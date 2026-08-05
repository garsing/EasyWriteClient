<template>
  <div class="pay-mask" @click.self="onClose">
    <div class="pay-dialog" role="dialog" aria-labelledby="pay-dialog-title">
      <header class="pay-header">
        <h2 id="pay-dialog-title" class="pay-title">{{ titleText }}</h2>
        <button type="button" class="pay-close" aria-label="关闭" @click="onClose">×</button>
      </header>

      <p class="pay-amount">¥{{ amountYuan }}</p>
      <p class="pay-desc">{{ description }}</p>

      <!-- 微信：二维码；支付宝：弹层打开时自动跳转浏览器 -->
      <div v-if="isWechat" class="qr-wrap">
        <img v-if="qrDataUrl" class="qr-img" :src="qrDataUrl" alt="支付二维码" />
        <p v-else-if="qrError" class="pay-error">{{ qrError }}</p>
        <p v-else class="pay-hint">生成二维码中…</p>
      </div>

      <p v-if="isMock" class="mock-note">{{ mockNote }}</p>
      <p v-if="statusTip" class="status-tip" :class="{ 'status-tip--err': isError }">{{ statusTip }}</p>
      <p v-if="remainText" class="remain">{{ remainText }}</p>

      <div class="pay-actions">
        <button type="button" class="btn-secondary" @click="onClose">取消</button>
        <button
          type="button"
          class="btn-primary"
          :class="{ 'btn-primary--alipay': !isWechat }"
          :disabled="confirming || paid"
          @click="onConfirmPaid"
        >
          {{ confirming ? '查询中…' : '我已支付' }}
        </button>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, onBeforeUnmount, watch } from 'vue'
import QRCode from 'qrcode'
import { getInspirationOrder, confirmInspirationPaid } from '../services/inspirationPaymentApi'
import { useWebViewBridge } from '../composables/useWebViewBridge'

const props = defineProps({
  orderId: { type: Number, required: true },
  channel: { type: String, default: 'alipay_page' },
  codeUrl: { type: String, default: '' },
  checkoutUrl: { type: String, default: '' },
  payFormHtml: { type: String, default: '' },
  amountYuan: { type: [Number, String], default: 0 },
  description: { type: String, default: '' },
  expireAt: { type: String, default: '' },
  isMock: { type: Boolean, default: false }
})

const emit = defineEmits(['close', 'paid'])
const { sendMessage, isWebView2 } = useWebViewBridge()

const qrDataUrl = ref('')
const qrError = ref('')
const statusTip = ref('')
const isError = ref(false)
const confirming = ref(false)
const paid = ref(false)
const remainText = ref('')

let pollTimer = null
let remainTimer = null

const isWechat = computed(() => (props.channel || '') === 'wechat_native')
const titleText = computed(() => (isWechat.value ? '微信支付' : '支付宝支付'))
const mockNote = computed(() =>
  isWechat.value
    ? '开发模式：未接微信时，点「我已支付」即可入账'
    : '开发模式：未接支付宝时，点「我已支付」即可入账'
)

const expireMs = computed(() => {
  if (!props.expireAt) return 0
  const t = Date.parse(props.expireAt.replace(' ', 'T'))
  return Number.isFinite(t) ? t : 0
})

async function renderQr () {
  if (!isWechat.value) return
  qrError.value = ''
  if (!props.codeUrl) {
    qrError.value = '缺少支付链接'
    return
  }
  try {
    qrDataUrl.value = await QRCode.toDataURL(props.codeUrl, {
      width: 220,
      margin: 2,
      errorCorrectionLevel: 'M'
    })
  } catch (e) {
    console.error(e)
    qrError.value = '二维码生成失败'
  }
}

async function openAlipay () {
  const url = props.checkoutUrl
  if (url && isWebView2) {
    try {
      const res = await sendMessage('openExternalUrl', { url })
      if (res && res.success === false) {
        window.open(url, '_blank')
      }
      statusTip.value = '已在浏览器打开支付宝，完成后点「我已支付」'
      return
    } catch {
      /* fallthrough */
    }
  }
  if (url) {
    window.open(url, '_blank')
    statusTip.value = '已在浏览器打开支付宝，完成后点「我已支付」'
    return
  }
  if (props.payFormHtml) {
    const w = window.open('', '_blank')
    if (w) {
      w.document.write(props.payFormHtml)
      w.document.close()
      statusTip.value = '已在浏览器打开支付宝，完成后点「我已支付」'
    } else {
      isError.value = true
      statusTip.value = '无法打开支付页，请允许弹窗后重试'
    }
  }
}

function clearTimers () {
  if (pollTimer) {
    clearInterval(pollTimer)
    pollTimer = null
  }
  if (remainTimer) {
    clearInterval(remainTimer)
    remainTimer = null
  }
}

function updateRemain () {
  if (!expireMs.value) {
    remainText.value = ''
    return
  }
  const left = expireMs.value - Date.now()
  if (left <= 0) {
    remainText.value = '订单已超时'
    isError.value = true
    statusTip.value = '支付超时，请重新下单'
    clearTimers()
    return
  }
  const sec = Math.ceil(left / 1000)
  const m = Math.floor(sec / 60)
  const s = sec % 60
  remainText.value = `请在 ${m}:${String(s).padStart(2, '0')} 内完成支付`
}

async function pollOnce () {
  if (paid.value) return
  try {
    const data = await getInspirationOrder(props.orderId)
    handleStatus(data)
  } catch (e) {
    console.warn('[PaymentQrDialog] poll', e)
  }
}

function handleStatus (data) {
  if (!data) return
  if (data.status === 'paid') {
    paid.value = true
    isError.value = false
    statusTip.value = '支付成功'
    clearTimers()
    emit('paid', data)
    return
  }
  if (data.status === 'expired' || data.status === 'closed') {
    isError.value = true
    statusTip.value = data.status === 'expired' ? '订单已超时' : '订单已关闭'
    clearTimers()
  }
}

async function onConfirmPaid () {
  if (confirming.value || paid.value) return
  confirming.value = true
  try {
    const data = await confirmInspirationPaid(props.orderId)
    handleStatus(data)
    if (data.status !== 'paid') {
      statusTip.value = '尚未检测到支付，请稍后再试'
    }
  } catch (e) {
    isError.value = true
    statusTip.value = e.message || '查询失败'
  } finally {
    confirming.value = false
  }
}

function onClose () {
  emit('close')
}

onMounted(async () => {
  if (isWechat.value) {
    statusTip.value = '请使用微信扫码支付'
    await renderQr()
  } else if (props.checkoutUrl || props.payFormHtml) {
    await openAlipay()
  } else {
    isError.value = true
    statusTip.value = '缺少支付宝支付链接'
  }
  updateRemain()
  remainTimer = setInterval(updateRemain, 1000)
  pollTimer = setInterval(pollOnce, 2000)
})

onBeforeUnmount(() => {
  clearTimers()
})

watch(() => props.codeUrl, () => {
  renderQr()
})
</script>

<style scoped>
.pay-mask {
  position: fixed;
  inset: 0;
  z-index: 1000;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(0, 0, 0, 0.35);
}

.pay-dialog {
  width: min(360px, calc(100vw - 32px));
  padding: 20px 20px 16px;
  background: #f7f7f5;
  border-radius: 12px;
  box-sizing: border-box;
}

.pay-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 8px;
}

.pay-title {
  margin: 0;
  font-size: 16px;
  font-weight: 600;
  color: #222;
}

.pay-close {
  border: none;
  background: transparent;
  font-size: 22px;
  line-height: 1;
  color: #666;
  cursor: pointer;
}

.pay-amount {
  margin: 0 0 4px;
  font-size: 28px;
  font-weight: 700;
  color: #1a1a1a;
  text-align: center;
}

.pay-desc {
  margin: 0 0 12px;
  font-size: 13px;
  color: #666;
  text-align: center;
}

.qr-wrap {
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 220px;
  margin-bottom: 10px;
}

.qr-img {
  width: 220px;
  height: 220px;
  display: block;
  background: #fff;
}

.pay-hint,
.pay-error {
  margin: 0;
  font-size: 13px;
  color: #666;
}

.pay-error {
  color: #b33;
}

.mock-note {
  margin: 0 0 8px;
  font-size: 12px;
  color: #888;
  text-align: center;
}

.status-tip {
  margin: 0 0 4px;
  font-size: 13px;
  color: #444;
  text-align: center;
}

.status-tip--err {
  color: #b33;
}

.remain {
  margin: 0 0 12px;
  font-size: 12px;
  color: #888;
  text-align: center;
}

.pay-actions {
  display: flex;
  gap: 10px;
}

.btn-secondary,
.btn-primary {
  flex: 1;
  padding: 10px 12px;
  border: none;
  border-radius: 8px;
  font-size: 14px;
  cursor: pointer;
}

.btn-secondary {
  background: #e4e4e2;
  color: #333;
}

.btn-primary {
  background: #07c160;
  color: #fff;
  font-weight: 500;
}

.btn-primary--alipay {
  background: #1677ff;
}

.btn-primary:disabled {
  opacity: 0.55;
  cursor: not-allowed;
}
</style>

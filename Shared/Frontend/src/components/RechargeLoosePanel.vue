<template>
  <div class="sub-page">
    <header class="sub-page-header">
      <button type="button" class="back-btn" aria-label="返回设置" @click="$emit('back')">
        <svg class="back-icon" viewBox="0 0 24 24" aria-hidden="true">
          <path
            d="M15 5L8 12l7 7M8 12h13"
            fill="none"
            stroke="currentColor"
            stroke-width="1.5"
            stroke-linecap="round"
            stroke-linejoin="round"
          />
        </svg>
      </button>
      <h1 class="sub-page-title">充值散装灵感值</h1>
    </header>

    <div class="sub-page-body">
      <p class="rate-line">1 元 = {{ LOOSE_RECHARGE_PER_CNY }} 灵感值</p>
      <p class="hint-line">入账后将先抵消负的散装余额。金额 {{ LOOSE_MIN_YUAN }}～{{ LOOSE_MAX_YUAN }} 整元。</p>

      <p class="field-label">选择金额</p>
      <div class="preset-row">
        <button
          v-for="preset in LOOSE_PRESET_YUAN"
          :key="preset"
          type="button"
          class="preset-chip"
          :class="{ 'preset-chip--active': Number(amountYuan) === preset }"
          @click="amountYuan = preset"
        >
          {{ preset }} 元
        </button>
      </div>

      <label class="field-label" for="custom-recharge-amount">自定义金额（元）</label>
      <input
        id="custom-recharge-amount"
        v-model.number="amountYuan"
        type="number"
        class="amount-input"
        :min="LOOSE_MIN_YUAN"
        :max="LOOSE_MAX_YUAN"
        step="1"
        placeholder="请输入金额"
      />

      <p class="preview-line">
        可得灵感值 <strong>{{ previewInspiration }}</strong>
      </p>

      <PayChannelPicker v-model="payChannel" />

      <p v-if="errorTip" class="mock-tip">{{ errorTip }}</p>

      <button
        type="button"
        class="confirm-btn"
        :disabled="!canConfirm || ordering"
        @click="onConfirm"
      >
        {{ ordering ? '下单中…' : '确认充值' }}
      </button>
    </div>

    <PaymentQrDialog
      v-if="payOrder"
      :order-id="payOrder.order_id"
      :channel="payOrder.channel || 'alipay_page'"
      :code-url="payOrder.code_url || ''"
      :checkout-url="payOrder.checkout_url || ''"
      :pay-form-html="payOrder.pay_form_html || ''"
      :amount-yuan="payOrder.amount_yuan"
      description="散装灵感值充值"
      :expire-at="payOrder.expire_at || ''"
      :is-mock="!!payOrder.mock"
      @close="payOrder = null"
      @paid="onPaid"
    />
  </div>
</template>

<script setup>
import { ref, computed } from 'vue'
import {
  LOOSE_RECHARGE_PER_CNY,
  LOOSE_PRESET_YUAN,
  LOOSE_MIN_YUAN,
  LOOSE_MAX_YUAN,
  looseInspirationFromYuan,
  isValidRechargeYuan
} from '../constants/inspirationCommerce'
import { createInspirationOrder } from '../services/inspirationPaymentApi'
import PaymentQrDialog from './PaymentQrDialog.vue'
import PayChannelPicker from './PayChannelPicker.vue'

const emit = defineEmits(['back', 'paid'])

const amountYuan = ref(50)
const payChannel = ref('alipay_page')
const errorTip = ref('')
const ordering = ref(false)
const payOrder = ref(null)

const previewInspiration = computed(() => looseInspirationFromYuan(amountYuan.value))
const canConfirm = computed(() => isValidRechargeYuan(amountYuan.value))

async function onConfirm () {
  if (!canConfirm.value || ordering.value) return
  errorTip.value = ''
  ordering.value = true
  try {
    const yuan = Math.floor(Number(amountYuan.value))
    const data = await createInspirationOrder({
      productType: 'loose',
      amountYuan: yuan,
      channel: payChannel.value
    })
    payOrder.value = data
  } catch (e) {
    errorTip.value = e.message || '下单失败'
  } finally {
    ordering.value = false
  }
}

function onPaid () {
  payOrder.value = null
  emit('paid')
}
</script>

<style scoped>
.sub-page {
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 0;
  box-sizing: border-box;
  padding: 16px 20px 20px;
  background: #ececec;
}

.sub-page-header {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-shrink: 0;
  margin-bottom: 12px;
}

.back-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 32px;
  height: 32px;
  padding: 0;
  border: none;
  border-radius: 0;
  background: transparent;
  color: #555;
  cursor: pointer;
}

.back-btn:hover {
  color: #222;
  background: transparent;
}

.back-icon {
  width: 22px;
  height: 22px;
  display: block;
}

.sub-page-title {
  margin: 0;
  font-size: 18px;
  font-weight: 600;
  color: #333;
}

.sub-page-body {
  flex: 1;
  min-height: 0;
  overflow-y: auto;
  width: 100%;
  max-width: 720px;
}

.rate-line {
  margin: 0 0 4px;
  font-size: 15px;
  font-weight: 500;
  color: #222;
}

.hint-line {
  margin: 0 0 16px;
  font-size: 13px;
  color: #666;
  line-height: 1.45;
}

.field-label {
  display: block;
  margin: 0 0 8px;
  font-size: 13px;
  color: #555;
}

.preset-row {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin-bottom: 14px;
}

.preset-chip {
  padding: 8px 14px;
  border: 1px solid #d0d0ce;
  border-radius: 8px;
  background: #f7f7f5;
  font-size: 13px;
  color: #333;
  cursor: pointer;
}

.preset-chip--active {
  border-color: #1a4d8f;
  background: #e8eef6;
  color: #1a4d8f;
}

.amount-input {
  width: 100%;
  box-sizing: border-box;
  margin-bottom: 12px;
  padding: 10px 12px;
  border: 1px solid #d0d0ce;
  border-radius: 8px;
  font-size: 15px;
  background: #fff;
}

.preview-line {
  margin: 0 0 12px;
  font-size: 14px;
  color: #333;
}

.mock-tip {
  margin: 0 0 12px;
  padding: 8px 12px;
  border-radius: 6px;
  background: #f7f7f5;
  color: #666;
  font-size: 13px;
  text-align: center;
}

.confirm-btn {
  width: 100%;
  padding: 12px;
  border: none;
  border-radius: 8px;
  background: #1a4d8f;
  color: #fff;
  font-size: 15px;
  font-weight: 500;
  cursor: pointer;
}

.confirm-btn:disabled {
  background: #c8c8c6;
  color: #666;
  cursor: not-allowed;
}
</style>

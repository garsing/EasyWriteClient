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
      <h1 class="sub-page-title">升级套餐</h1>
    </header>

    <PayChannelPicker v-model="payChannel" />

    <p v-if="errorTip" class="mock-tip">{{ errorTip }}</p>

    <div class="plans-scroll">
      <div class="plans-row">
        <article
          v-for="plan in PLAN_CATALOG"
          :key="plan.code"
          class="plan-card"
          :class="{ 'plan-card--current': isCurrent(plan) }"
        >
          <div class="plan-card-top">
            <h2 class="plan-name">{{ plan.name }}</h2>
            <span v-if="isCurrent(plan)" class="current-badge">当前套餐</span>
          </div>
          <p class="plan-price">
            ¥{{ plan.priceYuan }}<span class="plan-price-unit">/月</span>
          </p>
          <p class="plan-blurb">{{ plan.blurb }}</p>
          <ul class="plan-features">
            <li>充值汇率 1 元 = {{ plan.rechargePerCny }} 灵感值</li>
            <li>月限额 {{ monthlyQuota(plan) }} 灵感值</li>
          </ul>
          <button
            type="button"
            class="plan-cta"
            :disabled="ctaDisabled(plan) || ordering"
            @click="onUpgrade(plan)"
          >
            {{ ctaLabel(plan) }}
          </button>
        </article>
      </div>
    </div>

    <PaymentQrDialog
      v-if="payOrder"
      :order-id="payOrder.order_id"
      :channel="payOrder.channel || 'alipay_page'"
      :code-url="payOrder.code_url || ''"
      :checkout-url="payOrder.checkout_url || ''"
      :pay-form-html="payOrder.pay_form_html || ''"
      :amount-yuan="payOrder.amount_yuan"
      :description="payDescription"
      :expire-at="payOrder.expire_at || ''"
      :is-mock="!!payOrder.mock"
      @close="payOrder = null"
      @paid="onPaid"
    />
  </div>
</template>

<script setup>
import { ref } from 'vue'
import { PLAN_CATALOG, monthlyQuota, planTier } from '../constants/inspirationCommerce'
import { createInspirationOrder } from '../services/inspirationPaymentApi'
import PaymentQrDialog from './PaymentQrDialog.vue'
import PayChannelPicker from './PayChannelPicker.vue'

const props = defineProps({
  currentPlanCode: {
    type: String,
    default: ''
  },
  currentPlanName: {
    type: String,
    default: ''
  },
  planExpiresAt: {
    type: String,
    default: ''
  }
})

const emit = defineEmits(['back', 'paid'])

const errorTip = ref('')
const payChannel = ref('alipay_page')
const ordering = ref(false)
const payOrder = ref(null)
const payDescription = ref('')

function resolvedCurrentCode () {
  const code = (props.currentPlanCode || '').trim().toLowerCase()
  if (code && PLAN_CATALOG.some((p) => p.code === code)) return code
  const name = (props.currentPlanName || '').trim()
  const hit = PLAN_CATALOG.find((p) => p.name === name)
  return hit ? hit.code : ''
}

function isPaidActive () {
  if (!props.planExpiresAt) return false
  const t = Date.parse(String(props.planExpiresAt).replace(' ', 'T'))
  return Number.isFinite(t) && t > Date.now()
}

function isCurrent (plan) {
  const cur = resolvedCurrentCode()
  return !!cur && cur === plan.code && isPaidActive()
}

function ctaLabel (plan) {
  const cur = resolvedCurrentCode()
  if (!cur || !isPaidActive()) return '开通'
  if (cur === plan.code) return '续费'
  if (planTier(plan.code) > planTier(cur)) return '升级'
  if (planTier(plan.code) < planTier(cur)) return '不可降级'
  return '开通'
}

function ctaDisabled (plan) {
  const cur = resolvedCurrentCode()
  if (!cur || !isPaidActive()) return false
  return planTier(plan.code) < planTier(cur)
}

async function onUpgrade (plan) {
  if (ctaDisabled(plan) || ordering.value) return
  errorTip.value = ''
  ordering.value = true
  try {
    const data = await createInspirationOrder({
      productType: 'plan',
      planCode: plan.code,
      channel: payChannel.value
    })
    payDescription.value = `${plan.name} 月套餐`
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

.mock-tip {
  flex-shrink: 0;
  margin: 0 0 12px;
  padding: 8px 12px;
  border-radius: 6px;
  background: #f7f7f5;
  color: #666;
  font-size: 13px;
  text-align: center;
}

.plans-scroll {
  flex: 1;
  min-height: 0;
  overflow-x: auto;
  overflow-y: auto;
}

.plans-row {
  display: flex;
  gap: 12px;
  min-width: min-content;
  height: 100%;
  align-items: stretch;
}

.plan-card {
  display: flex;
  flex-direction: column;
  flex: 1 1 220px;
  min-width: 200px;
  max-width: 280px;
  padding: 18px 16px 16px;
  background: #f7f7f5;
  border-radius: 12px;
  border: 1px solid #e4e4e2;
  box-sizing: border-box;
}

.plan-card--current {
  background: #ececec;
  border-color: #d0d0ce;
}

.plan-card-top {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  margin-bottom: 8px;
}

.plan-name {
  margin: 0;
  font-size: 18px;
  font-weight: 600;
  color: #222;
}

.current-badge {
  flex-shrink: 0;
  font-size: 11px;
  color: #666;
  padding: 2px 8px;
  border-radius: 4px;
  background: rgba(0, 0, 0, 0.06);
}

.plan-price {
  margin: 0 0 8px;
  font-size: 28px;
  font-weight: 700;
  color: #1a1a1a;
  line-height: 1.2;
}

.plan-price-unit {
  margin-left: 2px;
  font-size: 14px;
  font-weight: 500;
  color: #666;
}

.plan-blurb {
  margin: 0 0 12px;
  font-size: 13px;
  line-height: 1.45;
  color: #666;
  min-height: 2.9em;
}

.plan-features {
  margin: 0 0 16px;
  padding: 0 0 0 18px;
  flex: 1;
  font-size: 13px;
  line-height: 1.55;
  color: #444;
}

.plan-features li {
  margin-bottom: 4px;
}

.plan-cta {
  width: 100%;
  margin-top: auto;
  padding: 10px 12px;
  border: none;
  border-radius: 8px;
  background: #1a4d8f;
  color: #fff;
  font-size: 14px;
  font-weight: 500;
  cursor: pointer;
}

.plan-cta:hover:not(:disabled) {
  background: #163f75;
}

.plan-cta:disabled {
  background: #c8c8c6;
  color: #666;
  cursor: not-allowed;
}
</style>

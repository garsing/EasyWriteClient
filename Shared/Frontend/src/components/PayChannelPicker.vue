<template>
  <div class="pay-channel">
    <p class="pay-channel-label">支付方式</p>
    <div class="pay-channel-list" role="radiogroup" aria-label="支付方式">
      <button
        v-for="opt in options"
        :key="opt.value"
        type="button"
        class="pay-channel-card"
        :class="{ 'pay-channel-card--active': modelValue === opt.value }"
        role="radio"
        :aria-checked="modelValue === opt.value"
        @click="$emit('update:modelValue', opt.value)"
      >
        <img
          class="pay-channel-logo"
          :src="opt.icon"
          :alt="opt.name"
          width="28"
          height="28"
        />
        <span class="pay-channel-text">
          <span class="pay-channel-name">{{ opt.name }}</span>
          <span v-if="opt.sub" class="pay-channel-sub">{{ opt.sub }}</span>
        </span>
      </button>
    </div>
  </div>
</template>

<script setup>
import iconAlipay from '../assets/pay/alipay.png'
import iconWechat from '../assets/pay/wechat.png'

defineProps({
  modelValue: {
    type: String,
    default: 'alipay_page'
  }
})

defineEmits(['update:modelValue'])

const options = [
  { value: 'alipay_page', name: '支付宝', sub: 'ALIPAY', icon: iconAlipay },
  { value: 'wechat_native', name: '微信支付', sub: '', icon: iconWechat }
]
</script>

<style scoped>
.pay-channel {
  margin-bottom: 14px;
}

.pay-channel-label {
  margin: 0 0 10px;
  font-size: 13px;
  color: #8a8a8a;
}

.pay-channel-list {
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.pay-channel-card {
  display: flex;
  align-items: center;
  gap: 12px;
  width: 100%;
  height: 46px;
  box-sizing: border-box;
  padding: 0 16px;
  border: 1px solid #e8e8e8;
  border-radius: 8px;
  background: #fff;
  text-align: left;
  cursor: pointer;
  transition: border-color 0.15s ease;
}

.pay-channel-card--active {
  border: 2px solid #111;
  padding: 0 15px;
}

.pay-channel-logo {
  flex-shrink: 0;
  width: 28px;
  height: 28px;
  object-fit: contain;
  border-radius: 6px;
}

.pay-channel-text {
  display: flex;
  flex-direction: column;
  justify-content: center;
  gap: 0;
  min-width: 0;
}

.pay-channel-name {
  font-size: 15px;
  font-weight: 500;
  color: #222;
  line-height: 1.15;
}

.pay-channel-card:not(.pay-channel-card--active) .pay-channel-name {
  color: #555;
}

.pay-channel-sub {
  font-size: 10px;
  font-weight: 500;
  letter-spacing: 0.05em;
  color: #a0a0a0;
  line-height: 1.1;
}
</style>

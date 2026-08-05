/** 散装充值：1 元 = 多少灵感值（充值换算，非消费系数） */
export const LOOSE_RECHARGE_PER_CNY = 100

export const LOOSE_PRESET_YUAN = [10, 30, 50, 100, 200]

export const LOOSE_MIN_YUAN = 1
export const LOOSE_MAX_YUAN = 5000

/** 付费套餐目录（月限额 = priceYuan * rechargePerCny） */
export const PLAN_CATALOG = [
  {
    code: 'pro',
    name: 'Pro',
    priceYuan: 60,
    rechargePerCny: 120,
    blurb: '升级后按自然月限额，适合日常写作。'
  },
  {
    code: 'pro_plus',
    name: 'Pro+',
    priceYuan: 100,
    rechargePerCny: 150,
    blurb: '更高月限额与更优充值汇率。'
  },
  {
    code: 'ultra',
    name: 'Ultra',
    priceYuan: 200,
    rechargePerCny: 180,
    blurb: '最高月限额与充值汇率。'
  }
]

const TIER = { pro: 1, pro_plus: 2, ultra: 3 }

export function planTier (code) {
  return TIER[code] || 0
}

export function monthlyQuota (plan) {
  return plan.priceYuan * plan.rechargePerCny
}

export function looseInspirationFromYuan (yuan) {
  const n = Number(yuan)
  if (!Number.isFinite(n) || n <= 0) return 0
  return Math.floor(n) * LOOSE_RECHARGE_PER_CNY
}

export function isValidRechargeYuan (yuan) {
  const n = Number(yuan)
  if (!Number.isFinite(n)) return false
  if (!Number.isInteger(n) && Math.floor(n) !== n) return false
  const i = Math.floor(n)
  return i >= LOOSE_MIN_YUAN && i <= LOOSE_MAX_YUAN
}

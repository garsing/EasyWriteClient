<template>
  <div class="settings-page">
    <!-- 关闭放在前端，避免被 WebView2 盖住 WinForms 按钮 -->
    <button
      type="button"
      class="settings-close-btn"
      title="关闭"
      aria-label="关闭"
      @click="handleClose"
    >
      ×
    </button>

    <!-- 升级 / 充值子页：整窗，隐藏侧栏 -->
    <UpgradePlansPanel
      v-if="viewMode === 'upgrade'"
      :current-plan-code="planCode"
      :current-plan-name="planName"
      :plan-expires-at="planExpiresAt"
      @back="backToMain"
      @paid="onPaymentPaid"
    />
    <RechargeLoosePanel
      v-else-if="viewMode === 'recharge'"
      @back="backToMain"
      @paid="onPaymentPaid"
    />

    <template v-else>
      <header class="settings-title-bar">
        <h1 class="settings-title">设置</h1>
      </header>

      <div class="settings-layout">
        <aside class="settings-sidebar">
          <nav class="sidebar-nav">
            <button
              v-for="item in navItems"
              :key="item.id"
              type="button"
              class="sidebar-item"
              :class="{ 'sidebar-item--active': activeSection === item.id }"
              @click="activeSection = item.id"
            >
              {{ item.label }}
            </button>
          </nav>
        </aside>

        <main class="settings-main">
          <div class="settings-card">
            <div v-if="loading" class="settings-body settings-loading">加载中…</div>

            <div v-else-if="!isLoggedIn" class="settings-body settings-login">
              <p class="login-tip">请先登录后使用设置</p>
              <button type="button" class="btn-login" :disabled="loginOpening" @click="handleLogin">
                {{ loginOpening ? '打开中…' : '登录' }}
              </button>
            </div>

            <template v-else>
              <!-- 计划 -->
              <template v-if="activeSection === 'overview'">
                <header class="settings-header">
                  <span class="username">{{ displayUsername }}</span>
                </header>
                <div class="settings-divider" />
                <div class="settings-body settings-body--stack">
                  <p v-if="quotaError" class="save-error">{{ quotaError }}</p>
                  <template v-else>
                    <p class="usage-label usage-label--row">
                      <span class="usage-label-text">
                        <span v-if="planName" class="plan-name">{{ planName }}</span>
                        灵感值 {{ displayUsedInspiration }}/{{ displayQuotaInspiration }}
                        <span v-if="expiresLabel" class="expires-inline">（{{ expiresLabel }}）</span>
                      </span>
                      <button type="button" class="btn-link" @click="viewMode = 'upgrade'">升级</button>
                    </p>
                    <div class="usage-bar" aria-hidden="true">
                      <div class="usage-bar-fill" :style="{ width: barWidthPercent + '%' }" />
                    </div>
                    <p class="usage-label usage-label--row">
                      <span class="usage-label-text usage-label-text--loose">
                        散装灵感值 {{ displayLooseInspiration }}
                        <span
                          class="loose-help"
                          tabindex="0"
                          role="img"
                          aria-label="今日额度用完会消耗散装灵感值"
                        >
                          ?
                          <span class="loose-help-tip" role="tooltip">今日额度用完会消耗散装灵感值</span>
                        </span>
                      </span>
                      <button type="button" class="btn-link" @click="viewMode = 'recharge'">充值</button>
                    </p>
                  </template>
                </div>
              </template>

              <!-- 通用：仅工作目录 -->
              <template v-else-if="activeSection === 'general'">
                <div class="settings-body settings-body--top">
                  <label class="workspace-label" for="workspace-root-input">工作目录</label>
                  <input
                    id="workspace-root-input"
                    v-model="workspaceRoot"
                    type="text"
                    class="workspace-input"
                    :disabled="saving"
                    spellcheck="false"
                    @blur="handleSaveWorkspace"
                  />
                  <p v-if="remapHint" class="remap-hint">{{ remapHint }}</p>
                  <p v-if="saveError" class="save-error">{{ saveError }}</p>
                </div>
              </template>

              <!-- 关于我们：本期留空 -->
              <template v-else-if="activeSection === 'about'">
                <div class="settings-body settings-body--stack settings-body--top settings-about-empty" />
              </template>

              <footer v-if="activeSection === 'overview'" class="settings-footer">
                <button
                  type="button"
                  class="btn-logout"
                  :disabled="loggingOut"
                  @click="handleLogout"
                >
                  {{ loggingOut ? '退出中…' : '退出登录' }}
                </button>
              </footer>
            </template>
          </div>
        </main>
      </div>
    </template>
  </div>
</template>

<script setup>
import { ref, computed, onMounted, watch } from 'vue'
import { useWebViewBridge } from '../composables/useWebViewBridge'
import UpgradePlansPanel from './UpgradePlansPanel.vue'
import RechargeLoosePanel from './RechargeLoosePanel.vue'

const { sendMessage, isWebView2 } = useWebViewBridge()

const navItems = [
  { id: 'overview', label: '计划' },
  { id: 'general', label: '通用' },
  { id: 'about', label: '关于我们' }
]

/** 'main' | 'upgrade' | 'recharge' */
const viewMode = ref('main')
const activeSection = ref('overview')
const loading = ref(true)
const isLoggedIn = ref(false)
const username = ref('')
const workspaceRoot = ref('')
const remapHint = ref('')
const saving = ref(false)
const saveError = ref('')
const loggingOut = ref(false)
const loginOpening = ref(false)

const planName = ref('')
const planCode = ref('')
const planExpiresAt = ref('')
const dailyUsedInspiration = ref(0)
const dailyQuotaInspiration = ref(0)
const looseInspiration = ref(0)
const quotaError = ref('')
const quotaLoading = ref(false)

const displayUsername = computed(() => {
  if (isLoggedIn.value && username.value) {
    return username.value
  }
  return '未登录'
})

const displayUsedInspiration = computed(() => {
  const n = Number(dailyUsedInspiration.value)
  if (!Number.isFinite(n) || n <= 0) return 0
  return Math.floor(n)
})

const displayQuotaInspiration = computed(() => {
  const n = Number(dailyQuotaInspiration.value)
  if (!Number.isFinite(n) || n <= 0) return 0
  return Math.floor(n)
})

const displayLooseInspiration = computed(() => {
  const n = Number(looseInspiration.value)
  if (!Number.isFinite(n)) return 0
  return Math.trunc(n)
})

const expiresLabel = computed(() => {
  if (!planExpiresAt.value) return ''
  const t = Date.parse(String(planExpiresAt.value).replace(' ', 'T'))
  if (!Number.isFinite(t) || t <= Date.now()) return ''
  const d = new Date(t)
  const m = d.getMonth() + 1
  const day = d.getDate()
  return `到期 ${m}/${day}`
})

const barWidthPercent = computed(() => {
  const used = Number(dailyUsedInspiration.value)
  const quota = Number(dailyQuotaInspiration.value)
  if (!Number.isFinite(used) || !Number.isFinite(quota) || quota <= 0 || used <= 0) return 0
  return Math.min(100, (used / quota) * 100)
})

async function loadLlmQuota () {
  quotaError.value = ''
  if (!isLoggedIn.value) return

  quotaLoading.value = true
  try {
    if (!isWebView2) {
      planName.value = '基础免费'
      planCode.value = 'basic_free'
      planExpiresAt.value = ''
      dailyUsedInspiration.value = 0
      dailyQuotaInspiration.value = 300
      looseInspiration.value = 0
      return
    }
    const res = await sendMessage('getLlmQuota', {})
    if (!res || !res.success) {
      quotaError.value = res?.message || '加载额度失败'
      planName.value = ''
      planCode.value = ''
      planExpiresAt.value = ''
      dailyUsedInspiration.value = 0
      dailyQuotaInspiration.value = 0
      looseInspiration.value = 0
      return
    }
    planName.value = res.planName || ''
    planCode.value = res.planCode || ''
    planExpiresAt.value = res.planExpiresAt || ''
    dailyUsedInspiration.value = Number(res.dailyUsedInspiration) || 0
    dailyQuotaInspiration.value = Number(res.dailyQuotaInspiration) || 0
    looseInspiration.value = Number(res.looseInspiration) || 0
  } catch (err) {
    console.error('[UserSettingsPanel] 加载额度失败:', err)
    quotaError.value = '加载额度失败'
  } finally {
    quotaLoading.value = false
  }
}

async function loadSettings () {
  loading.value = true
  saveError.value = ''
  try {
    if (!isWebView2) {
      isLoggedIn.value = true
      username.value = 'dev-user'
      workspaceRoot.value = 'D:/YiWrite/WorkSpace'
      remapHint.value = ''
      await loadLlmQuota()
      return
    }

    const res = await sendMessage('getUserSettings', {})
    if (!res || !res.success) {
      saveError.value = res?.message || '加载设置失败'
      return
    }

    isLoggedIn.value = !!res.isLoggedIn
    username.value = res.username || ''
    workspaceRoot.value = res.workspaceRoot || res.defaultRoot || ''
    remapHint.value = res.remapHint || ''
    if (isLoggedIn.value) {
      await loadLlmQuota()
    }
  } catch (err) {
    console.error('[UserSettingsPanel] 加载设置失败:', err)
    saveError.value = '加载设置失败'
  } finally {
    loading.value = false
  }
}

function backToMain () {
  viewMode.value = 'main'
  if (isLoggedIn.value) {
    loadLlmQuota()
  }
}

function onPaymentPaid () {
  viewMode.value = 'main'
  if (isLoggedIn.value) {
    loadLlmQuota()
  }
}

watch(activeSection, (id) => {
  if (id === 'overview' && isLoggedIn.value) {
    loadLlmQuota()
  }
})

async function handleSaveWorkspace () {
  if (!isLoggedIn.value || saving.value) return

  const path = workspaceRoot.value.trim()
  if (!path) return

  saving.value = true
  saveError.value = ''
  try {
    const res = await sendMessage('saveWorkspaceRoot', { path })
    if (!res || !res.success) {
      saveError.value = res?.message || '保存工作目录失败'
      await loadSettings()
      return
    }
    workspaceRoot.value = res.workspaceRoot || path
    remapHint.value = res.remapHint || ''
  } catch (err) {
    console.error('[UserSettingsPanel] 保存工作目录失败:', err)
    saveError.value = '保存工作目录失败'
  } finally {
    saving.value = false
  }
}

async function handleLogout () {
  if (loggingOut.value) return

  loggingOut.value = true
  try {
    const res = await sendMessage('logoutUser', {})
    if (!res || !res.success) {
      saveError.value = res?.message || '退出登录失败'
      return
    }
    await handleClose()
  } catch (err) {
    console.error('[UserSettingsPanel] 退出登录失败:', err)
    saveError.value = '退出登录失败'
  } finally {
    loggingOut.value = false
  }
}

async function handleLogin () {
  if (loginOpening.value) return

  loginOpening.value = true
  try {
    await sendMessage('openLoginWindow', {})
    await loadSettings()
  } catch (err) {
    console.error('[UserSettingsPanel] 打开登录窗口失败:', err)
  } finally {
    loginOpening.value = false
  }
}

async function handleClose () {
  try {
    if (isWebView2) {
      await sendMessage('closeSettingsWindow', {})
      return
    }
    window.close()
  } catch (err) {
    console.error('[UserSettingsPanel] 关闭窗口失败:', err)
  }
}

onMounted(() => {
  loadSettings()
})
</script>

<style scoped>
.settings-page {
  position: relative;
  display: flex;
  flex-direction: column;
  width: 100%;
  height: 100%;
  box-sizing: border-box;
  background: #ececec;
  border: 1px solid #e0e0e0;
  border-radius: 14px;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Microsoft YaHei', sans-serif;
  overflow: hidden;
}

.settings-close-btn {
  position: absolute;
  top: 10px;
  right: 12px;
  z-index: 20;
  width: 36px;
  height: 36px;
  margin: 0;
  padding: 0;
  border: none;
  border-radius: 8px;
  background: transparent;
  color: #666;
  font-size: 22px;
  line-height: 1;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
}

.settings-close-btn:hover {
  background: rgba(0, 0, 0, 0.06);
  color: #333;
}

.settings-title-bar {
  flex-shrink: 0;
  padding: 20px 56px 12px 24px;
}

.settings-title {
  margin: 0;
  font-size: 18px;
  font-weight: 600;
  color: #333;
}

.settings-layout {
  display: flex;
  flex: 1;
  min-height: 0;
  padding: 0 20px 20px;
  gap: 8px;
}

.settings-sidebar {
  flex-shrink: 0;
  width: 168px;
  padding-top: 4px;
}

.sidebar-nav {
  display: flex;
  flex-direction: column;
  gap: 4px;
}

.sidebar-item {
  width: 100%;
  padding: 10px 16px;
  border: none;
  border-radius: 6px;
  background: transparent;
  color: #333;
  font-size: 14px;
  text-align: left;
  cursor: pointer;
  transition: background-color 0.15s ease;
}

.sidebar-item:hover {
  background: rgba(0, 0, 0, 0.04);
}

.sidebar-item--active {
  background: #d6d6d6;
  font-weight: 500;
}

.settings-main {
  flex: 1;
  min-width: 0;
  min-height: 0;
}

.settings-card {
  display: flex;
  flex-direction: column;
  height: 100%;
  background: #f7f7f5;
  border-radius: 12px;
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.06);
  overflow: hidden;
}

.settings-header {
  padding: 22px 32px 16px;
}

.username {
  font-size: 16px;
  font-weight: 600;
  color: #1a4d8f;
}

.expires-inline {
  margin-left: 4px;
  font-weight: 400;
  color: #888;
  font-size: 12px;
}

.plan-name {
  margin-right: 8px;
  font-size: inherit;
  color: #555;
}

.settings-divider {
  height: 1px;
  margin: 0 32px;
  background: #e8e8e8;
}

.settings-body {
  flex: 1;
  padding: 36px 32px 24px;
  display: grid;
  grid-template-columns: auto 1fr;
  grid-template-rows: auto auto auto;
  column-gap: 24px;
  row-gap: 10px;
  align-items: center;
  align-content: start;
}

.settings-body--top {
  flex: 0 0 auto;
  padding-top: 22px;
}

.settings-body--stack {
  display: flex;
  flex-direction: column;
  align-items: stretch;
  gap: 12px;
}

.settings-about-empty {
  min-height: 120px;
}

.settings-loading,
.settings-login {
  display: block;
  text-align: center;
  color: #666;
  font-size: 14px;
}

.settings-login {
  padding-top: 80px;
}

.login-tip {
  margin: 0 0 16px;
}

.btn-login {
  padding: 8px 24px;
  font-size: 14px;
  color: #fff;
  background: #1a4d8f;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}

.btn-login:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.usage-label {
  margin: 0;
  font-size: 14px;
  color: #333;
}

.usage-label--row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}

.usage-label-text {
  min-width: 0;
}

.usage-label-text--loose {
  display: inline-flex;
  align-items: center;
  gap: 6px;
}

.btn-link {
  flex-shrink: 0;
  padding: 2px 10px;
  border: 1px solid #1a4d8f;
  border-radius: 4px;
  background: transparent;
  color: #1a4d8f;
  font-size: 13px;
  line-height: 1.4;
  cursor: pointer;
}

.btn-link:hover {
  background: rgba(26, 77, 143, 0.06);
}

.loose-help {
  position: relative;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 12px;
  height: 12px;
  border: 1px solid #999;
  border-radius: 50%;
  font-size: 9px;
  line-height: 1;
  color: #666;
  cursor: help;
  flex-shrink: 0;
  user-select: none;
}

.loose-help-tip {
  position: absolute;
  left: calc(100% + 8px);
  top: 50%;
  transform: translateY(-50%);
  z-index: 2;
  width: max-content;
  max-width: 220px;
  padding: 6px 10px;
  border-radius: 4px;
  background: #333;
  color: #fff;
  font-size: 12px;
  line-height: 1.4;
  white-space: normal;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.18);
  opacity: 0;
  visibility: hidden;
  pointer-events: none;
  transition: opacity 0.12s ease;
}

.loose-help:hover .loose-help-tip,
.loose-help:focus .loose-help-tip {
  opacity: 1;
  visibility: visible;
}

.usage-bar {
  width: 100%;
  height: 10px;
  border-radius: 5px;
  background: #ececec;
  overflow: hidden;
}

.usage-bar-fill {
  height: 100%;
  border-radius: 5px;
  background: #1a4d8f;
  transition: width 0.2s ease;
}

.workspace-label {
  font-size: 14px;
  color: #333;
  white-space: nowrap;
}

.workspace-input {
  width: 100%;
  min-width: 0;
  height: 34px;
  padding: 4px 12px;
  font-size: 14px;
  color: #333;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  outline: none;
  box-sizing: border-box;
}

.workspace-input:focus {
  border-color: #b0b0b0;
}

.workspace-input:disabled {
  background: #f8f8f8;
}

.remap-hint,
.save-error {
  grid-column: 1 / -1;
  margin: 4px 0 0;
  font-size: 12px;
  line-height: 1.5;
}

.remap-hint {
  color: #888;
}

.save-error {
  color: #c0392b;
}

.settings-footer {
  display: flex;
  justify-content: flex-end;
  margin-top: auto;
  padding: 0 32px 28px;
}

.btn-logout {
  padding: 8px 20px;
  font-size: 14px;
  color: #fff;
  background: #e74c3c;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}

.btn-logout:hover:not(:disabled) {
  background: #d63c2c;
}

.btn-logout:disabled {
  opacity: 0.6;
  cursor: not-allowed;
}
</style>

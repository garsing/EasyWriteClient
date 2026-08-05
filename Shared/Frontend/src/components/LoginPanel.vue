<template>
  <div class="login-page">
    <!-- 提示固定在窗口最顶部（登录 / 注册共用） -->
    <div
      v-if="(viewMode === 'login' && (loginNotice || loginError)) || (viewMode === 'register' && (registerBanner || registerError))"
      class="login-banners"
    >
      <template v-if="viewMode === 'login'">
        <p v-if="loginNotice" class="form-banner form-banner-success">{{ loginNotice }}</p>
        <p v-if="loginError" class="form-banner form-banner-error">{{ loginError }}</p>
      </template>
      <template v-else>
        <p v-if="registerBanner" class="form-banner form-banner-success">{{ registerBanner }}</p>
        <p v-if="registerError" class="form-banner form-banner-error">{{ registerError }}</p>
      </template>
    </div>

    <div class="login-body" :class="viewMode === 'login' ? 'login-body--login' : 'login-body--register'">
      <!-- 登录视图 -->
      <template v-if="viewMode === 'login'">
        <div class="login-stack">
          <div class="login-tabs">
            <button
              type="button"
              class="login-tab"
              :class="{ active: loginMethod === 'password' }"
              @click="loginMethod = 'password'"
            >
              密码登录
            </button>
            <button
              type="button"
              class="login-tab"
              :class="{ active: loginMethod === 'sms' }"
              @click="loginMethod = 'sms'"
            >
              验证码登录
            </button>
          </div>

          <div class="field field-user">
            <div class="input-wrap">
              <input
                v-model="credential"
                type="text"
                class="text-input"
                placeholder="手机号"
                autocomplete="tel"
                spellcheck="false"
                @keyup.enter="handleLogin"
              />
              <button
                v-if="credential"
                type="button"
                class="input-action"
                tabindex="-1"
                aria-label="清除"
                @mousedown.prevent
                @click="credential = ''"
              >
                ×
              </button>
            </div>
          </div>

          <div v-if="loginMethod === 'password'" class="field field-pass login-pivot">
            <div class="input-wrap">
              <input
                v-model="password"
                :type="showLoginPassword ? 'text' : 'password'"
                class="text-input"
                placeholder="密码"
                autocomplete="current-password"
                @keyup.enter="handleLogin"
              />
              <button
                type="button"
                class="input-action input-action-eye"
                tabindex="-1"
                :aria-label="showLoginPassword ? '隐藏密码' : '显示密码'"
                @mousedown.prevent
                @click="showLoginPassword = !showLoginPassword"
              >
                <span class="eye-icon" :class="{ open: showLoginPassword }" />
              </button>
            </div>
          </div>

          <div v-else class="field field-sms login-pivot">
            <div class="input-wrap input-wrap--sms">
              <input
                v-model="loginSmsCode"
                type="text"
                class="text-input"
                placeholder="验证码"
                maxlength="4"
                autocomplete="one-time-code"
                @keyup.enter="handleLogin"
              />
              <button
                type="button"
                class="btn-sms"
                :disabled="smsCooldownLeft > 0 || sendingSms"
                @click="handleSendSms('login')"
              >
                {{ smsCountdownLabel }}
              </button>
            </div>
          </div>

          <label class="agreement-row agreement-row--login">
            <input v-model="agreedToTerms" type="checkbox" class="agreement-check" />
            <span class="agreement-text">
              我已阅读并同意
              <button type="button" class="agreement-link" @click.prevent="openAgreement('user')">《用户协议》</button>
              和
              <button type="button" class="agreement-link" @click.prevent="openAgreement('privacy')">《隐私政策》</button>
            </span>
          </label>

          <button
            type="button"
            class="btn-primary btn-login"
            :disabled="loggingIn"
            @click="handleLogin"
          >
            {{ loggingIn ? '登录中…' : '登录' }}
          </button>
        </div>

        <div class="login-footer">
          <button type="button" class="link-btn" @click="switchToRegister">立即注册</button>
        </div>
      </template>

      <!-- 注册视图 -->
      <template v-else>
        <div class="register-stack">
          <div class="register-center">
            <h1 class="login-title login-title--register">注册账号</h1>

            <div class="field field-hint-outside">
              <div class="input-wrap">
                <input
                  v-model="regPhone"
                  type="text"
                  class="text-input"
                  placeholder="手机号"
                  autocomplete="tel"
                  spellcheck="false"
                  @input="refreshRegHints"
                />
                <p
                  :class="hintClass(regPhoneHint)"
                  :title="regPhoneHint.text || undefined"
                >
                  {{ regPhoneHint.text }}
                </p>
              </div>
            </div>

            <div class="field field-hint-outside">
              <div class="input-wrap input-wrap--sms">
                <input
                  v-model="regSmsCode"
                  type="text"
                  class="text-input"
                  placeholder="验证码"
                  maxlength="4"
                  autocomplete="one-time-code"
                  @input="refreshRegHints"
                />
                <button
                  type="button"
                  class="btn-sms"
                  :disabled="smsCountdownLeft > 0 || sendingSms"
                  @click="handleSendSms('register')"
                >
                  {{ smsCountdownLabel }}
                </button>
                <p
                  :class="hintClass(regSmsHint)"
                  :title="regSmsHint.text || undefined"
                >
                  {{ regSmsHint.text }}
                </p>
              </div>
            </div>

            <div class="field field-hint-outside">
              <div class="input-wrap input-wrap--has-action">
                <input
                  v-model="regPassword"
                  :type="showRegPassword ? 'text' : 'password'"
                  class="text-input"
                  placeholder="密码"
                  autocomplete="new-password"
                  @input="refreshRegHints"
                />
                <p
                  :class="hintClass(regPasswordHint)"
                  :title="regPasswordHint.text || undefined"
                >
                  {{ regPasswordHint.text }}
                </p>
                <button
                  type="button"
                  class="input-action input-action-eye"
                  tabindex="-1"
                  :aria-label="showRegPassword ? '隐藏密码' : '显示密码'"
                  @mousedown.prevent
                  @click="showRegPassword = !showRegPassword"
                >
                  <span class="eye-icon" :class="{ open: showRegPassword }" />
                </button>
              </div>
            </div>

            <div class="field field-hint-outside field-hint-outside-last">
              <div class="input-wrap input-wrap--has-action">
                <input
                  v-model="regConfirm"
                  :type="showRegConfirm ? 'text' : 'password'"
                  class="text-input"
                  placeholder="确认密码"
                  autocomplete="new-password"
                  @input="refreshRegHints"
                />
                <p
                  :class="hintClass(regConfirmHint)"
                  :title="regConfirmHint.text || undefined"
                >
                  {{ regConfirmHint.text }}
                </p>
                <button
                  type="button"
                  class="input-action input-action-eye"
                  tabindex="-1"
                  :aria-label="showRegConfirm ? '隐藏密码' : '显示密码'"
                  @mousedown.prevent
                  @click="showRegConfirm = !showRegConfirm"
                >
                  <span class="eye-icon" :class="{ open: showRegConfirm }" />
                </button>
              </div>
            </div>

            <label class="agreement-row">
              <input v-model="agreedToTerms" type="checkbox" class="agreement-check" />
              <span class="agreement-text">
                我已阅读并同意
                <button type="button" class="agreement-link" @click.prevent="openAgreement('user')">《用户协议》</button>
                和
                <button type="button" class="agreement-link" @click.prevent="openAgreement('privacy')">《隐私政策》</button>
              </span>
            </label>

            <button
              type="button"
              class="btn-primary btn-register"
              :disabled="!canSubmitRegister || registering"
              @click="handleRegister"
            >
              {{ registering ? '注册中…' : '注册' }}
            </button>

            <div class="register-footer">
              <button type="button" class="link-btn" @click="switchToLogin">已有账号？去登录</button>
            </div>
          </div>
        </div>
      </template>
    </div>

    <!-- 与登录窗同风格的协议确认框 -->
    <div
      v-if="agreementDialogVisible"
      class="agree-dialog-mask"
      @click.self="resolveAgreementDialog(false)"
    >
      <div class="agree-dialog" role="dialog" aria-modal="true" aria-labelledby="agree-dialog-title">
        <h2 id="agree-dialog-title" class="agree-dialog-title">用户协议与隐私政策</h2>
        <p class="agree-dialog-body">
          继续操作前需同意
          <button type="button" class="agreement-link" @click="openAgreement('user')">《用户协议》</button>
          和
          <button type="button" class="agreement-link" @click="openAgreement('privacy')">《隐私政策》</button>。
          点击「确定」即表示您已阅读并同意上述协议。
        </p>
        <div class="agree-dialog-actions">
          <button type="button" class="btn-dialog-cancel" @click="resolveAgreementDialog(false)">取消</button>
          <button type="button" class="btn-dialog-ok" @click="resolveAgreementDialog(true)">确定</button>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, onUnmounted } from 'vue'
import { useWebViewBridge } from '../composables/useWebViewBridge'
import {
  validatePhone,
  validateSmsCode,
  validatePassword,
  validateConfirmPassword
} from '../utils/registerValidation'

const { sendMessage, isWebView2 } = useWebViewBridge()

const viewMode = ref('login')
const loginMethod = ref('password')

const credential = ref('')
const password = ref('')
const loginSmsCode = ref('')
const showLoginPassword = ref(false)
const loginError = ref('')
const loginNotice = ref('')
const loggingIn = ref(false)
const agreedToTerms = ref(false)
const agreementDialogVisible = ref(false)
let agreementDialogResolver = null

const regPhone = ref('')
const regSmsCode = ref('')
const regPassword = ref('')
const regConfirm = ref('')
const showRegPassword = ref(false)
const showRegConfirm = ref(false)
const registering = ref(false)
const registerError = ref('')
const registerBanner = ref('')

const sendingSms = ref(false)
const smsCountdownLeft = ref(0)
let smsTimer = null

const emptyHint = () => ({ kind: 'none', text: '' })

const regPhoneHint = ref(emptyHint())
const regSmsHint = ref(emptyHint())
const regPasswordHint = ref(emptyHint())
const regConfirmHint = ref(emptyHint())

const smsCountdownLabel = computed(() => {
  if (sendingSms.value) return '发送中…'
  if (smsCountdownLeft.value > 0) return `${smsCountdownLeft.value}s`
  return '获取验证码'
})

function hintFromValidation (result, hasInput) {
  if (result.valid) {
    return hasInput ? { kind: 'ok', text: '格式正确' } : emptyHint()
  }
  return hasInput ? { kind: 'error', text: result.message } : emptyHint()
}

function hintClass (hint) {
  if (!hint.text) {
    return 'field-hint field-hint-inline field-hint-empty'
  }
  return hint.kind === 'ok'
    ? 'field-hint field-hint-inline field-hint-ok'
    : 'field-hint field-hint-inline field-hint-error'
}

function refreshRegHints () {
  registerError.value = ''
  regPhoneHint.value = hintFromValidation(validatePhone(regPhone.value), !!regPhone.value)
  regSmsHint.value = hintFromValidation(validateSmsCode(regSmsCode.value), !!regSmsCode.value)
  regPasswordHint.value = hintFromValidation(
    validatePassword(regPassword.value),
    !!regPassword.value
  )
  regConfirmHint.value = hintFromValidation(
    validateConfirmPassword(regPassword.value, regConfirm.value),
    !!regConfirm.value
  )
}

const canSubmitRegister = computed(() =>
  validatePhone(regPhone.value).valid &&
  validateSmsCode(regSmsCode.value).valid &&
  validatePassword(regPassword.value).valid &&
  validateConfirmPassword(regPassword.value, regConfirm.value).valid
)

async function openAgreement (kind) {
  if (!isWebView2) {
    console.log('[LoginPanel] openAgreement', kind)
    return
  }
  try {
    await sendMessage('openAgreement', { kind })
  } catch (err) {
    console.error('[LoginPanel] 打开协议失败:', err)
  }
}

function resolveAgreementDialog (ok) {
  agreementDialogVisible.value = false
  const resolve = agreementDialogResolver
  agreementDialogResolver = null
  if (ok) agreedToTerms.value = true
  if (resolve) resolve(ok)
}

/** 未勾选协议时弹同风格确认框；确定则勾选并返回 true */
function ensureAgreedToTerms () {
  if (agreedToTerms.value) return Promise.resolve(true)
  if (agreementDialogVisible.value) {
    return new Promise((resolve) => {
      agreementDialogResolver = resolve
    })
  }
  agreementDialogVisible.value = true
  return new Promise((resolve) => {
    agreementDialogResolver = resolve
  })
}

function startSmsCountdown () {
  smsCountdownLeft.value = 60
  if (smsTimer) clearInterval(smsTimer)
  smsTimer = setInterval(() => {
    if (smsCountdownLeft.value <= 1) {
      smsCountdownLeft.value = 0
      clearInterval(smsTimer)
      smsTimer = null
      return
    }
    smsCountdownLeft.value -= 1
  }, 1000)
}

onUnmounted(() => {
  if (smsTimer) clearInterval(smsTimer)
})

function switchToRegister () {
  viewMode.value = 'register'
  loginError.value = ''
  loginNotice.value = ''
  registerError.value = ''
  registerBanner.value = ''
  refreshRegHints()
}

function switchToLogin () {
  viewMode.value = 'login'
  registerError.value = ''
  registerBanner.value = ''
}

async function handleSendSms (scene) {
  if (sendingSms.value || smsCountdownLeft.value > 0) return

  const phone = scene === 'register' ? regPhone.value.trim() : credential.value.trim()
  const phoneCheck = validatePhone(phone)
  if (!phoneCheck.valid) {
    if (scene === 'register') {
      registerError.value = phoneCheck.message
      refreshRegHints()
    } else {
      loginError.value = phoneCheck.message
    }
    return
  }

  sendingSms.value = true
  loginError.value = ''
  registerError.value = ''
  try {
    if (!isWebView2) {
      console.log('[LoginPanel] dev mock sendSms', { phone, scene })
      startSmsCountdown()
      return
    }
    const res = await sendMessage('sendSmsCode', { phone, scene })
    if (!res || !res.success) {
      const msg = res?.message || '验证码发送失败'
      if (scene === 'register') registerError.value = msg
      else loginError.value = msg
      return
    }
    startSmsCountdown()
  } catch (err) {
    console.error('[LoginPanel] 发码失败:', err)
    const msg = '验证码发送失败，请稍后重试'
    if (scene === 'register') registerError.value = msg
    else loginError.value = msg
  } finally {
    sendingSms.value = false
  }
}

async function handleLogin () {
  if (loggingIn.value) return

  const cred = credential.value.trim()
  if (!cred) {
    loginError.value = '请输入手机号'
    return
  }

  if (!(await ensureAgreedToTerms())) return

  loggingIn.value = true
  loginError.value = ''
  try {
    if (loginMethod.value === 'sms') {
      const phoneCheck = validatePhone(cred)
      if (!phoneCheck.valid) {
        loginError.value = phoneCheck.message
        return
      }
      const codeCheck = validateSmsCode(loginSmsCode.value)
      if (!codeCheck.valid) {
        loginError.value = codeCheck.message
        return
      }
      if (!isWebView2) {
        console.log('[LoginPanel] dev mock sms login')
        return
      }
      const res = await sendMessage('loginBySms', {
        phone: cred,
        sms_code: loginSmsCode.value.trim()
      })
      if (!res || !res.success) {
        loginError.value = res?.message || '验证码错误或已过期'
      }
      return
    }

    if (!password.value) {
      loginError.value = '请输入密码'
      return
    }
    if (!isWebView2) {
      console.log('[LoginPanel] dev mock login')
      return
    }
    const res = await sendMessage('loginUser', { credential: cred, password: password.value })
    if (!res || !res.success) {
      loginError.value = res?.message || '手机号或密码错误'
    }
  } catch (err) {
    console.error('[LoginPanel] 登录失败:', err)
    loginError.value = '登录失败，请稍后重试'
  } finally {
    loggingIn.value = false
  }
}

async function handleRegister () {
  if (registering.value || !canSubmitRegister.value) return

  if (!(await ensureAgreedToTerms())) return

  registering.value = true
  registerError.value = ''
  registerBanner.value = ''
  try {
    const phone = regPhone.value.trim()
    const smsCode = regSmsCode.value.trim()
    const pwd = regPassword.value
    if (!isWebView2) {
      console.log('[LoginPanel] dev mock register + auto login')
      return
    }
    const res = await sendMessage('registerUser', {
      phone,
      sms_code: smsCode,
      password: pwd
    })
    if (!res || !res.success) {
      registerError.value = res?.message || '注册失败'
      return
    }
    // 成功时由 Bridge 自动登录并关闭窗口
  } catch (err) {
    console.error('[LoginPanel] 注册失败:', err)
    registerError.value = '注册失败，请稍后重试'
  } finally {
    registering.value = false
  }
}
</script>

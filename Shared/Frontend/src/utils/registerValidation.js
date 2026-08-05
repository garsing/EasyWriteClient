const CJK_RE = /[\u4e00-\u9fff]/

/** @returns {{ valid: boolean, message: string | null }} */
export function validateUsername (raw) {
  const username = (raw ?? '').trim()
  if (!username) return { valid: false, message: '用户名不能为空' }
  if (username.length < 3 || username.length > 50) {
    return { valid: false, message: '用户名长度应在3-50字符之间' }
  }
  if (username.includes('@')) return { valid: false, message: '用户名不能包含@符号' }
  if (CJK_RE.test(username)) return { valid: false, message: '用户名不能包含中文' }
  return { valid: true, message: null }
}

/** @returns {{ valid: boolean, message: string | null }} */
export function validatePhone (raw) {
  const phone = (raw ?? '').trim()
  if (!phone) return { valid: false, message: '手机号不能为空' }
  if (!/^1[3-9]\d{9}$/.test(phone)) return { valid: false, message: '手机号格式不正确' }
  return { valid: true, message: null }
}

/** @returns {{ valid: boolean, message: string | null }} */
export function validateSmsCode (raw) {
  const code = (raw ?? '').trim()
  if (!code) return { valid: false, message: '请输入验证码' }
  if (!/^\d{4}$/.test(code)) return { valid: false, message: '请输入4位验证码' }
  return { valid: true, message: null }
}

/** @returns {{ valid: boolean, message: string | null }} */
export function validatePassword (raw) {
  const password = raw ?? ''
  if (!password.trim()) return { valid: false, message: '密码不能为空' }
  if (password.length < 6) return { valid: false, message: '密码长度至少6位' }
  return { valid: true, message: null }
}

/** @returns {{ valid: boolean, message: string | null }} */
export function validateConfirmPassword (password, confirm) {
  if (!confirm) return { valid: false, message: '请再次输入密码' }
  if (confirm !== password) return { valid: false, message: '两次输入的密码不一致' }
  return { valid: true, message: null }
}

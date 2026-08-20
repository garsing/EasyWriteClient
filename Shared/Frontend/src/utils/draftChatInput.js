/** 输入框草稿（未发送文本），切历史 / 切回「新对话」时读写。 */

let current = ''
let stashed = ''

export function syncLiveDraftInput (value) {
  current = value == null ? '' : String(value)
}

export function getLiveDraftInput () {
  return current
}

/** 离开新对话去看历史时调用 */
export function stashDraftInputFromLive () {
  stashed = current
  return stashed
}

export function peekStashedDraftInput () {
  return stashed
}

export function clearStashedDraftInput () {
  stashed = ''
}

export function clearAllDraftInput () {
  current = ''
  stashed = ''
}

export const DRAFT_TASK_ID = '__draft_new__'

export function makeDraftTaskItem () {
  return {
    id: DRAFT_TASK_ID,
    title: '新对话',
    isDraft: true,
    last_message_at: null,
    updated_at: null
  }
}

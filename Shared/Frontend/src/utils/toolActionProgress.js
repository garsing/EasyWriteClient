const ACTION_LIST_TOOLS = new Set(['F_process_document_actions'])

/**
 * 批量文字操作折叠标题右侧的 current/total。
 * 只信宿主推送的进度；单条不显示。历史会话没有进度则不画。
 */
export function formatActionProgressLabel({
  toolName,
  progressCurrent,
  progressTotal
}) {
  if (!ACTION_LIST_TOOLS.has(toolName)) {
    return ''
  }

  const total = Number(progressTotal)
  const current = Number(progressCurrent)
  if (!(total > 1) || !(current > 0)) {
    return ''
  }

  let cur = current
  if (cur < 1) cur = 1
  if (cur > total) cur = total
  return `${cur}/${total}`
}

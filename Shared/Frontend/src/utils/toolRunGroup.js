import { isHiddenTool } from './toolVisibility.js'

/**
 * 把连续、中间没有可见正文的工具收成 toolRun。
 * 可见工具卡不足两张时不包外层。B_todo 丢掉。
 * 第一张工具前的思考留在组外；工具之间以及最后一张工具之后的思考留在组内。
 *
 * @param {Array} segments 已去掉 todoList 与纯空白正文
 * @returns {Array}
 */
export function groupToolRuns (segments) {
  const output = []
  let lead = []
  let run = []
  let tools = 0

  function emitLead () {
    if (lead.length) output.push(...lead)
    lead = []
  }

  function flush () {
    if (tools >= 2) {
      output.push(makeRun(run))
    } else if (tools === 1) {
      output.push(...run)
    } else {
      emitLead()
      if (run.length) output.push(...run)
    }
    lead = []
    run = []
    tools = 0
  }

  for (const seg of segments || []) {
    if (!seg) continue
    if (seg.type === 'toolCall' && isHiddenTool(seg.toolName)) continue
    if (seg.type === 'text') {
      flush()
      output.push(seg)
      continue
    }
    if (seg.type === 'thinking') {
      if (tools === 0) lead.push(seg)
      else run.push(seg)
      continue
    }
    if (seg.type === 'toolCall') {
      if (tools === 0) emitLead()
      run.push(seg)
      tools += 1
      continue
    }
    flush()
    output.push(seg)
  }
  flush()
  return output
}

function makeRun (run) {
  const first = run.find((s) => s && s.type === 'toolCall')
  return {
    type: 'toolRun',
    startIndex: first?.startIndex,
    toolCallId: first?.toolCallId,
    segments: run.slice()
  }
}

/**
 * 折叠条标题：从后往前找当前这张已知名的工具。
 * 尾部还没补上 function.name 时沿用上一张。
 */
export function currentToolName (segments) {
  const segs = segments || []
  for (let i = segs.length - 1; i >= 0; i--) {
    if (segs[i]?.type !== 'toolCall') continue
    const name = String(segs[i].toolName || '').trim()
    if (name) return name
  }
  return ''
}

/**
 * 外层折叠标题。末段仍是进行中的思考时显示「思考」，否则显示当前工具。
 */
export function currentRunTitle (segments) {
  const segs = segments || []
  const last = segs.length ? segs[segs.length - 1] : null
  if (last && last.type === 'thinking' && !last.isComplete) {
    return { thinking: true, toolName: '' }
  }
  return { thinking: false, toolName: currentToolName(segs) }
}

export function toolRunKey (run, index) {
  if (run?.startIndex != null && run.startIndex !== '') return `run-${run.startIndex}`
  if (run?.toolCallId) return `run-${run.toolCallId}`
  return `run-${index}`
}

import { categoryOf } from './toolRunCategory.js'
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
 * 组标题。尾段是思考 → thinking。
 * 尾部工具还没有名字 → 往前找上一张已知名，中间的思考不抢标题。
 * 一张已知名都没有 → unknown。
 * @returns {'thinking'|'explore'|'operate'|'unknown'}
 */
export function resolveRunHeading (segments) {
  const segs = segments || []
  for (let i = segs.length - 1; i >= 0; i--) {
    const seg = segs[i]
    if (!seg) continue
    if (seg.type === 'thinking') return 'thinking'
    if (seg.type !== 'toolCall') continue
    const cat = categoryOf(seg.toolName)
    if (cat) return cat
    for (let j = i - 1; j >= 0; j--) {
      if (segs[j]?.type !== 'toolCall') continue
      const prev = categoryOf(segs[j].toolName)
      if (prev) return prev
    }
    return 'unknown'
  }
  return 'unknown'
}

export function toolRunKey (run, index) {
  if (run?.startIndex != null && run.startIndex !== '') return `run-${run.startIndex}`
  if (run?.toolCallId) return `run-${run.toolCallId}`
  return `run-${index}`
}

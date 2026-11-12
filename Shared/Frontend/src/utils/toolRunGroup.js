import { isHiddenTool } from './toolVisibility.js'

/**
 * 两次思考中间夹着正文时：思考并成一段，中间的字挪到这段思考后面，
 * 和紧跟着的正文直接接上。工具隔开的思考不动。
 */
export function foldTextStuckBetweenThinking (segments) {
  const segs = (segments || []).filter(Boolean)
  const out = []
  let i = 0
  while (i < segs.length) {
    const seg = segs[i]
    if (seg.type !== 'thinking') {
      out.push({ ...seg })
      i += 1
      continue
    }
    const thinkings = [seg]
    const texts = []
    let j = i
    while (
      segs[j + 1] && segs[j + 1].type === 'text' &&
      segs[j + 2] && segs[j + 2].type === 'thinking'
    ) {
      texts.push(segs[j + 1])
      thinkings.push(segs[j + 2])
      j += 2
    }
    if (thinkings.length === 1) {
      out.push({ ...seg })
      i += 1
      continue
    }
    if (segs[j + 1] && segs[j + 1].type === 'text') {
      texts.push(segs[j + 1])
      j += 1
    }
    out.push(mergeThinkings(thinkings))
    const joined = texts.map((t) => t.content || '').join('')
    if (joined) {
      const lastText = texts[texts.length - 1]
      out.push({ ...lastText, type: 'text', content: joined })
    }
    i = j + 1
  }
  return out
}

function mergeThinkings (thinkings) {
  const first = thinkings[0]
  return {
    ...first,
    type: 'thinking',
    content: thinkings.map((t) => t.content || '').filter((s) => s).join('\n'),
    isComplete: thinkings.every((t) => t.isComplete !== false)
  }
}

/**
 * 把连续、中间没有可见正文的工具收成 toolRun。
 * 可见工具卡不足两张、且没有思考时不包外层。B_todo 丢掉。
 * 思考跟相邻工具走：工具前、工具间、正文后的思考都收进同一组，不单独成行。
 *
 * @param {Array} segments 已去掉 todoList 与纯空白正文
 * @returns {Array}
 */
export function groupToolRuns (segments) {
  const output = []
  let lead = []
  let run = []
  let tools = 0

  function flush () {
    if (tools >= 2 || (tools === 1 && run.some((s) => s && s.type === 'thinking'))) {
      output.push(makeRun(run))
    } else if (tools === 1) {
      output.push(...run)
    } else {
      if (lead.length) output.push(...lead)
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
      if (tools > 0) run.push(seg)
      else if (lead.length) lead.push(seg)
      else if (!attachThinking(output, seg)) lead.push(seg)
      continue
    }
    if (seg.type === 'toolCall') {
      if (tools === 0 && lead.length) {
        run.push(...lead)
        lead = []
      }
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

/** 正文后面才到的思考，收回前面那组工具，避免在答案下面再单列一行。 */
function attachThinking (output, seg) {
  for (let i = output.length - 1; i >= 0; i--) {
    const item = output[i]
    if (!item) continue
    if (item.type === 'text') continue
    if (item.type === 'toolRun') {
      item.segments.push(seg)
      return true
    }
    if (item.type === 'toolCall') {
      output[i] = makeRun([item, seg])
      return true
    }
    break
  }
  return false
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
 * 折叠条当前工具：从后往前找已知名的那张。
 * 尾部还没补上 function.name 时沿用上一张，进度也跟这张走。
 */
export function currentToolSegment (segments) {
  const segs = segments || []
  for (let i = segs.length - 1; i >= 0; i--) {
    if (segs[i]?.type !== 'toolCall') continue
    const name = String(segs[i].toolName || '').trim()
    if (name) return segs[i]
  }
  return null
}

export function currentToolName (segments) {
  const seg = currentToolSegment(segments)
  return seg ? String(seg.toolName || '').trim() : ''
}

/**
 * 外层折叠标题。末段仍是进行中的思考时显示「思考」，否则显示当前工具。
 * 当前工具带进度时带上 current/total，标题栏折叠时也能看见。
 */
export function currentRunTitle (segments) {
  const segs = segments || []
  const last = segs.length ? segs[segs.length - 1] : null
  if (last && last.type === 'thinking' && !last.isComplete) {
    return { thinking: true, toolName: '', progressCurrent: 0, progressTotal: 0 }
  }
  const seg = currentToolSegment(segs)
  return {
    thinking: false,
    toolName: seg ? String(seg.toolName || '').trim() : '',
    progressCurrent: seg?.progressCurrent || 0,
    progressTotal: seg?.progressTotal || 0
  }
}

export function toolRunKey (run, index) {
  if (run?.startIndex != null && run.startIndex !== '') return `run-${run.startIndex}`
  if (run?.toolCallId) return `run-${run.toolCallId}`
  return `run-${index}`
}

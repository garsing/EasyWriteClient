/**
 * toolCallAccumulator 单测（node toolCallAccumulator.test.mjs）
 */
import { ToolCallAccumulator } from './toolCallAccumulator.js'

function assert(cond, msg) {
  if (!cond) throw new Error(msg)
}

// 单轮：正文 + tool
const acc = new ToolCallAccumulator()
acc.feedDeltaToolCalls([
  { index: 0, id: 'call_1', function: { name: 'B_calculate', arguments: '{"expr' } }
])
acc.feedDeltaToolCalls([
  { index: 0, function: { arguments: 'ession":"1+1"}' } }
])
acc.feedFinishReason('tool_calls', '我来算一下')

const segments = acc.buildSegments('我来算一下')
assert(segments.length === 2, 'text + toolCall')
assert(segments[0].type === 'text', 'first is text')
assert(segments[1].type === 'toolCall', 'second is toolCall')
assert(segments[1].toolName === 'B_calculate', 'tool name')
assert(segments[1].isComplete === true, 'complete after finish_reason')

// 多轮 Orchestrator：正文(前) → tool → 正文(后)
const acc2 = new ToolCallAccumulator()
acc2.feedDeltaToolCalls([
  {
    index: 0,
    id: 'call_2',
    function: { name: 'B_calculate', arguments: '{"expression": "128 * 129"}' }
  }
])
acc2.feedFinishReason('tool_calls', '好的，我来计算这个表达式。')

const full = '好的，我来计算这个表达式。**128 × 129 = 16512** ✅'
const segs2 = acc2.buildSegments(full)
assert(segs2.length === 3, 'before + tool + after')
assert(segs2[0].type === 'text' && segs2[0].content === '好的，我来计算这个表达式。', 'before text')
assert(segs2[1].type === 'toolCall' && segs2[1].toolName === 'B_calculate', 'tool middle')
assert(segs2[2].type === 'text' && segs2[2].content === '**128 × 129 = 16512** ✅', 'after text')

// 多轮多工具：每轮 index 均为 0，应顺序追加卡片而非覆盖
const acc3 = new ToolCallAccumulator()
acc3.feedDeltaToolCalls([
  { index: 0, id: 'c1', function: { name: 'B_read_skill_file', arguments: '{"skill":"data-fetch"}' } }
])
acc3.feedFinishReason('tool_calls', '先看 skill。')

acc3.feedDeltaToolCalls([
  { index: 0, id: 'c2', function: { name: 'B_write_python', arguments: '{"path":"a.py"}' } }
])
acc3.feedFinishReason('tool_calls', '先看 skill。写脚本。')

acc3.feedDeltaToolCalls([
  { index: 0, id: 'c3', function: { name: 'B_data_fetcher', arguments: '{"query":"x"}' } }
])
acc3.feedFinishReason('tool_calls', '先看 skill。写脚本。拉数据。')

const segs3 = acc3.buildSegments('先看 skill。写脚本。拉数据。分析完成。')
assert(segs3.length === 7, 'multi-round: text/tool alternating + final')
assert(segs3[0].content === '先看 skill。', 'round1 text')
assert(segs3[1].toolName === 'B_read_skill_file', 'tool1')
assert(segs3[2].content === '写脚本。', 'round2 text')
assert(segs3[3].toolName === 'B_write_python', 'tool2')
assert(segs3[4].content === '拉数据。', 'round3 text')
assert(segs3[5].toolName === 'B_data_fetcher', 'tool3')
assert(segs3[6].content === '分析完成。', 'final text')

// 仅有工具名、arguments 仍空：不能当成 complete（否则写 .py 的卡片一开始就折叠）
const acc4 = new ToolCallAccumulator()
acc4.feedDeltaToolCalls([
  { index: 0, id: 'c4', function: { name: 'F_write_file', arguments: '' } }
])
const segs4 = acc4.buildSegments('')
assert(segs4.length === 1 && segs4[0].type === 'toolCall', 'in-progress write')
assert(segs4[0].isComplete === false, 'empty args is not complete')

acc4.feedDeltaToolCalls([
  { index: 0, function: { arguments: '{"path":"calc.py"}' } }
])
const segs5 = acc4.buildSegments('')
assert(segs5[0].isComplete === true, 'path-only JSON is syntactically complete')

console.log('[OK] toolCallAccumulator tests passed')

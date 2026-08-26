/**
 * 消费 SSE delta.tool_calls / reasoning，产出与 ToolCallBox / ThinkingBox 兼容的 segments。
 *
 * 多轮 Orchestrator：每轮 finish_reason=tool_calls 时提交「thinking + 正文 + 工具卡片」。
 * thinking 按时间线插入当前轮位置，不永久置顶整条气泡。
 */

function isValidJson(str) {
  if (!str || !str.trim()) return false
  try {
    JSON.parse(str)
    return true
  } catch {
    return false
  }
}

export class ToolCallAccumulator {
  constructor() {
    /** 当前轮次进行中的 tool_calls（按 SSE index 拼接） */
    this._partials = new Map()
    this._finishReason = null
    /** 已提交、按顺序固定的 segments */
    this._committedSegments = []
    /** 已写入 committed 的正文长度（对应 C# 累计 content 前缀） */
    this._committedContentLength = 0
    /** 单调递增，用于 ToolCallBox key */
    this._toolSeq = 0
    /** 单调递增，用于 ThinkingBox key */
    this._thinkingSeq = 0
    /** 当前未提交的 thinking 段（本轮进行中） */
    this._activeThinking = null
    /** 已消费的累计 reasoning 长度（对应 C# reasoningContent） */
    this._reasoningAppliedLength = 0
  }

  /**
   * @param {string|null} reason
   * @param {string|null} contentAtFinish 收到 tool_calls 完成时当前累计正文
   */
  feedFinishReason(reason, contentAtFinish = null) {
    if (reason === 'tool_calls' || reason === 'function_call') {
      this._commitRound(contentAtFinish ?? '')
      this._finishReason = reason
      return
    }
    if (reason) {
      this._finishReason = reason
    }
  }

  feedDeltaToolCalls(deltaToolCalls) {
    if (!Array.isArray(deltaToolCalls) || deltaToolCalls.length === 0) return

    for (const item of deltaToolCalls) {
      const index = Number(item.index ?? 0)
      if (!this._partials.has(index)) {
        this._partials.set(index, { index, id: null, name: null, argumentsParts: [] })
      }
      const partial = this._partials.get(index)
      if (item.id) partial.id = item.id
      const fn = item.function || {}
      if (fn.name) partial.name = fn.name
      if (fn.arguments) partial.argumentsParts.push(fn.arguments)
    }
  }

  /**
   * 从 C# 累计 reasoningContent 喂入增量；按时间线落段。
   * @param {string} fullReasoning
   * @param {string} contentSoFar 当前累计正文，用于判断本轮是否已有正文
   */
  feedReasoningFromCumulative(fullReasoning, contentSoFar = '') {
    const full = fullReasoning || ''
    if (full.length < this._reasoningAppliedLength) {
      this._reasoningAppliedLength = 0
    }
    const delta = full.slice(this._reasoningAppliedLength)
    if (!delta) return
    this._reasoningAppliedLength = full.length
    this._feedReasoningDelta(delta, contentSoFar || '')
  }

  _feedReasoningDelta(delta, contentSoFar) {
    const uncommittedText = contentSoFar.slice(this._committedContentLength)

    // 本轮已有正文后再来 thinking：先落下一段 thinking（若有）与正文，再开新 thinking
    if (uncommittedText) {
      this._flushActiveThinking({ complete: true })
      this._committedSegments.push({ type: 'text', content: uncommittedText })
      this._committedContentLength = contentSoFar.length
    }

    if (!this._activeThinking) {
      this._activeThinking = {
        type: 'thinking',
        content: '',
        isComplete: false,
        startIndex: this._thinkingSeq++
      }
    }
    this._activeThinking.content += delta
  }

  _flushActiveThinking({ complete }) {
    if (!this._activeThinking) return
    if (this._activeThinking.content) {
      this._committedSegments.push({
        type: 'thinking',
        content: this._activeThinking.content,
        isComplete: !!complete,
        startIndex: this._activeThinking.startIndex
      })
    }
    this._activeThinking = null
  }

  /** 流结束：进行中的 thinking 标完成 */
  markStreamComplete() {
    if (this._activeThinking) {
      this._activeThinking.isComplete = true
    }
    for (const seg of this._committedSegments) {
      if (seg.type === 'thinking') {
        seg.isComplete = true
      }
    }
  }

  _partialsToToolSegments({ forceComplete }) {
    const segments = []
    const indices = [...this._partials.keys()].sort((a, b) => a - b)

    for (const index of indices) {
      const partial = this._partials.get(index)
      const args = partial.argumentsParts.join('')
      const complete = forceComplete || isValidJson(args)
      segments.push({
        type: 'toolCall',
        toolName: partial.name || '',
        toolCallId: partial.id || '',
        content: args,
        isComplete: complete,
        startIndex: this._toolSeq++
      })
    }
    return segments
  }

  _commitRound(contentAtFinish) {
    // 工具轮结束：thinking → 正文 → 工具，保持时间顺序
    this._flushActiveThinking({ complete: true })

    const full = contentAtFinish || ''
    const newText = full.slice(this._committedContentLength)
    const toolSegs = this._partialsToToolSegments({ forceComplete: true })

    if (!newText && toolSegs.length === 0) {
      return
    }

    if (newText) {
      this._committedSegments.push({ type: 'text', content: newText })
    }
    this._committedSegments.push(...toolSegs)
    this._partials.clear()
    this._committedContentLength = full.length
    this._finishReason = null
  }

  /**
   * 合并正文 / thinking / tool_calls 为渲染 segments。
   * 已提交轮次 + 当前轮（thinking → 尾部正文 → 进行中 tool）。
   */
  buildSegments(content) {
    const full = content || ''
    const tailText = full.slice(this._committedContentLength)
    const inProgressTools = this._partialsToToolSegments({ forceComplete: false })

    const segments = [...this._committedSegments]

    if (this._activeThinking && this._activeThinking.content) {
      segments.push({ ...this._activeThinking })
    }

    if (inProgressTools.length > 0) {
      if (tailText) {
        segments.push({ type: 'text', content: tailText })
      }
      segments.push(...inProgressTools)
      return segments
    }

    if (tailText) {
      segments.push({ type: 'text', content: tailText })
    }

    if (segments.length === 0 && !full) {
      return []
    }
    return segments
  }
}

export function createToolCallAccumulator() {
  return new ToolCallAccumulator()
}

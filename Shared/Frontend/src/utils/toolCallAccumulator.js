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
    /**
     * 本轮已提交，但提交时还没有 function.name 的调用。
     * 名字后到时补进原卡，不再另开一行。
     */
    this._awaitingName = new Map()
    /** 刚提交的这一轮，index → 已进入 _committedSegments 的工具卡 */
    this._lastRoundTools = new Map()
    this._finishReason = null
    /** 已提交、按顺序固定的 segments */
    this._committedSegments = []
    /** 已写入 committed 的正文长度（对应 C# 累计 content 前缀） */
    this._committedContentLength = 0
    /** 单调递增，用于 ToolCallBox key。一张卡只分配一次 */
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
      this._feedOneToolDelta(item)
    }
  }

  _feedOneToolDelta(item) {
    const index = Number(item.index ?? 0)
    if (this._partials.has(index)) {
      this._applyDelta(this._partials.get(index), item)
      return
    }

    const awaiting = this._awaitingName.get(index)
    if (awaiting && !this._isDifferentCall(awaiting.id, item.id)) {
      this._applyDelta(awaiting, item)
      if (awaiting.name) {
        const seg = this._toSegment(awaiting, true)
        this._insertCommittedTool(seg)
        this._awaitingName.delete(index)
        this._lastRoundTools.set(index, seg)
      }
      return
    }

    const committed = this._lastRoundTools.get(index)
    if (
      committed &&
      !this._isDifferentCall(committed.toolCallId, item.id) &&
      !this._nameConflicts(committed.toolName, item)
    ) {
      this._patchSegment(committed, item)
      return
    }

    // 序号被下一轮复用（id 或名字对不上）：丢掉上一轮还没补上的空位
    if (
      (awaiting && this._isDifferentCall(awaiting.id, item.id)) ||
      (committed &&
        (this._isDifferentCall(committed.toolCallId, item.id) ||
          this._nameConflicts(committed.toolName, item)))
    ) {
      this._awaitingName.clear()
      this._lastRoundTools.clear()
    }

    const partial = {
      index,
      id: null,
      name: null,
      argumentsParts: [],
      startIndex: this._toolSeq++
    }
    this._partials.set(index, partial)
    this._applyDelta(partial, item)
  }

  /** 两边都有 id 且不同，才是另一次调用 */
  _isDifferentCall(existingId, incomingId) {
    return !!(existingId && incomingId && existingId !== incomingId)
  }

  _nameConflicts(existingName, item) {
    const incoming = item && item.function && item.function.name
    return !!(existingName && incoming && existingName !== incoming)
  }

  _applyDelta(partial, item) {
    if (item.id) partial.id = item.id
    const fn = item.function || {}
    if (fn.name) partial.name = fn.name
    if (fn.arguments) partial.argumentsParts.push(fn.arguments)
  }

  _patchSegment(seg, item) {
    if (item.id && !seg.toolCallId) seg.toolCallId = item.id
    const fn = item.function || {}
    if (fn.name && !seg.toolName) seg.toolName = fn.name
    if (fn.arguments) seg.content = (seg.content || '') + fn.arguments
    seg.isComplete = isValidJson(seg.content)
  }

  /** 按本轮 callIndex 插回工具卡序列，避免后到的名字排到下一轮正文后面 */
  _insertCommittedTool(seg) {
    let runStart = this._committedSegments.length
    while (runStart > 0 && this._committedSegments[runStart - 1].type === 'toolCall') {
      runStart--
    }
    let insertAt = this._committedSegments.length
    for (let i = runStart; i < this._committedSegments.length; i++) {
      const callIndex = this._committedSegments[i].callIndex
      if (callIndex !== undefined && callIndex !== null && callIndex > seg.callIndex) {
        insertAt = i
        break
      }
    }
    this._committedSegments.splice(insertAt, 0, seg)
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

  _toSegment(partial, forceComplete) {
    const args = partial.argumentsParts.join('')
    return {
      type: 'toolCall',
      toolName: partial.name || '',
      toolCallId: partial.id || '',
      content: args,
      outputText: '',
      outputDone: false,
      isComplete: !!forceComplete || isValidJson(args),
      startIndex: partial.startIndex,
      callIndex: partial.index
    }
  }

  _partialsToToolSegments({ forceComplete }) {
    const segments = []
    const indices = [...this._partials.keys()].sort((a, b) => a - b)

    for (const index of indices) {
      const partial = this._partials.get(index)
      // 名字还没到：先不画卡片，避免只剩一个折角
      if (!partial.name) continue
      segments.push(this._toSegment(partial, forceComplete))
    }
    return segments
  }

  _commitRound(contentAtFinish) {
    // 工具轮结束：thinking → 正文 → 工具，保持时间顺序
    this._flushActiveThinking({ complete: true })

    const full = contentAtFinish || ''
    const newText = full.slice(this._committedContentLength)
    const named = []
    const stillWaiting = new Map()
    const indices = [...this._partials.keys()].sort((a, b) => a - b)
    for (const index of indices) {
      const partial = this._partials.get(index)
      if (partial.name) named.push(this._toSegment(partial, true))
      else stillWaiting.set(index, partial)
    }

    // 收尾包会再送一次 finish_reason。没有新片段时保持上一轮的补名入口
    if (!newText && named.length === 0 && stillWaiting.size === 0) {
      return
    }

    this._partials.clear()
    this._awaitingName = stillWaiting
    this._lastRoundTools = new Map()
    for (const seg of named) {
      this._lastRoundTools.set(seg.callIndex, seg)
    }

    if (newText) {
      this._committedSegments.push({ type: 'text', content: newText })
    }
    this._committedSegments.push(...named)
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

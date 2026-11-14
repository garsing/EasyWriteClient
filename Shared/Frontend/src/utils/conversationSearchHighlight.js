const SKIP = 'pre, .work-fold, .tool-call-box, .tool-run-fold, .thinking-box, .hint-text, mark.conv-search-hit'

let marks = []

function clearMarks (root) {
  if (!root) {
    marks = []
    return
  }
  root.querySelectorAll('mark.conv-search-hit').forEach((mark) => {
    const parent = mark.parentNode
    if (!parent) return
    while (mark.firstChild) parent.insertBefore(mark.firstChild, mark)
    parent.removeChild(mark)
    parent.normalize()
  })
  marks = []
}

function searchableRoots (root) {
  const found = []
  root.querySelectorAll('.user-message .message-text, .assistant-reply-wrap .message-text').forEach((el) => {
    if (el.closest('.work-fold')) return
    found.push(el)
  })
  return found
}

function textNodesUnder (el) {
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {
    acceptNode (node) {
      if (!node.nodeValue) return NodeFilter.FILTER_REJECT
      const parent = node.parentElement
      if (!parent || parent.closest(SKIP)) return NodeFilter.FILTER_REJECT
      return NodeFilter.FILTER_ACCEPT
    }
  })
  const list = []
  let current = walker.nextNode()
  while (current) {
    list.push(current)
    current = walker.nextNode()
  }
  return list
}

function matchRanges (text, query) {
  const hay = text.toLowerCase()
  const needle = query.toLowerCase()
  const ranges = []
  let from = 0
  while (from <= hay.length - needle.length) {
    const at = hay.indexOf(needle, from)
    if (at < 0) break
    ranges.push([at, at + needle.length])
    from = at + needle.length
  }
  return ranges
}

function wrapTextNode (textNode, ranges) {
  const created = []
  let node = textNode
  for (let i = ranges.length - 1; i >= 0; i--) {
    const [start, end] = ranges[i]
    node.splitText(end)
    const mid = node.splitText(start)
    const mark = document.createElement('mark')
    mark.className = 'conv-search-hit'
    mid.parentNode.insertBefore(mark, mid)
    mark.appendChild(mid)
    created.push(mark)
  }
  created.reverse()
  return created
}

export function applyConversationSearch (root, query) {
  clearMarks(root)
  const q = String(query || '').trim()
  if (!root || !q) return 0
  searchableRoots(root).forEach((el) => {
    textNodesUnder(el).forEach((node) => {
      const ranges = matchRanges(node.nodeValue || '', q)
      if (!ranges.length) return
      marks.push(...wrapTextNode(node, ranges))
    })
  })
  return marks.length
}

export function setConversationSearchCurrent (index, scroll) {
  const live = marks.filter((mark) => mark.isConnected)
  if (live.length !== marks.length) marks = live
  marks.forEach((mark, i) => {
    mark.classList.toggle('conv-search-hit--current', i === index)
  })
  if (scroll && marks[index]) {
    marks[index].scrollIntoView({ block: 'center', inline: 'nearest' })
  }
}

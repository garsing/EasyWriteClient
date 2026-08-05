function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

export function escapeHtml(text) {
  if (typeof text !== 'string') {
    return ''
  }
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
}

export function highlightPython(text) {
  const markers = new Map()
  let markerCounter = 0

  function getMarker() {
    return `__MARKER_${markerCounter++}__`
  }

  let highlighted = text

  highlighted = highlighted.replace(/"""([\s\S]*?)"""/g, (match) => {
    const marker = getMarker()
    markers.set(marker, `<span class="string">${escapeHtml(match)}</span>`)
    return marker
  })
  highlighted = highlighted.replace(/'''([\s\S]*?)'''/g, (match) => {
    const marker = getMarker()
    markers.set(marker, `<span class="string">${escapeHtml(match)}</span>`)
    return marker
  })
  highlighted = highlighted.replace(/"([^"]*(?:\\.[^"]*)*)"/g, (match) => {
    const marker = getMarker()
    markers.set(marker, `<span class="string">${escapeHtml(match)}</span>`)
    return marker
  })
  highlighted = highlighted.replace(/'([^']*(?:\\.[^']*)*)'/g, (match) => {
    const marker = getMarker()
    markers.set(marker, `<span class="string">${escapeHtml(match)}</span>`)
    return marker
  })

  highlighted = highlighted.replace(/#([^\n]*)/g, (match) => {
    const marker = getMarker()
    markers.set(marker, `<span class="comment">${escapeHtml(match)}</span>`)
    return marker
  })

  const keywords = ['nonlocal', 'finally', 'lambda', 'assert', 'global', 'yield', 'async', 'await', 'break', 'class', 'continue', 'def', 'del', 'elif', 'else', 'except', 'False', 'for', 'from', 'if', 'import', 'in', 'is', 'None', 'not', 'or', 'pass', 'raise', 'return', 'True', 'try', 'while', 'with', 'as', 'and']
  keywords.sort((a, b) => b.length - a.length)

  const markerPositions = []
  const markerPattern = /__MARKER_\d+__/g
  let markerMatch
  while ((markerMatch = markerPattern.exec(highlighted)) !== null) {
    markerPositions.push({
      start: markerMatch.index,
      end: markerMatch.index + markerMatch[0].length
    })
  }

  function isInMarker(pos) {
    return markerPositions.some(m => pos >= m.start && pos < m.end)
  }

  keywords.forEach(keyword => {
    const regex = new RegExp(`\\b${escapeRegex(keyword)}\\b`, 'g')
    const matches = []
    let match

    while ((match = regex.exec(highlighted)) !== null) {
      matches.push({
        index: match.index,
        length: match[0].length,
        text: match[0]
      })
    }

    for (let i = matches.length - 1; i >= 0; i--) {
      const m = matches[i]
      const matchStart = m.index
      const matchEnd = matchStart + m.length

      if (!isInMarker(matchStart) && !isInMarker(matchEnd - 1)) {
        const marker = getMarker()
        markers.set(marker, `<span class="keyword">${escapeHtml(m.text)}</span>`)
        highlighted = highlighted.substring(0, matchStart) + marker + highlighted.substring(matchEnd)
      }
    }
  })

  const operators = ['==', '!=', '<=', '>=', '+=', '-=', '*=', '/=', '%=', '**', '//', '<<=', '>>=', '&=', '^=', '|=', '>>', '<<', '+', '-', '*', '/', '%', '=', '<', '>', '!', '&', '|', '^', '~', '(', ')', '[', ']', '{', '}', '.', ',', ':', ';']
  operators.sort((a, b) => b.length - a.length)

  markerPositions.length = 0
  markerPattern.lastIndex = 0
  while ((markerMatch = markerPattern.exec(highlighted)) !== null) {
    markerPositions.push({
      start: markerMatch.index,
      end: markerMatch.index + markerMatch[0].length
    })
  }

  operators.forEach(op => {
    const regex = new RegExp(escapeRegex(op), 'g')
    const matches = []
    let match

    while ((match = regex.exec(highlighted)) !== null) {
      if (!isInMarker(match.index)) {
        matches.push({
          index: match.index,
          length: match[0].length,
          text: match[0]
        })
      }
    }

    for (let i = matches.length - 1; i >= 0; i--) {
      const m = matches[i]
      const marker = getMarker()
      markers.set(marker, `<span class="operator">${escapeHtml(m.text)}</span>`)
      highlighted = highlighted.substring(0, m.index) + marker + highlighted.substring(m.index + m.length)
    }
  })

  markerPositions.length = 0
  markerPattern.lastIndex = 0
  while ((markerMatch = markerPattern.exec(highlighted)) !== null) {
    markerPositions.push({
      start: markerMatch.index,
      end: markerMatch.index + markerMatch[0].length
    })
  }

  highlighted = highlighted.replace(/@(\w+)/g, (match, decoratorName) => {
    if (!isInMarker(match.index)) {
      const marker = getMarker()
      markers.set(marker, `<span class="decorator">@${escapeHtml(decoratorName)}</span>`)
      return marker
    }
    return match
  })

  const builtins = ['abs', 'all', 'any', 'ascii', 'bin', 'bool', 'bytearray', 'bytes', 'callable', 'chr', 'classmethod', 'compile', 'complex', 'delattr', 'dict', 'dir', 'divmod', 'enumerate', 'eval', 'exec', 'filter', 'float', 'format', 'frozenset', 'getattr', 'globals', 'hasattr', 'hash', 'help', 'hex', 'id', 'input', 'int', 'isinstance', 'issubclass', 'iter', 'len', 'list', 'locals', 'map', 'max', 'memoryview', 'min', 'next', 'object', 'oct', 'open', 'ord', 'pow', 'print', 'property', 'range', 'repr', 'reversed', 'round', 'set', 'setattr', 'slice', 'sorted', 'staticmethod', 'str', 'sum', 'super', 'tuple', 'type', 'vars', 'zip', '__import__']
  builtins.sort((a, b) => b.length - a.length)

  markerPositions.length = 0
  markerPattern.lastIndex = 0
  while ((markerMatch = markerPattern.exec(highlighted)) !== null) {
    markerPositions.push({
      start: markerMatch.index,
      end: markerMatch.index + markerMatch[0].length
    })
  }

  builtins.forEach(builtin => {
    const regex = new RegExp(`\\b${escapeRegex(builtin)}\\b(?=\\s*\\()`, 'g')
    const matches = []
    let match

    while ((match = regex.exec(highlighted)) !== null) {
      matches.push({
        index: match.index,
        length: match[0].length,
        text: match[0]
      })
    }

    for (let i = matches.length - 1; i >= 0; i--) {
      const m = matches[i]
      if (!isInMarker(m.index) && !isInMarker(m.index + m.length - 1)) {
        const marker = getMarker()
        markers.set(marker, `<span class="builtin">${escapeHtml(m.text)}</span>`)
        highlighted = highlighted.substring(0, m.index) + marker + highlighted.substring(m.index + m.length)
      }
    }
  })

  markerPositions.length = 0
  markerPattern.lastIndex = 0
  while ((markerMatch = markerPattern.exec(highlighted)) !== null) {
    markerPositions.push({
      start: markerMatch.index,
      end: markerMatch.index + markerMatch[0].length
    })
  }

  const classRegex = /\bclass\s+(\w+)/g
  const classMatches = []
  let classMatch
  while ((classMatch = classRegex.exec(highlighted)) !== null) {
    const classNameStart = classMatch.index + classMatch[0].indexOf(classMatch[1])
    if (!isInMarker(classNameStart)) {
      classMatches.push({
        start: classNameStart,
        end: classNameStart + classMatch[1].length,
        className: classMatch[1],
        fullMatch: classMatch[0]
      })
    }
  }

  for (let i = classMatches.length - 1; i >= 0; i--) {
    const m = classMatches[i]
    const marker = getMarker()
    markers.set(marker, `class <span class="class-name">${escapeHtml(m.className)}</span>`)
    highlighted = highlighted.substring(0, m.start - (m.fullMatch.length - m.className.length - 6)) + marker + highlighted.substring(m.end)
  }

  markerPositions.length = 0
  markerPattern.lastIndex = 0
  while ((markerMatch = markerPattern.exec(highlighted)) !== null) {
    markerPositions.push({
      start: markerMatch.index,
      end: markerMatch.index + markerMatch[0].length
    })
  }

  const funcRegex = /\b([a-zA-Z_][a-zA-Z0-9_]*(?:\.[a-zA-Z_][a-zA-Z0-9_]*)*)\s*(?=\()/g
  const funcMatches = []
  let funcMatch
  while ((funcMatch = funcRegex.exec(highlighted)) !== null) {
    const funcName = funcMatch[1]
    if (keywords.includes(funcName) || builtins.includes(funcName)) {
      continue
    }

    if (funcName.includes('.')) {
      const parts = funcName.split('.')
      const methodName = parts[parts.length - 1]
      const methodStart = funcMatch.index + funcMatch[0].indexOf(methodName)
      if (!isInMarker(methodStart)) {
        funcMatches.push({
          start: methodStart,
          end: methodStart + methodName.length,
          methodName,
          fullMatch: funcMatch[0]
        })
      }
    } else if (!isInMarker(funcMatch.index)) {
      funcMatches.push({
        start: funcMatch.index,
        end: funcMatch.index + funcName.length,
        methodName: funcName,
        fullMatch: funcMatch[0]
      })
    }
  }

  for (let i = funcMatches.length - 1; i >= 0; i--) {
    const m = funcMatches[i]
    const marker = getMarker()
    markers.set(marker, `<span class="function">${escapeHtml(m.methodName)}</span>`)
    highlighted = highlighted.substring(0, m.start) + marker + highlighted.substring(m.end)
  }

  highlighted = highlighted.replace(/\b(\d+\.?\d*)\b/g, (match) => {
    if (!isInMarker(match.index)) {
      const marker = getMarker()
      markers.set(marker, `<span class="number">${escapeHtml(match)}</span>`)
      return marker
    }
    return match
  })

  highlighted = escapeHtml(highlighted)

  const sortedMarkers = Array.from(markers.entries()).reverse()
  sortedMarkers.forEach(([marker, replacement]) => {
    highlighted = highlighted.replace(new RegExp(escapeRegex(marker), 'g'), replacement)
  })

  return highlighted
}

export function highlightYaml(text) {
  const markers = new Map()
  let markerCounter = 0

  function getMarker() {
    return `__YAML_MARKER_${markerCounter++}__`
  }

  let highlighted = text

  highlighted = highlighted.replace(/'((?:[^']|'')*)'/g, (match, content) => {
    const marker = getMarker()
    markers.set(marker, `<span class="string">'${escapeHtml(content)}'</span>`)
    return marker
  })
  highlighted = highlighted.replace(/"([^"]*(?:\\.[^"]*)*)"/g, (match) => {
    const marker = getMarker()
    const content = match.substring(1, match.length - 1)
    markers.set(marker, `<span class="string">"${escapeHtml(content)}"</span>`)
    return marker
  })

  highlighted = highlighted.replace(/#([^\n]*)/g, (match) => {
    const marker = getMarker()
    markers.set(marker, `<span class="comment">${escapeHtml(match)}</span>`)
    return marker
  })

  highlighted = highlighted.replace(/^(\s*)([^#:\n]+?)(\s*):/gm, (match, indent, key, spaces) => {
    if (!match.includes('__YAML_MARKER_')) {
      const marker = getMarker()
      markers.set(marker, `${indent}<span class="key">${escapeHtml(key)}</span>${spaces}:`)
      return marker
    }
    return match
  })

  highlighted = highlighted.replace(/\b(true|false|yes|no|on|off)\b/gi, (match) => {
    if (!match.includes('__YAML_MARKER_')) {
      const marker = getMarker()
      markers.set(marker, `<span class="boolean">${escapeHtml(match)}</span>`)
      return marker
    }
    return match
  })

  highlighted = highlighted.replace(/\b(\d+\.?\d*)\b/g, (match) => {
    if (!match.includes('__YAML_MARKER_')) {
      const marker = getMarker()
      markers.set(marker, `<span class="number">${escapeHtml(match)}</span>`)
      return marker
    }
    return match
  })

  highlighted = escapeHtml(highlighted)

  const sortedMarkers = Array.from(markers.entries()).reverse()
  sortedMarkers.forEach(([marker, replacement]) => {
    highlighted = highlighted.replace(new RegExp(escapeRegex(marker), 'g'), replacement)
  })

  return highlighted
}

export function highlightJson(text) {
  let highlighted = escapeHtml(text)
  highlighted = highlighted.replace(/"([^"]+)":/g, '<span class="key">"$1"</span>:')
  highlighted = highlighted.replace(/:\s*"([^"]*)"/g, ': <span class="string">"$1"</span>')
  highlighted = highlighted.replace(/\b(true|false|null)\b/g, '<span class="boolean">$&</span>')
  highlighted = highlighted.replace(/\b\d+\.?\d*\b/g, '<span class="number">$&</span>')
  return highlighted
}

export function highlightByLanguage(text, language) {
  if (language === 'python') return highlightPython(text)
  if (language === 'yaml') return highlightYaml(text)
  if (language === 'json') return highlightJson(text)
  return escapeHtml(text)
}

/**
 * 按行增量高亮：完整行缓存为 HTML，仅重算末尾未闭合行。
 */
export function highlightPythonIncremental(text, lineCache) {
  if (!text) {
    lineCache.stableLineCount = 0
    lineCache.prefixHtml = ''
    return ''
  }

  const lines = text.split('\n')
  const stableCount = lines.length > 1 ? lines.length - 1 : 0

  if (stableCount < lineCache.stableLineCount) {
    lineCache.stableLineCount = 0
    lineCache.prefixHtml = ''
    return highlightPython(text)
  }

  if (stableCount > lineCache.stableLineCount) {
    for (let i = lineCache.stableLineCount; i < stableCount; i++) {
      const lineHtml = highlightPython(lines[i])
      lineCache.prefixHtml += (lineCache.prefixHtml ? '\n' : '') + lineHtml
    }
    lineCache.stableLineCount = stableCount
  }

  const tailText = lines.slice(stableCount).join('\n')
  if (!tailText) {
    return lineCache.prefixHtml
  }

  const tailHtml = highlightPython(tailText)
  if (!lineCache.prefixHtml) {
    return tailHtml
  }
  return `${lineCache.prefixHtml}\n${tailHtml}`
}

export function resetPythonLineCache(lineCache) {
  lineCache.stableLineCount = 0
  lineCache.prefixHtml = ''
}

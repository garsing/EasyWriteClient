import { IMAGE_DOCUMENT_EXTS } from './kbImageThumb.js'

const MIME_TO_EXT = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/jpg': 'jpg',
  'image/webp': 'webp',
  'image/bmp': 'bmp',
  'image/gif': 'gif'
}

function extFromMime (mime) {
  return MIME_TO_EXT[String(mime || '').toLowerCase()] || ''
}

function pad2 (n) {
  return String(n).padStart(2, '0')
}

function formatStamp () {
  const d = new Date()
  return (
    `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}-` +
    `${pad2(d.getHours())}${pad2(d.getMinutes())}${pad2(d.getSeconds())}`
  )
}

function isImageFile (file) {
  if (!file) return false
  const mimeExt = extFromMime(file.type)
  if (mimeExt && IMAGE_DOCUMENT_EXTS.includes(mimeExt)) return true
  const name = file.name || ''
  const ext = name.includes('.') ? name.split('.').pop().toLowerCase() : ''
  return IMAGE_DOCUMENT_EXTS.includes(ext)
}

function isGenericClipboardName (name) {
  const n = String(name || '').trim().toLowerCase()
  return !n || n === 'blob' || n === 'untitled' || /^image\.(png|jpe?g|webp|bmp|gif)$/.test(n)
}

function namedClipboardFile (file, index) {
  if (!isImageFile(file) || !isGenericClipboardName(file.name)) {
    return file
  }
  const ext =
    extFromMime(file.type) ||
    (file.name && file.name.includes('.') ? file.name.split('.').pop().toLowerCase() : 'png')
  const name = `截图-${formatStamp()}${index > 0 ? `-${index + 1}` : ''}.${ext}`
  return new File([file], name, {
    type: file.type || `image/${ext}`,
    lastModified: file.lastModified || Date.now()
  })
}

export function clipboardHasPlainText (clipboardData) {
  if (!clipboardData) return false
  try {
    const t = clipboardData.getData('text/plain') || clipboardData.getData('text') || ''
    return String(t).trim().length > 0
  } catch (_) {
    return false
  }
}

/** 资源管理器「复制文件」后，文本里常常是一条 Windows 路径 */
export function looksLikeWindowsFilePath (text) {
  const t = String(text || '').trim()
  if (!t || /\r|\n/.test(t)) return false
  return /^[a-zA-Z]:[\\/]/.test(t) || t.startsWith('\\\\')
}

export function clipboardPlainText (clipboardData) {
  if (!clipboardData) return ''
  try {
    return String(clipboardData.getData('text/plain') || clipboardData.getData('text') || '')
  } catch (_) {
    return ''
  }
}

export function clipboardHasFileHint (clipboardData) {
  if (!clipboardData) return false
  try {
    if (clipboardData.files && clipboardData.files.length) return true
    const types = clipboardData.types
    if (types) {
      for (let i = 0; i < types.length; i++) {
        const t = String(types[i]).toLowerCase()
        if (t === 'files' || t.startsWith('image/')) return true
      }
    }
    if (clipboardData.items) {
      for (let i = 0; i < clipboardData.items.length; i++) {
        if (clipboardData.items[i].kind === 'file') return true
      }
    }
    if (looksLikeWindowsFilePath(clipboardPlainText(clipboardData))) return true
  } catch (_) {
    /* ignore */
  }
  return false
}

/** @deprecated 用 clipboardHasFileHint */
export function clipboardHasImageHint (clipboardData) {
  return clipboardHasFileHint(clipboardData)
}

/**
 * 从 paste 事件取出 File[]：截图、复制的图片、资源管理器复制的文件。
 */
export function collectClipboardFiles (clipboardData) {
  if (!clipboardData) return []
  const seen = new Set()
  const out = []

  const push = (file) => {
    if (!file) return
    const key = `${file.size}:${file.type}:${file.name}`
    if (seen.has(key)) return
    seen.add(key)
    out.push(namedClipboardFile(file, out.length))
  }

  if (clipboardData.items && clipboardData.items.length) {
    for (let i = 0; i < clipboardData.items.length; i++) {
      const item = clipboardData.items[i]
      if (item.kind !== 'file') continue
      const f = item.getAsFile()
      if (f) push(f)
    }
  }

  if (!out.length && clipboardData.files && clipboardData.files.length) {
    for (const f of clipboardData.files) push(f)
  }

  return out
}

/** @deprecated 用 collectClipboardFiles */
export function collectClipboardImageFiles (clipboardData) {
  return collectClipboardFiles(clipboardData)
}

export function shouldTryHostClipboard (clipboardData) {
  if (collectClipboardFiles(clipboardData).length) return false
  if (clipboardHasFileHint(clipboardData)) return true
  if (clipboardHasPlainText(clipboardData)) return false
  return true
}

/** @deprecated 用 shouldTryHostClipboard */
export function shouldTryHostClipboardImage (clipboardData) {
  return shouldTryHostClipboard(clipboardData)
}

export function shouldBlockPasteForHost (clipboardData) {
  if (clipboardHasFileHint(clipboardData)) return true
  return !clipboardHasPlainText(clipboardData)
}

export function fileFromHostClipboardPayload (payload) {
  const b64 = payload?.base64
  if (!b64 || typeof b64 !== 'string') return null
  try {
    const bin = atob(b64)
    const bytes = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
    const mime = payload.mimeType || 'image/png'
    const ext = extFromMime(mime) || 'png'
    const name = payload.fileName || `截图-${formatStamp()}.${ext}`
    return new File([bytes], name, { type: mime, lastModified: Date.now() })
  } catch (_) {
    return null
  }
}

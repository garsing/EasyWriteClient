/**
 * 禁止 Ctrl+滚轮 / 触控板捏合缩放页面。
 * WebView2 宿主也会关缩放，此处兜底未走宿主设置的入口。
 */
export function disableUserZoom() {
  if (typeof window === 'undefined' || window.__ewDisableUserZoom) {
    return
  }
  window.__ewDisableUserZoom = true

  const blockCtrlWheel = (e) => {
    if (e.ctrlKey) {
      e.preventDefault()
    }
  }
  window.addEventListener('wheel', blockCtrlWheel, { passive: false, capture: true })
  window.addEventListener('gesturestart', (e) => {
    e.preventDefault()
  }, { passive: false, capture: true })
}

disableUserZoom()

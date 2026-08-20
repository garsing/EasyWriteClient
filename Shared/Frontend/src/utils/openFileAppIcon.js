import wordAppIcon from '../assets/images/word.png'
import wpsAppIcon from '../assets/images/wps.png'
import excelAppIcon from '../assets/images/excel.png'
import pptAppIcon from '../assets/images/ppt.png'
import yiWriteLogoIcon from '../assets/images/yi-write_logo1.png'
import chromeAppIcon from '../assets/images/chrome.png'
import edgeAppIcon from '../assets/images/edge.png'

const GROUP_ORDER = [
  'word',
  'wps',
  'excel',
  'et',
  'ppt',
  'wpp',
  'chrome',
  'edge',
  'yiwrite',
  'other'
]

const GROUP_LABELS = {
  word: 'Word',
  wps: 'WPS 文字',
  excel: 'Excel',
  et: 'WPS 表格',
  ppt: 'PowerPoint',
  wpp: 'WPS 演示',
  chrome: 'Chrome',
  edge: 'Edge',
  yiwrite: '易写浏览器',
  other: '其他'
}

const GROUP_ICONS = {
  word: wordAppIcon,
  wps: wpsAppIcon,
  excel: excelAppIcon,
  et: wpsAppIcon,
  ppt: pptAppIcon,
  wpp: wpsAppIcon,
  chrome: chromeAppIcon,
  edge: edgeAppIcon,
  yiwrite: yiWriteLogoIcon,
  other: wordAppIcon
}

/** 侧栏树第一级分组键（应用）。 */
export function resolveOpenFileAppGroupKey (item) {
  const type = String(item?.appType || item?.app_type || '').toLowerCase()
  const id = String(item?.id || item?.channelId || item?.channel_id || '').toLowerCase()

  if (id.includes('browser:attach:chrome:') || id.startsWith('browser:attach:chrome:')) {
    return 'chrome'
  }
  if (id.includes('browser:attach:edge:') || id.startsWith('browser:attach:edge:')) {
    return 'edge'
  }
  if (type === 'browser' || id.startsWith('browser:agent:') || id.startsWith('browser:')) {
    return 'yiwrite'
  }
  if (type === 'excel') return 'excel'
  if (type === 'ppt') return 'ppt'
  if (type === 'wps') return 'wps'
  if (type === 'et') return 'et'
  if (type === 'wpp') return 'wpp'
  if (type === 'word') return 'word'
  return type || 'other'
}

export function resolveOpenFileAppGroupLabel (groupKey) {
  return GROUP_LABELS[groupKey] || groupKey || '其他'
}

export function resolveOpenFileAppGroupIcon (groupKey) {
  return GROUP_ICONS[groupKey] || wordAppIcon
}

/** 将扁平打开文件列表编成应用树（保持各组内原有顺序）。 */
export function groupOpenFilesByApp (items) {
  const list = Array.isArray(items) ? items : []
  const map = new Map()
  for (const item of list) {
    const key = resolveOpenFileAppGroupKey(item)
    if (!map.has(key)) {
      map.set(key, {
        key,
        label: resolveOpenFileAppGroupLabel(key),
        icon: resolveOpenFileAppGroupIcon(key),
        items: []
      })
    }
    map.get(key).items.push(item)
  }

  const ordered = []
  for (const key of GROUP_ORDER) {
    if (map.has(key)) {
      ordered.push(map.get(key))
      map.delete(key)
    }
  }
  for (const g of map.values()) {
    ordered.push(g)
  }
  return ordered
}

/** 侧栏「打开文件」与输入区芯片共用应用图标（按条目）。 */
export function resolveOpenFileAppIcon (item) {
  return resolveOpenFileAppGroupIcon(resolveOpenFileAppGroupKey(item))
}

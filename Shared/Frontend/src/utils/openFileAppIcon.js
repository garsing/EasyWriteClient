import wordAppIcon from '../assets/images/word.png'
import wpsAppIcon from '../assets/images/wps.png'
import excelAppIcon from '../assets/images/excel.png'
import pptAppIcon from '../assets/images/ppt.png'
import yiWriteLogoIcon from '../assets/images/yi-write_logo1.png'
import chromeAppIcon from '../assets/images/chrome.png'
import edgeAppIcon from '../assets/images/edge.png'

/** 侧栏「打开文件」与输入区芯片共用应用图标。 */
export function resolveOpenFileAppIcon (item) {
  const type = String(item?.appType || item?.app_type || '').toLowerCase()
  const id = String(item?.id || item?.channelId || item?.channel_id || '').toLowerCase()

  // Chrome / Edge 附着页：靠 channel_id，不用文字前缀区分
  if (id.includes('browser:attach:chrome:') || id.startsWith('browser:attach:chrome:')) {
    return chromeAppIcon
  }
  if (id.includes('browser:attach:edge:') || id.startsWith('browser:attach:edge:')) {
    return edgeAppIcon
  }

  if (type === 'excel') return excelAppIcon
  if (type === 'ppt') return pptAppIcon
  if (type === 'wps' || type === 'et' || type === 'wpp') return wpsAppIcon
  // 易写浏览页：与 Desktop 标题栏同源 Logo
  if (type === 'browser' || id.startsWith('browser:')) return yiWriteLogoIcon
  return wordAppIcon
}

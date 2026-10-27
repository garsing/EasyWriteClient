/** 聊天面板整行不画的工具。分组时跳过，不拆开连续工具折叠。 */
export const toolsHidden = [
  'B_todo'
]

export function isHiddenTool (toolName) {
  return toolsHidden.includes(toolName)
}

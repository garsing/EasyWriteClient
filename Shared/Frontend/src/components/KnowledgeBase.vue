<template>
  <div class="knowledge-base-container">
    <!-- 文档列表页 -->
    <DocumentList
      v-if="showDocumentList"
      :knowledge-base-name="selectedKnowledgeBase?.name || ''"
      :knowledge-base-uuid="selectedKnowledgeBase?.uuid || ''"
      @back="handleBackToList"
    />
    
    <!-- 知识库列表页 -->
    <template v-else>
    <!-- 顶部标题栏 -->
    <div class="top-header">
      <div class="header-left">
        <h1 class="title">知识库</h1>
      </div>
    </div>

    <!-- 导航和筛选区域 -->
    <div class="nav-filter-section">
      <div class="nav-tabs">
        <el-tabs v-model="activeTab" class="kb-tabs">
          <el-tab-pane label="全部" name="all"></el-tab-pane>
          <el-tab-pane label="我的关注" name="favorites"></el-tab-pane>
          <el-tab-pane label="最近编辑" name="recent"></el-tab-pane>
        </el-tabs>
      </div>
      
      <div class="filter-actions">
        <el-input
          v-model="searchText"
          placeholder="搜索知识库"
          class="search-input"
          :prefix-icon="Search"
          clearable
        />
        
        <el-select v-model="selectedGroup" placeholder="全部分组" class="filter-select">
          <el-option 
            v-for="option in groupOptions" 
            :key="option.value"
            :label="option.label" 
            :value="option.value"
          ></el-option>
        </el-select>
        
        <el-select v-model="selectedTag" placeholder="全部标签" class="filter-select">
          <el-option 
            v-for="option in tagOptions" 
            :key="option.value"
            :label="option.label" 
            :value="option.value"
          ></el-option>
        </el-select>
        
        <el-select v-model="selectedStatus" placeholder="全部状态" class="filter-select">
          <el-option 
            v-for="option in statusOptions" 
            :key="option.value"
            :label="option.label" 
            :value="option.value"
          ></el-option>
        </el-select>
        
        <el-button type="default" :icon="Plus" class="new-group-btn" @click="openCreateGroupDialog">
          新建分组
        </el-button>
        <el-button type="primary" :icon="Plus" class="new-kb-btn" @click="openCreateDialog">
          新建知识库
        </el-button>
      </div>
    </div>

    <!-- 表格列表 -->
    <div class="table-section">
      <el-table
        :data="filteredKnowledgeBaseList"
        style="width: 100%"
        class="kb-table"
        v-loading="loading"
        :header-cell-style="{ background: '#f5f5f5', color: '#333', fontWeight: 'normal' }"
        :border="false"
        row-key="id"
        :tree-props="{ children: 'children' }"
        default-expand-all
        :indent="0"
        :row-class-name="({ row }) => {
          if (row.isGroup) return 'group-row'
          if (row.isPlaceholder) return 'placeholder-row'
          return ''
        }"
        @row-mouse-enter="handleRowEnter"
        @row-mouse-leave="handleRowLeave"
        @row-click="handleRowClick"
      >
        <el-table-column prop="name" label="知识库" min-width="200">
          <template #default="{ row }">
            <template v-if="row.isGroup">
              <!-- 分组行 -->
              <div class="group-name-cell">
                <span class="group-name">
                  {{ row.name }}
                  <span v-if="row.children && row.children.filter(child => !child.isPlaceholder).length > 0" class="group-count">
                    {{ row.children.filter(child => !child.isPlaceholder).length }}
                  </span>
                </span>
                <!-- 未分组不显示三个点 -->
                <el-dropdown 
                  v-if="row.uuid !== null" 
                  trigger="click" 
                  @command="handleGroupCommand"
                  class="group-dropdown"
                >
                  <el-icon class="more-icon group-more-icon" :class="{ 'is-visible': row.showMore }"><MoreFilled /></el-icon>
                  <template #dropdown>
                    <el-dropdown-menu>
                      <el-dropdown-item :command="{ action: 'create', group: row }">新建知识库</el-dropdown-item>
                      <el-dropdown-item :command="{ action: 'rename', group: row }">重命名</el-dropdown-item>
                      <el-dropdown-item :command="{ action: 'delete', group: row }" class="delete-item" style="color: #f56c6c !important;">删除</el-dropdown-item>
                    </el-dropdown-menu>
                  </template>
                </el-dropdown>
              </div>
            </template>
            <template v-else-if="!row.isPlaceholder">
              <!-- 知识库行 -->
              <div class="kb-name-cell">
                <div class="kb-icon">{{ row.icon }}</div>
                <div class="kb-name-info">
                  <span class="kb-name">{{ row.name }}</span>
                  <el-tag v-if="row.type" size="small" class="kb-type-tag">{{ row.type }}</el-tag>
                  <el-icon 
                    class="star-icon" 
                    :class="{ 'is-favorite': row.isFavorite, 'is-visible': row.isFavorite || row.showStar }" 
                    @click.stop="toggleFavorite(row)"
                  >
                    <StarFilled v-if="row.isFavorite" />
                    <Star v-else />
                  </el-icon>
                </div>
              </div>
            </template>
          </template>
        </el-table-column>
        
        <el-table-column prop="tags" label="标签" min-width="120">
          <template #default="{ row }">
            <div class="tags-cell" v-if="!row.isGroup && !row.isPlaceholder">
              <el-tag 
                v-if="row.tags && row.tags.trim()" 
                size="small" 
                class="tag-item"
                @click.stop="handleAddTag(row)"
                style="cursor: pointer;"
              >
                {{ row.tags }}
              </el-tag>
              <div 
                v-if="row.showAddTag || !row.tags || !row.tags.trim()"
                class="add-tag-container"
                :class="{ 'is-hover': row.showAddTag }"
                @click.stop="handleAddTag(row)"
              >
                <span class="add-tag-text">{{ row.tags && row.tags.trim() ? '编辑标签' : '添加标签' }}</span>
              </div>
            </div>
          </template>
        </el-table-column>
        
        <el-table-column prop="documentCount" label="文档数量" min-width="120">
          <template #default="{ row }">
            <div class="count-cell" v-if="!row.isGroup && !row.isPlaceholder">
              <el-icon><Document /></el-icon>
              <span>{{ row.documentCount }}</span>
            </div>
          </template>
        </el-table-column>
        
        <el-table-column prop="lastEdited" label="最近编辑" min-width="120" sortable>
          <template #default="{ row }">
            <span v-if="!row.isGroup && !row.isPlaceholder">{{ row.lastEdited }}</span>
          </template>
        </el-table-column>
        
        <el-table-column prop="creator" label="创建人" min-width="120">
          <template #default="{ row }">
            <span v-if="!row.isGroup && !row.isPlaceholder">{{ row.creator }}</span>
          </template>
        </el-table-column>
        
        <el-table-column label="操作" min-width="120" fixed="right">
          <template #default="{ row }">
            <div class="action-cell" v-if="!row.isGroup && !row.isPlaceholder">
              <el-link type="primary" :underline="false" @click="handleDelete(row)">
                删除
              </el-link>
              <el-dropdown trigger="hover" placement="bottom-end" @command="handleMoveToGroup" popper-class="move-group-dropdown">
                <el-link type="primary" :underline="false" class="move-link">
                  移动
                </el-link>
                <template #dropdown>
                  <el-dropdown-menu>
                    <el-dropdown-item 
                      v-for="group in groupOptionsForMove" 
                      :key="group.value || 'ungrouped'"
                      :command="{ knowledgeBase: row, group: group }"
                    >
                      {{ group.label }}
                    </el-dropdown-item>
                  </el-dropdown-menu>
                </template>
              </el-dropdown>
            </div>
          </template>
        </el-table-column>
      </el-table>
    </div>
    
    <!-- 新建知识库对话框 -->
    <el-dialog
      v-model="showCreateDialog"
      title="新建知识库"
      width="500px"
      @close="newKnowledgeBaseName = ''; newKnowledgeBaseGroupUuid = null"
    >
      <el-form>
        <el-form-item label="知识库名称">
          <el-input
            v-model="newKnowledgeBaseName"
            placeholder="请输入知识库名称（1-200字符）"
            maxlength="200"
            show-word-limit
            @keyup.enter="() => handleCreateKnowledgeBase()"
          />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="showCreateDialog = false">取消</el-button>
        <el-button type="primary" @click="() => handleCreateKnowledgeBase()">确定</el-button>
      </template>
    </el-dialog>
    
    <!-- 新建分组对话框 -->
    <el-dialog
      v-model="showCreateGroupDialog"
      title="新建分组"
      width="500px"
      @close="newGroupName = ''"
    >
      <el-form>
        <el-form-item label="分组名称">
          <el-input
            v-model="newGroupName"
            placeholder="输入分组名称"
            maxlength="20"
            show-word-limit
            @keyup.enter="handleCreateGroup"
          />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="showCreateGroupDialog = false">取消</el-button>
        <el-button type="primary" @click="handleCreateGroup">确定</el-button>
      </template>
    </el-dialog>
    
    <!-- 重命名分组对话框 -->
    <el-dialog
      v-model="showRenameGroupDialog"
      title="重命名分组"
      width="500px"
      @close="renameGroupName = ''; currentRenameGroup = null"
    >
      <el-form>
        <el-form-item label="分组名称">
          <el-input
            v-model="renameGroupName"
            placeholder="输入分组名称"
            maxlength="200"
            show-word-limit
            @keyup.enter="saveRenameGroup"
          />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="showRenameGroupDialog = false">取消</el-button>
        <el-button type="primary" @click="saveRenameGroup">确定</el-button>
      </template>
    </el-dialog>
    
    <!-- 添加标签对话框 -->
    <el-dialog
      v-model="showTagDialog"
      title="编辑标签"
      width="500px"
      @close="tagInput = ''; currentTagRow = null"
    >
      <el-form>
        <el-form-item label="标签">
          <el-input
            v-model="tagInput"
            placeholder="请输入标签，留空则清空标签"
            maxlength="200"
            show-word-limit
            @keyup.enter="saveTags"
          />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="showTagDialog = false">取消</el-button>
        <el-button type="primary" @click="saveTags">确定</el-button>
      </template>
    </el-dialog>
    </template>
  </div>
</template>

<script setup>
import { ref, onMounted, watch, computed, nextTick } from 'vue'
import { ElMessage, ElMessageBox } from 'element-plus'
import {
  Plus,
  Search,
  Document,
  MoreFilled,
  Star,
  StarFilled
} from '@element-plus/icons-vue'
import unfoldIcon from '@/assets/images/unfold.png'
import DocumentList from './DocumentList.vue'
import {
  getKnowledgeBaseList,
  createKnowledgeBase,
  moveKnowledgeBaseToGroup,
  deleteKnowledgeBase,
  updateKnowledgeBaseFollow,
  updateKnowledgeBaseTags,
  createGroup,
  getGroupList,
  updateGroupName,
  deleteGroup
} from '../services/knowledgeBaseApi'

// 页面状态
const showDocumentList = ref(false)
const selectedKnowledgeBase = ref(null)

// 当前选中的标签页
const activeTab = ref('all')

// 搜索文本
const searchText = ref('')

// 筛选器
const selectedGroup = ref('all')
const selectedTag = ref('all')
const selectedStatus = ref('all')

// 加载状态
const loading = ref(false)

// 原始知识库列表数据（从后端获取的所有数据）
const allKnowledgeBaseList = ref([])

// 所有分组信息（包括空分组）
const allGroups = ref(new Map())

// 筛选后的知识库列表数据（用于显示）
const filteredKnowledgeBaseList = ref([])

// 新建知识库对话框
const showCreateDialog = ref(false)
const newKnowledgeBaseName = ref('')
const newKnowledgeBaseGroupUuid = ref(null)

// 新建分组对话框
const showCreateGroupDialog = ref(false)
const newGroupName = ref('')

// 重命名分组对话框
const showRenameGroupDialog = ref(false)
const renameGroupName = ref('')
const currentRenameGroup = ref(null)

// 添加标签对话框
const showTagDialog = ref(false)
const currentTagRow = ref(null)
const tagInput = ref('')

/**
 * 将 API 返回的数据转换为前端需要的格式
 */
function transformKnowledgeBaseData(apiData) {
  return {
    id: apiData.uuid,
    uuid: apiData.uuid,
    icon: apiData.name ? apiData.name.charAt(0).toUpperCase() : 'N',
    name: apiData.name || '',
    type: apiData.type || null,
    tags: apiData.tags ? apiData.tags.trim() : '',
    documentCount: apiData.document_count || 0,
    lastEdited: formatTime(apiData.last_edited_at),
    lastEditedAt: apiData.last_edited_at || null, // 保存原始时间用于排序
    creator: apiData.created_by_username || '',
    isFavorite: apiData.is_followed || false,
    groupUuid: apiData.group_uuid || null,
    groupName: apiData.group_name || null,
    isEnabled: apiData.is_enabled !== undefined ? apiData.is_enabled : true,
    showAddTag: false,
    showStar: false,
    showMore: false, // 用于知识库行的更多按钮显示
    isGroup: false // 标记是否为分组行
  }
}

/**
 * 格式化时间显示
 */
function formatTime(timeString) {
  if (!timeString) return '未知'
  
  try {
    const date = new Date(timeString)
    const now = new Date()
    const diff = now - date
    const days = Math.floor(diff / (1000 * 60 * 60 * 24))
    
    if (days === 0) {
      const hours = Math.floor(diff / (1000 * 60 * 60))
      if (hours === 0) {
        const minutes = Math.floor(diff / (1000 * 60))
        return minutes <= 0 ? '刚刚' : `${minutes}分钟前`
      }
      return `${hours}小时前`
    } else if (days === 1) {
      return '1天前'
    } else if (days < 7) {
      return `${days}天前`
    } else {
      return date.toLocaleDateString('zh-CN', { year: 'numeric', month: '2-digit', day: '2-digit' })
    }
  } catch (error) {
    return timeString
  }
}

/**
 * 加载知识库列表（先获取分组列表，再获取知识库列表）
 */
async function loadKnowledgeBaseList() {
  loading.value = true
  try {
    // 先获取所有分组列表
    const groupsData = await getGroupList()
    console.log('[KnowledgeBase] 获取到的分组数据:', groupsData)
    
    // 将所有分组保存到 allGroups
    allGroups.value.clear()
    if (Array.isArray(groupsData)) {
      groupsData.forEach(group => {
        if (group.uuid && group.group_name) {
          allGroups.value.set(group.uuid, {
            uuid: group.uuid,
            name: group.group_name
          })
        }
      })
    }
    
    console.log('[KnowledgeBase] 保存的分组数量:', allGroups.value.size)
    
    // 再获取知识库列表
    const data = await getKnowledgeBaseList({})
    
    // 转换数据格式并保存原始数据
    allKnowledgeBaseList.value = Array.isArray(data) 
      ? data.map(transformKnowledgeBaseData)
      : []
    
    // 应用筛选
    applyFilters()
  } catch (error) {
    console.error('加载知识库列表失败:', error)
    ElMessage.error(error.message || '加载知识库列表失败')
  } finally {
    loading.value = false
  }
}

/**
 * 应用筛选条件并组织成树形结构
 */
function applyFilters() {
  let filtered = [...allKnowledgeBaseList.value]
  
  // 根据标签页筛选
  if (activeTab.value === 'favorites') {
    filtered = filtered.filter(item => item.isFavorite)
  } else if (activeTab.value === 'recent') {
    // 最近编辑：按 lastEditedAt 排序（最新的在前）
    filtered = filtered.sort((a, b) => {
      if (!a.lastEditedAt && !b.lastEditedAt) return 0
      if (!a.lastEditedAt) return 1
      if (!b.lastEditedAt) return -1
      const dateA = new Date(a.lastEditedAt)
      const dateB = new Date(b.lastEditedAt)
      return dateB - dateA
    })
  }
  
  // 根据搜索文本筛选
  if (searchText.value && searchText.value.trim()) {
    const searchLower = searchText.value.toLowerCase()
    filtered = filtered.filter(item => 
      item.name.toLowerCase().includes(searchLower)
    )
  }
  
  // 根据分组筛选
  if (selectedGroup.value !== 'all') {
    filtered = filtered.filter(item => {
      if (selectedGroup.value === 'none') {
        // 筛选无分组的知识库
        return !item.groupUuid
      }
      return item.groupUuid === selectedGroup.value
    })
  }
  
  // 根据标签筛选
  if (selectedTag.value !== 'all') {
    filtered = filtered.filter(item => 
      item.tags && item.tags.trim() === selectedTag.value
    )
  }
  
  // 根据状态筛选
  if (selectedStatus.value !== 'all') {
    filtered = filtered.filter(item => {
      if (selectedStatus.value === 'available') {
        return item.isEnabled === true
      } else if (selectedStatus.value === 'unavailable') {
        return item.isEnabled === false
      }
      return true
    })
  }
  
  // 按分组组织成树形结构
  const groupMap = new Map()
  const ungroupedItems = []
  
  // 先创建所有分组节点（包括空分组）
  // 为空分组添加一个隐藏的占位子节点，确保展开符号显示
  allGroups.value.forEach((groupInfo, uuid) => {
    groupMap.set(uuid, {
      id: `group-${uuid}`,
      uuid: uuid,
      name: groupInfo.name,
      groupUuid: uuid,
      groupName: groupInfo.name,
      isGroup: true,
      children: [{ 
        id: `placeholder-${uuid}`, 
        isPlaceholder: true,
        isGroup: false
      }] // 添加占位子节点，确保展开符号显示
    })
  })
  
  // 将知识库按分组分类
  filtered.forEach(item => {
    if (item.groupUuid && item.groupName) {
      // 根据 groupUuid 找到对应的分组
      if (groupMap.has(item.groupUuid)) {
        const group = groupMap.get(item.groupUuid)
        // 如果第一个子节点是占位符，移除它
        if (group.children.length === 1 && group.children[0].isPlaceholder) {
          group.children = []
        }
        group.children.push(item)
      } else {
        // 如果分组不存在于 allGroups 中，创建一个新分组节点
        groupMap.set(item.groupUuid, {
          id: `group-${item.groupUuid}`,
          uuid: item.groupUuid,
          name: item.groupName,
          groupUuid: item.groupUuid,
          groupName: item.groupName,
          isGroup: true,
          children: [item]
        })
      }
    } else {
      ungroupedItems.push(item)
    }
  })
  
  // 构建树形结构数组
  const treeData = []
  
  // 添加所有分组（包括空分组）
  groupMap.forEach(group => {
    treeData.push(group)
  })
  
  // 添加"未分组"分组（即使为空也要显示）
  treeData.push({
    id: 'group-none',
    uuid: null,
    name: '未分组',
    groupUuid: null,
    groupName: '未分组',
    isGroup: true,
    children: ungroupedItems.length > 0 ? ungroupedItems : [{ 
      id: 'placeholder-none', 
      isPlaceholder: true,
      isGroup: false
    }] // 如果为空，添加占位子节点确保展开符号显示
  })
  
  filteredKnowledgeBaseList.value = treeData
}

/**
 * 获取所有可用的分组选项（用于筛选）
 */
const groupOptions = computed(() => {
  const groupMap = new Map()
  allKnowledgeBaseList.value.forEach(item => {
    if (item.groupUuid && item.groupName) {
      // 如果有分组UUID和名称，添加到选项中
      if (!groupMap.has(item.groupUuid)) {
        groupMap.set(item.groupUuid, {
          uuid: item.groupUuid,
          name: item.groupName
        })
      }
    }
  })
  const options = [{ label: '全部分组', value: 'all' }]
  // 如果有无分组的数据，添加"无分组"选项
  const hasNoGroup = allKnowledgeBaseList.value.some(item => !item.groupUuid)
  if (hasNoGroup) {
    options.push({ label: '无分组', value: 'none' })
  }
  // 添加其他分组选项，显示分组名称
  groupMap.forEach(group => {
    options.push({ label: group.name, value: group.uuid })
  })
  return options
})

/**
 * 获取所有可用的分组选项（用于新建知识库）
 */
const groupOptionsForCreate = computed(() => {
  const options = []
  // 添加所有分组选项
  allGroups.value.forEach((groupInfo, uuid) => {
    options.push({ 
      label: groupInfo.name, 
      value: groupInfo.name // 使用分组名称作为值
    })
  })
  return options
})

/**
 * 获取所有可用的标签选项
 */
const tagOptions = computed(() => {
  const tags = new Set()
  allKnowledgeBaseList.value.forEach(item => {
    if (item.tags && item.tags.trim()) {
      tags.add(item.tags.trim())
    }
  })
  const options = [{ label: '全部标签', value: 'all' }]
  Array.from(tags).sort().forEach(tag => {
    options.push({ label: tag, value: tag })
  })
  return options
})

/**
 * 获取所有可用的状态选项
 */
const statusOptions = computed(() => {
  return [
    { label: '全部状态', value: 'all' },
    { label: '可用', value: 'available' },
    { label: '不可用', value: 'unavailable' }
  ]
})

/**
 * 获取所有可用的分组选项（用于移动知识库）
 */
const groupOptionsForMove = computed(() => {
  const options = []
  // 添加"未分组"选项
  options.push({ 
    label: '未分组', 
    value: null 
  })
  // 添加所有分组选项
  allGroups.value.forEach((groupInfo, uuid) => {
    options.push({ 
      label: groupInfo.name, 
      value: uuid // 使用分组UUID作为值
    })
  })
  return options
})

/**
 * 切换关注状态
 */
async function toggleFavorite(row) {
  try {
    const newFollowStatus = !row.isFavorite
    await updateKnowledgeBaseFollow(row.uuid, newFollowStatus)
    row.isFavorite = newFollowStatus
    
    // 更新原始数据中的关注状态
    const originalItem = allKnowledgeBaseList.value.find(item => item.uuid === row.uuid)
    if (originalItem) {
      originalItem.isFavorite = newFollowStatus
    }
    
    ElMessage.success(newFollowStatus ? '已关注' : '已取消关注')
    
    // 重新应用筛选（如果当前在"我的关注"标签页）
    applyFilters()
  } catch (error) {
    console.error('更新关注状态失败:', error)
    ElMessage.error(error.message || '更新关注状态失败')
  }
}

/**
 * 添加标签
 */
function handleAddTag(row) {
  currentTagRow.value = row
  tagInput.value = row.tags || ''
  showTagDialog.value = true
}

/**
 * 保存标签
 */
async function saveTags() {
  if (!currentTagRow.value) return
  
  try {
    const tags = tagInput.value.trim() || null
    await updateKnowledgeBaseTags(currentTagRow.value.uuid, tags)
    
    // 更新本地数据
    currentTagRow.value.tags = tags || ''
    
    // 更新原始数据中的标签
    const originalItem = allKnowledgeBaseList.value.find(item => item.uuid === currentTagRow.value.uuid)
    if (originalItem) {
      originalItem.tags = tags || ''
    }
    
    showTagDialog.value = false
    currentTagRow.value = null
    tagInput.value = ''
    ElMessage.success('标签更新成功')
    
    // 重新应用筛选（标签选项可能已变化）
    applyFilters()
  } catch (error) {
    console.error('更新标签失败:', error)
    ElMessage.error(error.message || '更新标签失败')
  }
}

/**
 * 删除操作
 */
async function handleDelete(row) {
  try {
    await ElMessageBox.confirm(
      `确定要删除知识库"${row.name}"吗？删除后该知识库下的所有文档也将被删除。`,
      '确认删除',
      {
        confirmButtonText: '确定',
        cancelButtonText: '取消',
        type: 'warning'
      }
    )
    
    await deleteKnowledgeBase(row.uuid)
    
    // 从列表中移除
    const index = allKnowledgeBaseList.value.findIndex(item => item.uuid === row.uuid)
    if (index >= 0) {
      allKnowledgeBaseList.value.splice(index, 1)
      applyFilters()
    }
    
    ElMessage.success('删除成功')
  } catch (error) {
    if (error !== 'cancel') {
      console.error('删除知识库失败:', error)
      ElMessage.error(error.message || '删除知识库失败')
    }
  }
}

/**
 * 打开新建知识库对话框
 */
function openCreateDialog(groupUuid = null) {
  newKnowledgeBaseName.value = ''
  newKnowledgeBaseGroupUuid.value = groupUuid
  showCreateDialog.value = true
}

/**
 * 打开新建分组对话框
 */
function openCreateGroupDialog() {
  newGroupName.value = ''
  showCreateGroupDialog.value = true
}

/**
 * 创建知识库
 */
async function handleCreateKnowledgeBase() {
  if (!newKnowledgeBaseName.value.trim()) {
    ElMessage.warning('请输入知识库名称')
    return
  }
  
  try {
    // 获取分组UUID（如果选择了分组）
    // 确保 groupUuid 是字符串或 null，不会是对象
    let groupUuid = newKnowledgeBaseGroupUuid.value
    if (groupUuid && typeof groupUuid !== 'string') {
      console.warn('[创建知识库] groupUuid 不是字符串，重置为 null:', groupUuid)
      groupUuid = null
    }
    groupUuid = groupUuid || null
    
    console.log('[创建知识库] 分组UUID:', groupUuid, '类型:', typeof groupUuid)
    console.log('[创建知识库] 知识库名称:', newKnowledgeBaseName.value.trim())
    
    const data = await createKnowledgeBase(
      newKnowledgeBaseName.value.trim(),
      groupUuid
    )
    
    console.log('[创建知识库] 返回数据:', data)
    
    // 重新加载数据
    await loadKnowledgeBaseList()
    
    showCreateDialog.value = false
    newKnowledgeBaseName.value = ''
    newKnowledgeBaseGroupUuid.value = null
    ElMessage.success('创建成功')
  } catch (error) {
    console.error('创建知识库失败:', error)
    ElMessage.error(error.message || '创建知识库失败')
  }
}

/**
 * 创建分组
 */
async function handleCreateGroup() {
  if (!newGroupName.value.trim()) {
    ElMessage.warning('请输入分组名称')
    return
  }
  
  if (newGroupName.value.trim().length > 20) {
    ElMessage.warning('分组名称不能超过20个字符')
    return
  }
  
  try {
    const result = await createGroup(newGroupName.value.trim())
    
    // 将新创建的分组添加到 allGroups
    if (result && result.uuid && result.group_name) {
      allGroups.value.set(result.uuid, {
        uuid: result.uuid,
        name: result.group_name
      })
      
      showCreateGroupDialog.value = false
      newGroupName.value = ''
      ElMessage.success('创建成功')
      
      // 重新应用筛选以显示新分组
      applyFilters()
    } else {
      // 如果返回数据不完整，重新加载列表（会重新获取分组列表）
      showCreateGroupDialog.value = false
      newGroupName.value = ''
      ElMessage.success('创建成功')
      await loadKnowledgeBaseList()
    }
  } catch (error) {
    console.error('创建分组失败:', error)
    ElMessage.error(error.message || '创建分组失败')
  }
}

/**
 * 打开重命名分组对话框
 */
function openRenameGroupDialog(group) {
  currentRenameGroup.value = group
  renameGroupName.value = group.name
  showRenameGroupDialog.value = true
}

/**
 * 处理移动知识库到分组
 */
async function handleMoveToGroup(command) {
  const { knowledgeBase, group } = command
  try {
    const groupUuid = group.value !== null ? group.value : null
    await moveKnowledgeBaseToGroup(knowledgeBase.uuid, groupUuid)
    
    // 重新加载数据
    await loadKnowledgeBaseList()
    
    ElMessage.success(groupUuid ? `已移动到"${group.label}"` : '已移动到"未分组"')
  } catch (error) {
    console.error('移动知识库失败:', error)
    ElMessage.error(error.message || '移动知识库失败')
  }
}

/**
 * 处理分组菜单命令
 */
function handleGroupCommand(command) {
  const { action, group } = command
  
  if (action === 'create') {
    // 新建知识库，传递当前分组的UUID
    openCreateDialog(group.uuid)
  } else if (action === 'rename') {
    // 重命名分组
    openRenameGroupDialog(group)
  } else if (action === 'delete') {
    // 删除分组
    handleDeleteGroup(group)
  }
}

/**
 * 删除分组
 */
async function handleDeleteGroup(group) {
  try {
    await ElMessageBox.confirm(
      `确定要删除分组"${group.name}"吗？删除后该分组下的知识库将移动到"未分组"。`,
      '确认删除',
      {
        confirmButtonText: '确定',
        cancelButtonText: '取消',
        type: 'warning'
      }
    )
    
    // 调用删除分组API
    await deleteGroup(group.uuid)
    
    // 从 allGroups 中移除
    allGroups.value.delete(group.uuid)
    
    // 重新加载数据
    await loadKnowledgeBaseList()
    
    ElMessage.success('删除成功')
  } catch (error) {
    if (error !== 'cancel') {
      console.error('删除分组失败:', error)
      ElMessage.error(error.message || '删除分组失败')
    }
  }
}

/**
 * 保存重命名分组
 */
async function saveRenameGroup() {
  if (!renameGroupName.value.trim()) {
    ElMessage.warning('请输入分组名称')
    return
  }
  
  if (renameGroupName.value.trim().length > 200) {
    ElMessage.warning('分组名称不能超过200个字符')
    return
  }
  
  if (!currentRenameGroup.value) {
    return
  }
  
  const newGroupName = renameGroupName.value.trim()
  const oldGroupName = currentRenameGroup.value.name
  
  // 如果名称没有变化，直接关闭对话框
  if (newGroupName === oldGroupName) {
    showRenameGroupDialog.value = false
    renameGroupName.value = ''
    currentRenameGroup.value = null
    return
  }
  
  try {
    // 调用更新分组名称API
    await updateGroupName(currentRenameGroup.value.uuid, newGroupName)
    
    // 更新本地分组数据
    const groupInfo = allGroups.value.get(currentRenameGroup.value.uuid)
    if (groupInfo) {
      groupInfo.name = newGroupName
    }
    
    showRenameGroupDialog.value = false
    renameGroupName.value = ''
    currentRenameGroup.value = null
    ElMessage.success('重命名成功')
    
    // 重新加载数据
    await loadKnowledgeBaseList()
  } catch (error) {
    console.error('重命名分组失败:', error)
    ElMessage.error(error.message || '重命名分组失败')
  }
}

// 打开文档列表页
const handleOpenDocumentList = (row) => {
  selectedKnowledgeBase.value = row
  showDocumentList.value = true
}

// 表格行点击事件
const handleRowClick = (row, column, event) => {
  // 如果是分组行或占位符行，不触发跳转
  if (row.isGroup || row.isPlaceholder) {
    return
  }
  
  // 检查点击的目标元素，如果是按钮、链接、下拉菜单等交互元素，则不触发跳转
  const target = event.target
  const isInteractiveElement = target.closest('button') || 
                                target.closest('a') || 
                                target.closest('.el-link') ||
                                target.closest('.el-dropdown') ||
                                target.closest('.el-icon') ||
                                target.closest('.el-switch') ||
                                target.closest('.el-checkbox') ||
                                target.closest('.add-tag-container') ||
                                target.closest('.tag-item') ||
                                target.closest('.tags-cell')
  
  if (!isInteractiveElement) {
    handleOpenDocumentList(row)
  }
}

// 返回知识库列表页
const handleBackToList = () => {
  showDocumentList.value = false
  selectedKnowledgeBase.value = null
  // 确保表格重新渲染后删除多余的缩进元素
  nextTick(() => {
    removeIndentElements()
    // 延迟再次检查，确保 Element Plus 完全渲染完成
    setTimeout(() => {
      removeIndentElements()
    }, 100)
  })
}

// 鼠标进入行
const handleRowEnter = (row) => {
  if (row.isGroup) {
    row.showMore = true
  } else if (!row.isPlaceholder) {
    row.showAddTag = true
    row.showStar = true
  }
}

// 鼠标离开行
const handleRowLeave = (row) => {
  if (row.isGroup) {
    row.showMore = false
  } else if (!row.isPlaceholder) {
    row.showAddTag = false
    // 如果已关注，保持显示；否则隐藏
    if (!row.isFavorite) {
      row.showStar = false
    }
  }
}

// 监听标签页变化
watch(activeTab, () => {
  applyFilters()
})

// 监听搜索文本变化（防抖）
let searchTimeout = null
watch(searchText, () => {
  if (searchTimeout) {
    clearTimeout(searchTimeout)
  }
  searchTimeout = setTimeout(() => {
    applyFilters()
  }, 500)
})

// 监听筛选器变化
watch(selectedGroup, () => {
  applyFilters()
})

watch(selectedTag, () => {
  applyFilters()
})

watch(selectedStatus, () => {
  applyFilters()
})

// 监听 showDocumentList 变化，确保返回时删除多余的缩进元素
watch(showDocumentList, (newVal) => {
  if (!newVal) {
    // 从文档列表页返回时，删除多余的缩进元素
    nextTick(() => {
      removeIndentElements()
      // 延迟再次检查，确保 Element Plus 完全渲染完成
      setTimeout(() => {
        removeIndentElements()
      }, 100)
    })
  }
})

// 删除 Element Plus 自动生成的缩进元素
function removeIndentElements() {
  nextTick(() => {
    const indentElements = document.querySelectorAll('.kb-table .el-table__indent, .kb-table .el-table__placeholder')
    indentElements.forEach(el => {
      if (el && el.parentNode) {
        el.parentNode.removeChild(el)
      }
    })
  })
}

// 组件挂载时加载数据
onMounted(() => {
  loadKnowledgeBaseList()
  // 删除缩进元素
  removeIndentElements()
  // 使用 MutationObserver 监听 DOM 变化，持续删除新生成的缩进元素
  const observer = new MutationObserver(() => {
    removeIndentElements()
  })
  const tableElement = document.querySelector('.kb-table')
  if (tableElement) {
    observer.observe(tableElement, {
      childList: true,
      subtree: true
    })
  }
})
</script>

<style scoped>
/* 重置基础样式 */
.knowledge-base-container {
  display: flex;
  flex-direction: column;
  height: 100vh;
  width: 100%;
  background-color: #f5f5f5;
  margin: 0;
  padding: 0;
  box-sizing: border-box;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
}

/* 顶部标题栏 */
.top-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 20px 24px;
  background-color: #fff;
  border-bottom: 1px solid #e0e0e0;
}

.header-left {
  display: flex;
  align-items: center;
}

.title {
  margin: 0;
  font-size: 24px;
  font-weight: 600;
  color: #333;
  line-height: 1.5;
}


/* 导航和筛选区域 */
.nav-filter-section {
  padding: 16px 24px;
  background-color: #fff;
  border-bottom: 1px solid #e0e0e0;
}

.nav-tabs {
  margin-bottom: 16px;
}

.kb-tabs :deep(.el-tabs__item) {
  padding: 0 20px;
  font-size: 14px;
}

.kb-tabs :deep(.el-tabs__item.is-active) {
  color: #722ed1;
  font-weight: 500;
}

.kb-tabs :deep(.el-tabs__active-bar) {
  background-color: #722ed1;
}

.filter-actions {
  display: flex;
  align-items: center;
  gap: 12px;
}

.search-input {
  width: 300px;
}

.filter-select {
  width: 120px;
}

.new-group-btn {
  margin-left: auto;
}

.new-kb-btn {
  margin-left: 8px;  /* 新建分组和新建知识库按钮之间的间距 */
}

/* 表格区域 */
.table-section {
  flex: 1;
  padding: 24px;
  overflow: auto;
  background-color: #fff;
}

.kb-table {
  background-color: #fff;
}

/* 确保 Element Plus 组件样式正确应用 */
.kb-table :deep(.el-table__header-wrapper) {
  background-color: #f5f5f5;
}

.kb-table :deep(.el-table__body-wrapper) {
  background-color: #fff;
  border-bottom: none !important;
}

.kb-table :deep(.el-table__inner-wrapper) {
  border-bottom: none !important;
}

.kb-table :deep(.el-table__inner-wrapper::after) {
  display: none !important;
}

.kb-table :deep(.el-table__inner-wrapper::before) {
  display: none !important;
}

.kb-table :deep(.el-table th) {
  background-color: #f5f5f5 !important;
  color: #333;
  font-weight: normal;
}

.kb-table :deep(.el-table td) {
  background-color: #fff;
  padding-top: 16px;
  padding-bottom: 16px;
  border-bottom: none !important;
  vertical-align: middle !important;
}

.kb-table :deep(.el-table__body-wrapper .el-table__body tr td) {
  vertical-align: middle !important;
}

.kb-table :deep(.el-table td .cell) {
  display: flex;
  align-items: center;
  justify-content: flex-start;
  vertical-align: middle;
  min-height: 100%;
  height: 100%;
  line-height: normal;
  align-content: center;
}

/* 完全移除 Element Plus 表格树形结构的缩进元素，因为我们已经用 padding-left 实现了缩进 */
.kb-table :deep(.el-table td .cell .el-table__indent),
.kb-table :deep(.el-table td .cell .el-table__placeholder) {
  display: none !important;
  width: 0 !important;
  padding: 0 !important;
  margin: 0 !important;
  visibility: hidden !important;
  position: absolute !important;
  left: -9999px !important;
}

/* 确保知识库名称列的单元格内容垂直居中 */
.kb-table :deep(.el-table td .cell .group-name-cell),
.kb-table :deep(.el-table td .cell .kb-name-cell) {
  display: flex;
  align-items: center;
  flex: 1;
  align-self: center;
  min-height: auto;
  height: auto;
}

.kb-table :deep(.el-table__row) {
  border-bottom: none !important;
}

.kb-table :deep(.el-table__row.group-row) {
  border-bottom: none !important;
}

.kb-table :deep(.el-table__body tr) {
  border-bottom: none !important;
}

.kb-table :deep(.el-table__body td) {
  border-bottom: none !important;
}

/* 移除表格底部的边框 */
.kb-table :deep(.el-table__body-wrapper) {
  border-bottom: none !important;
}

.kb-table :deep(.el-table) {
  border-bottom: none !important;
}

.kb-table :deep(.el-table::after) {
  display: none !important;
}

.kb-table :deep(.el-table::before) {
  display: none !important;
}

.kb-table :deep(.el-table tr:hover > td) {
  background-color: #f5f7fa;
}

.kb-table :deep(.el-table__row) {
  cursor: pointer;
  height: auto;
}

.kb-table :deep(.el-table__row:hover > td) {
  background-color: #f5f7fa;
}

/* 分组行不显示 hover 效果 */
.kb-table :deep(.el-table__row.group-row:hover > td) {
  background-color: #fff !important;
}

.kb-table :deep(.el-table__row.group-row) {
  cursor: default;
}

/* 分组行的展开图标样式 - 使用自定义图片 */
.kb-table :deep(.el-table__row.group-row .el-table__expand-icon) {
  display: inline-flex !important;
  align-items: center;
  vertical-align: middle;
  line-height: 1.5;
  margin-right: 8px;
  flex-shrink: 0;
  visibility: visible !important;
  opacity: 1 !important;
  width: 12px;
  height: 12px;
  background-image: url('@/assets/images/unfold.png');
  background-size: contain;
  background-repeat: no-repeat;
  background-position: center;
  color: transparent !important;
  font-size: 0 !important;
  position: relative;
}

/* 隐藏默认的展开图标内容（包括 SVG 和伪元素） */
.kb-table :deep(.el-table__row.group-row .el-table__expand-icon::before),
.kb-table :deep(.el-table__row.group-row .el-table__expand-icon svg),
.kb-table :deep(.el-table__row.group-row .el-table__expand-icon i) {
  display: none !important;
}

/* 展开状态 - 旋转图片（向下） */
.kb-table :deep(.el-table__row.group-row .el-table__expand-icon--expanded) {
  transform: rotate(0deg);
  transition: transform 0.3s ease;
}

/* 折叠状态 - 旋转图片（向右） */
.kb-table :deep(.el-table__row.group-row .el-table__expand-icon:not(.el-table__expand-icon--expanded)) {
  transform: rotate(-90deg);
  transition: transform 0.3s ease;
}

/* 为空分组强制显示展开图标占位符 */
.kb-table :deep(.el-table__row.group-row:not(.el-table__row--level-0) .el-table__expand-icon),
.kb-table :deep(.el-table__row.group-row .el-table__expand-icon-placeholder) {
  display: inline-flex !important;
  visibility: visible !important;
  opacity: 1 !important;
}

/* 确保展开图标和内容在同一行 */
.kb-table :deep(.el-table__row.group-row > td:first-child) {
  display: flex;
  align-items: center;
  flex-wrap: nowrap;
  line-height: 1.5;
}

.kb-table :deep(.el-table__row.group-row > td:first-child .cell) {
  display: flex;
  align-items: center;
  flex-wrap: nowrap;
  width: 100%;
}

/* 知识库行添加缩进，区分层级 */
.kb-table :deep(.el-table__row:not(.group-row) > td:first-child) {
  padding-left: 24px;
}

/* 隐藏占位符行 */
.kb-table :deep(.el-table__row.placeholder-row) {
  display: none !important;
}

/* 分组行的单元格样式 */
.kb-table :deep(.el-table__row.group-row > td) {
  padding-top: 16px;
  padding-bottom: 16px;
  background-color: #fff;
  vertical-align: middle;
}

.kb-table :deep(.el-table__row.group-row > td .cell) {
  display: flex;
  align-items: center;
  vertical-align: middle;
}

.kb-table :deep(.el-table__row.group-row > td:first-child) {
  padding-left: 0;
  line-height: 1.5;
}

.kb-table :deep(.el-table__row.group-row > td:first-child .cell) {
  display: flex;
  align-items: center;
  flex-wrap: nowrap;
  line-height: 1.5;
}

/* 确保 hover 时显示添加标签按钮 */
.kb-table :deep(.el-table__row:hover .add-tag-btn) {
  display: inline-block;
}

.kb-name-cell {
  display: flex;
  align-items: center;
  gap: 12px;
  margin: 0;
  padding: 0;
  line-height: normal;
  align-self: stretch;
  flex: 1;
  justify-content: flex-start;
}

.group-name-cell {
  display: flex;
  align-items: center;
  gap: 12px;
  height: 100%;
  min-height: 100%;
  margin: 0;
  padding: 0;
  line-height: normal;
}

.group-name {
  font-size: 14px;
  color: #333;
  font-weight: 500;
  display: inline-block;
  margin-left: 0;
  padding-left: 0;
  line-height: 1.5;
  vertical-align: middle;
}

.group-count {
  display: inline-block;
  background-color: #f0f0f0;
  color: #666;
  border-radius: 10px;
  padding: 2px 8px;
  margin-left: 6px;
  font-size: 12px;
  font-weight: normal;
  min-width: 20px;
  text-align: center;
}

.group-dropdown {
  display: inline-flex;
  align-items: center;
  margin-left: 8px;
}

.group-dropdown .group-more-icon {
  font-size: 24px;
  padding: 4px 6px;
  border-radius: 4px;
  transition: background-color 0.2s;
}

.group-dropdown:hover .group-more-icon {
  background-color: #f5f5f5;
}

.kb-icon {
  width: 32px;
  height: 32px;
  border-radius: 50%;
  background: linear-gradient(135deg, #b37feb 0%, #722ed1 100%);
  display: flex;
  align-items: center;
  justify-content: center;
  color: #fff;
  font-weight: 600;
  font-size: 14px;
  flex-shrink: 0;
  align-self: center;
  margin: 0;
  padding: 0;
}

.kb-name-info {
  display: flex;
  align-items: center;
  gap: 8px;
  align-self: center;
  line-height: normal;
  margin: 0;
  padding: 0;
}

.kb-name {
  color: #333;
  font-size: 14px;
  line-height: normal;
  display: inline-block;
  vertical-align: middle;
  margin: 0;
  padding: 0;
}

.star-icon {
  font-size: 16px;
  color: #999;
  cursor: pointer;
  transition: all 0.2s;
  opacity: 0;
  visibility: hidden;
}

.star-icon.is-visible {
  opacity: 1 !important;
  visibility: visible !important;
}

.star-icon.is-favorite {
  opacity: 1 !important;
  visibility: visible !important;
  color: #f5a623;
}

.star-icon:hover {
  color: #f5a623;
}

/* 确保 hover 时显示星星图标 */
.kb-table :deep(.el-table__row:hover .star-icon) {
  opacity: 1 !important;
  visibility: visible !important;
}

.kb-type-tag {
  background-color: #f0f0f0;
  color: #666;
  border: none;
}

.count-cell {
  display: flex;
  align-items: center;
  gap: 6px;
  color: #666;
}

.count-cell .el-icon {
  font-size: 16px;
}

.empty-text {
  color: #999;
}

.tags-cell {
  display: flex;
  align-items: center;
  gap: 8px;
  position: relative;
}

.tags-list {
  display: flex;
  align-items: center;
  gap: 4px;
}

.add-tag-container {
  display: inline-block;
  padding: 2px 8px;
  border: 1px dashed #d9d9d9;
  border-radius: 4px;
  background-color: #fff;
  cursor: pointer;
  transition: all 0.2s;
  opacity: 0;
}

.add-tag-container.is-hover {
  opacity: 1;
}

.add-tag-text {
  font-size: 11px;
  color: #666;
  line-height: 1.4;
}

.add-tag-container:hover {
  border-color: #409eff;
  color: #409eff;
}

.add-tag-container:hover .add-tag-text {
  color: #409eff;
}

/* 确保 hover 时显示添加标签容器 */
.kb-table :deep(.el-table__row:hover .add-tag-container) {
  opacity: 1 !important;
}

.action-cell {
  display: flex;
  align-items: center;
  gap: 12px;
}

.more-icon {
  font-size: 18px;
  color: #999;
  cursor: pointer;
  opacity: 0;
  visibility: hidden;
  transition: all 0.2s;
}

.more-icon:hover {
  color: #333;
}

.more-icon.is-visible {
  opacity: 1 !important;
  visibility: visible !important;
}

.group-more-icon {
  opacity: 0;
  visibility: hidden;
  transition: opacity 0.2s, visibility 0.2s, background-color 0.2s;
}

.group-more-icon.is-visible {
  opacity: 1 !important;
  visibility: visible !important;
}

/* 分组行 hover 时显示更多图标 */
.kb-table :deep(.el-table__row.group-row:hover .group-more-icon) {
  opacity: 1 !important;
  visibility: visible !important;
}

/* 删除菜单项样式 - 红色 */
:deep(.delete-item),
:deep(.delete-item span),
:deep(.el-dropdown-menu__item.delete-item),
:deep(.el-dropdown-menu__item.delete-item span) {
  color: #f56c6c !important;
}

:deep(.delete-item:hover),
:deep(.delete-item:hover span),
:deep(.el-dropdown-menu__item.delete-item:hover),
:deep(.el-dropdown-menu__item.delete-item:hover span) {
  color: #f56c6c !important;
  background-color: #fef0f0 !important;
}

.tag-item {
  margin-right: 4px;
}

.move-link {
  cursor: pointer;
}

/* 移动至分组下拉菜单样式 - 当分组太多时显示滚动条 */
:deep(.move-group-dropdown .el-dropdown-menu) {
  max-height: 300px;
  overflow-y: auto;
}

/* 自定义滚动条样式（可选，让滚动条更美观） */
:deep(.move-group-dropdown .el-dropdown-menu::-webkit-scrollbar) {
  width: 6px;
}

:deep(.move-group-dropdown .el-dropdown-menu::-webkit-scrollbar-track) {
  background: #f1f1f1;
  border-radius: 3px;
}

:deep(.move-group-dropdown .el-dropdown-menu::-webkit-scrollbar-thumb) {
  background: #c1c1c1;
  border-radius: 3px;
}

:deep(.move-group-dropdown .el-dropdown-menu::-webkit-scrollbar-thumb:hover) {
  background: #a8a8a8;
}
</style>

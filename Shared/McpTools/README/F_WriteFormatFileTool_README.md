# F_WriteFormatFileTool 使用指南

## 概述

`F_WriteFormatFileTool` 是一个用于创建Word文档格式配置文件的工具，支持创建**表格**和**图表**两种类型的YAML配置文件。

## 支持的格式类型

### 1. 表格格式 (Table)
用于 `create_table_from_yaml` 工具，创建Word表格。

### 2. 图表格式 (Chart)
用于 `create_chart_from_yaml` 工具，创建Word图表（柱状图、折线图、饼图）。

## 使用方法

通过MCP调用 `write_format_file` 工具，传入YAML格式的配置内容。

### 参数
- `content`: YAML格式的配置文件内容

### 返回值
- `filename`: 生成的文件名（带时间戳）
- `file_path`: 完整文件路径
- `directory`: 保存目录（format_files）
- `message`: 成功消息

## 表格格式示例

```yaml
Location:
  Type: char_pos
  Start: 120
  End: 120

Table:
  Rows: 3
  Cols: 4
  ColWidths: [60.0, 80.0, 100.0, 120.0]
  Style: 'Grid Table 4 - Accent 1'
  TableFontColor: 'FF0000'

Merge:
  - [1,3, 1,4]

Data:
  - ['姓名', '部门', 'Q1销量', 'Q2销量']
  - ['张三', '华东区', '120', '150']
  - ['李四', '华南区', '90', '<b style="color:red">180</b>']
```

## 图表格式示例

### 柱状图
```yaml
operation: create_chart
location:
  type: after
  anchor: "end"
chart:
  type: column_clustered
  title: 月度销售统计
  data:
    - ['月份', '销售额', '利润']
    - ['1月', 15000, 3000]
    - ['2月', 18000, 4500]
  theme: Office
  width: 500
  height: 350
  showDataLabels: true
  legendPosition: bottom
  plotColor: "4472C4"
  gapWidth: 150
```

### 折线图
```yaml
operation: create_chart
location:
  type: after
  anchor: "end"
chart:
  type: line
  title: 温度变化趋势图
  data:
    - ['日期', '北京', '上海']
    - ['1月', 2, 8]
    - ['2月', 5, 12]
  theme: Office
  width: 550
  height: 400
  dataMarkers: true
  markerSize: 6
  lineWeight: 2.5
  yAxisMin: 0
  yAxisMax: 35
```

### 饼图
```yaml
operation: create_chart
location:
  type: after
  anchor: "end"
chart:
  type: pie3d
  title: 市场份额分布图
  data:
    - ['产品', '份额']
    - ['产品A', 35]
    - ['产品B', 25]
  theme: Office
  width: 450
  height: 450
  showDataLabels: true
  dataLabelType: both
  legendPosition: right
  explosion: 10
```

## 配置文件说明

### 表格配置参数

| 参数 | 类型 | 说明 |
|------|------|------|
| Location.Type | string | 位置类型（char_pos） |
| Location.Start | int | 起始字符位置（0-based） |
| Location.End | int | 结束字符位置 |
| Table.Rows | int | 表格行数 |
| Table.Cols | int | 表格列数 |
| Table.ColWidths | float[] | 列宽数组（磅） |
| Table.Style | string | Word内置表样式名 |
| Table.TableFontColor | string | 表格字体颜色（十六进制） |
| Merge | int[][] | 单元格合并配置 |
| Data | string[][] | 表格数据内容 |

### 图表配置参数

| 参数 | 类型 | 说明 |
|------|------|------|
| operation | string | 必须为"create_chart" |
| location.type | string | 位置类型（after） |
| location.anchor | string | 锚点位置（end） |
| chart.type | string | 图表类型 |
| chart.title | string | 图表标题 |
| chart.data | array | 图表数据 |
| chart.width | int | 图表宽度（磅） |
| chart.height | int | 图表高度（磅） |
| chart.theme | string | 主题样式 |
| chart.showDataLabels | bool | 是否显示数据标签 |
| chart.legendPosition | string | 图例位置 |

### 图表类型特有参数

#### 柱状图
- `gapWidth`: 柱间隙百分比（10-500）

#### 折线图
- `dataMarkers`: 是否显示数据点
- `markerSize`: 标记大小
- `lineWeight`: 线条粗细
- `yAxisMin/Max/MajorUnit`: Y轴设置

#### 饼图
- `dataLabelType`: 标签类型（value/percent/both）
- `explosion`: 分离程度（0-100）

## 文件保存

- 保存目录：`tmp/{username}/format_files/`
- 文件命名：`format_YYYYMMDD_HHMMSS_fff.yaml`
- 编码：UTF-8

## 注意事项

1. 表格数据行数必须等于Table.Rows
2. 表格数据列数必须等于Table.Cols
3. 图表数据第一行为标题行
4. 颜色值使用十六进制RGB格式，无#前缀
5. HTML样式仅在表格中支持

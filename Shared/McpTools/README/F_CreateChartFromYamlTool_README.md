# F_CreateChartFromYamlTool 使用说明

## 概述

`F_CreateChartFromYamlTool` 是一个MCP工具，用于从YAML配置文件在Word文档中创建图表。该工具支持创建柱状图、折线图和3D饼图。

## 支持的图表类型

### 1. 柱状图 (column_clustered)
```yaml
operation: create_chart
location: {type: after, anchor: "end"}
chart:
  type: column_clustered
  title: 各季度产品销量对比
  data:
    - [季度, 产品A, 产品B]
    - [Q1, 120, 80]
    - [Q2, 180, 150]
  theme: Office
  width: 400
  height: 300
  show_data_labels: true
  legend_position: bottom
  plot_color: "ED7D31"
  title_font_size: 14
  title_font_bold: true
  title_font_color: "000000"
  gridlines: major
  gap_width: 80
```

### 2. 折线图 (line)
```yaml
operation: create_chart
location: {type: after, anchor: "end"}
chart:
  type: line
  title: 温度变化趋势
  data:
    - [日期, 最高温, 最低温]
    - [1月, 8, -2]
    - [2月, 12, 2]
  theme: Office
  width: 420
  height: 280
  show_data_labels: false
  data_markers: true
  marker_size: 7
  line_weight: 2.25
  legend_position: right
  plot_color: "4472C4"
  title_font_size: 12
  title_font_bold: false
  gridlines: major
  y_axis_min: -5
  y_axis_max: 30
  y_axis_major_unit: 5
```

### 3. 3D饼图 (pie3d)
```yaml
operation: create_chart
location: {type: after, anchor: "end"}
chart:
  type: pie3d
  title: 市场份额分布
  data:
    - [产品, 份额]
    - [产品A, 35]
    - [产品B, 25]
  theme: Office
  width: 350
  height: 350
  show_data_labels: true
  data_label_type: percent
  legend_position: right
  plot_color: "70AD47"
  title_font_size: 14
  title_font_bold: true
  explosion: 5
```

## 配置参数说明

### 通用参数
- `operation`: 必须为 "create_chart"
- `location.type`: 插入位置类型，目前支持 "after"
- `location.anchor`: 锚点位置，目前支持 "end"

### 图表通用配置
- `chart.type`: 图表类型 (column_clustered|line|pie3d)
- `chart.title`: 图表标题
- `chart.data`: 数据数组，第一行为标题行
- `chart.theme`: 主题样式 (目前支持 "Office")
- `chart.width`: 图表宽度（磅）
- `chart.height`: 图表高度（磅）
- `chart.show_data_labels`: 是否显示数据标签
- `chart.legend_position`: 图例位置 (top|bottom|left|right)
- `chart.plot_color`: 主色系（RGB十六进制，无#前缀）
- `chart.title_font_size`: 标题字体大小
- `chart.title_font_bold`: 标题是否加粗
- `chart.title_font_color`: 标题颜色（RGB十六进制）
- `chart.gridlines`: 网格线类型 (none|major|minor|both)

### 柱状图特有参数
- `chart.gap_width`: 柱间隙百分比 (10-500)

### 折线图特有参数
- `chart.data_markers`: 是否显示数据点标记
- `chart.marker_size`: 标记大小
- `chart.line_weight`: 线条粗细（磅）
- `chart.y_axis_min`: Y轴最小值
- `chart.y_axis_max`: Y轴最大值
- `chart.y_axis_major_unit`: Y轴主刻度间隔

### 饼图特有参数
- `chart.data_label_type`: 数据标签类型 (value|percent|both)
- `chart.explosion`: 分离程度 (0-100)

## 使用方法

1. 创建YAML配置文件，放置在 `tmp/{username}/format_files/` 目录下
2. 通过MCP调用 `create_chart_from_yaml` 工具
3. 提供 `filename` 参数指定YAML文件名

## 测试文件

项目中包含以下测试文件：
- `chart_test.yaml`: 柱状图示例
- `line_chart_test.yaml`: 折线图示例
- `pie_chart_test.yaml`: 饼图示例

## 注意事项

1. 数据格式：第一行为标题行，后续为数据行
2. 颜色值：使用RGB十六进制格式，不包含#前缀
3. 位置信息：目前只支持文档末尾插入
4. 主题：目前只支持Office主题
5. 文件路径：YAML文件必须放在 `tmp/{username}/format_files/` 目录下

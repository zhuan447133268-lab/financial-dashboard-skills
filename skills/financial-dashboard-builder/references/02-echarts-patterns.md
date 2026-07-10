# ECharts 常用模式

financial-dashboard-builder 使用的 ECharts 配置规范。

## 通用配置

```javascript
{
    title: { text: '图表标题', left: 'center', textStyle: { fontSize: 14 } },
    tooltip: { trigger: 'axis' },
    legend: { bottom: 0 },
    grid: { left: '3%', right: '4%', bottom: '15%', containLabel: true },
    xAxis: { type: 'category' },
    yAxis: { type: 'value' }
}
```

## 颜色方案

默认使用蓝色系 + 辅助色：
- 主色：#3b82f6（蓝）
- 收入/利润：#10b981（绿）
- 成本/费用：#f59e0b（橙）
- 预警/异常：#dc2626（红）
- 预算/目标：#6b7280（灰）

## 格式化

- 金额：大于 1 万显示为 X.XX 万，大于 1 亿显示为 X.XX 亿
- 百分比：保留 1 位小数
- 数值轴：使用千分位

## 响应式

每个图表容器使用 `width: 100%`，并通过 `window.addEventListener('resize', ...)` 调用 `chart.resize()`。

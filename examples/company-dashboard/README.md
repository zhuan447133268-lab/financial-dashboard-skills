# 公司经营分析看板示例

这是一个端到端示例，演示如何使用 `financial-data-teller` 和 `financial-dashboard-builder` 从原始财务 Excel 生成经营分析看板。

## 文件说明

```
company-dashboard/
├── input/
│   └── 公司经营数据.xlsx          # 示例原始数据（含预算表、实际发生表、公司基础信息）
├── output/
│   ├── 数据资产说明书.md          # financial-data-teller 输出
│   ├── 看板配置.json             # 用户确认后的看板配置
│   └── 经营分析看板.html          # financial-dashboard-builder 输出
└── README.md                      # 本文件
```

## 数据说明

示例数据包含 3 个工作表：

1. **公司基础信息**：公司编码、公司名称、城市、区域、成立日期
2. **预算表**：按年月和公司编码的收入/成本预算
3. **实际发生表**：按年月和公司编码的实际收入/成本

数据中故意埋入了两个典型问题：
- 部分表头存在首尾空格
- 预算成本列中存在"暂估"非数字占位值

## 运行步骤

### 1. 识别数据

```bash
python ../../skills/financial-data-teller/data_teller.py \
  --input ./input \
  --output ./output/数据资产说明书.md
```

输出：
- 数据资产总览
- 每张表的字段清单和语义猜测
- 异常清单
- 可分析维度候选

### 2. 确认口径并配置看板

根据《数据资产说明书》确认：
- 关联字段：`年月` + `公司编码`
- 分析维度：月度趋势、预算达成率、城市对比、主体排行
- 单位：元
- 阈值：达成率 < 80% 红色预警，成本偏差 > 10% 红色预警

这些配置已写入 `output/看板配置.json`。

### 3. 生成看板

```bash
python ../../skills/financial-dashboard-builder/dashboard_builder.py \
  --config ./output/看板配置.json \
  --output ./output/经营分析看板.html
```

生成后用浏览器打开 `经营分析看板.html` 即可查看。

## 预期输出

看板包含：
- 4 个 KPI 卡片：实际收入、预算收入、收入达成率、成本偏差率
- 4 个图表：月度收入/成本趋势、月度预算达成率、城市利润对比、主体利润排行
- 核心发现与建议：基于规则自动触发预警

## 可改进方向

1. 替换为你的真实脱敏数据
2. 根据业务口径调整 `references/03-financial-kpis.md` 中的 KPI 定义
3. 在 `references/02-echarts-patterns.md` 中定制品牌配色
4. 增加更多维度，如"科目结构分析""同比环比"

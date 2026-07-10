# Financial Dashboard Skills

一个面向经营分析场景的 Claude Code Skill 套件：从原始财务 Excel 表到可交互 HTML 看板。

## 解决什么问题

财务人员/业务分析师经常遇到这样的场景：
- 客户丢过来一堆 Excel，表头不规范、单位不统一、合并单元格满天飞
- 需要快速识别这些数据能支撑哪些经营分析维度
- 需要根据确认的需求生成可交互的数据看板

本套件把这件事拆成两个「AI 岗位」：

1. **financial-data-teller**：读取 Excel → 输出《数据资产说明书》
2. **financial-dashboard-builder**：确认维度 → 生成 ECharts HTML 看板

## 核心特点

- **纯 Python**：只需要 `pandas` + `openpyxl` + `jinja2`，无需 MCP server / Node.js
- **可验证**：每个输出都带数据来源、公式、异常报告
- **中文友好**：面向中文财务表头、中文经营分析场景
- **规则驱动**：经营建议基于阈值规则，不接 AI 自由生成
- **可扩展**：按行业/场景增加 references 和模板即可

## 快速开始

```bash
git clone https://github.com/YOUR_USERNAME/financial-dashboard-skills.git
cd financial-dashboard-skills
pip install -r requirements.txt
```

### 第一步：识别数据

```bash
python skills/financial-data-teller/data_teller.py \
  --input examples/company-dashboard/input/ \
  --output examples/company-dashboard/output/数据资产说明书.md
```

### 第二步：生成看板

```bash
python skills/financial-dashboard-builder/dashboard_builder.py \
  --config examples/company-dashboard/output/看板配置.json \
  --output examples/company-dashboard/output/经营分析看板.html
```

### 运行测试

```bash
python -m pytest tests/ -v
```

### 运行评估

```bash
python scripts/run_evals.py
```

## 目录说明

```
skills/
├── financial-data-teller/        # 数据识岗员
│   ├── SKILL.md                  # Skill 本体
│   ├── data_teller.py            # 读取 Excel 脚本
│   ├── references/               # 知识文件
│   └── evals/                    # 评估用例
└── financial-dashboard-builder/  # 看板设计师
    ├── SKILL.md
    ├── dashboard_builder.py
    ├── references/
    └── evals/

examples/company-dashboard/       # 端到端示例（公司经营数据）
scripts/                          # 辅助脚本
```

## 设计哲学

- **一个 Skill 一个岗位**：每个 Skill 只负责一个明确职责
- **先确认、后生成**：不猜测业务口径，关键判断必须停下来让用户确认
- **对外交付前必须校验**：所有数字必须可追溯、可复现

## License

MIT

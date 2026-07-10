"""
financial-dashboard-builder: 根据确认口径生成 ECharts HTML 经营看板。

Usage:
    python dashboard_builder.py --config config.json --output dashboard.html
"""

import argparse
import json
from pathlib import Path

import pandas as pd
from jinja2 import Template


HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>经营分析看板</title>
    <script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
            background: #f5f7fa;
            color: #1f2937;
            padding: 20px;
        }
        .header {
            background: #fff;
            border-radius: 12px;
            padding: 24px;
            margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        .header h1 { font-size: 24px; margin-bottom: 8px; }
        .header .meta { color: #6b7280; font-size: 14px; }
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 16px;
            margin-bottom: 20px;
        }
        .kpi-card {
            background: #fff;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        .kpi-card .label { color: #6b7280; font-size: 14px; margin-bottom: 8px; }
        .kpi-card .value { font-size: 28px; font-weight: 600; color: #111827; }
        .kpi-card .delta { font-size: 13px; margin-top: 6px; }
        .kpi-card .delta.up { color: #dc2626; }
        .kpi-card .delta.down { color: #16a34a; }
        .chart-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(450px, 1fr));
            gap: 20px;
            margin-bottom: 20px;
        }
        .chart-card {
            background: #fff;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        .chart-card h3 {
            font-size: 16px;
            margin-bottom: 12px;
            color: #374151;
        }
        .chart { width: 100%; height: 320px; }
        .insights {
            background: #fff;
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
            margin-bottom: 20px;
        }
        .insights h3 { font-size: 16px; margin-bottom: 12px; }
        .insights ul { list-style: none; }
        .insights li {
            padding: 10px 0;
            border-bottom: 1px solid #f3f4f6;
            font-size: 14px;
            line-height: 1.6;
        }
        .insights li:last-child { border-bottom: none; }
        .insights .level-high { color: #dc2626; }
        .insights .level-medium { color: #d97706; }
        .insights .level-low { color: #16a34a; }
        .footer {
            text-align: center;
            color: #9ca3af;
            font-size: 12px;
            padding: 20px;
        }
        @media (max-width: 768px) {
            .chart-grid { grid-template-columns: 1fr; }
            .kpi-grid { grid-template-columns: 1fr 1fr; }
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>经营分析看板</h1>
        <div class="meta">数据期间: {{ period }} &nbsp;|&nbsp; 单位: {{ unit }} &nbsp;|&nbsp; 生成时间: {{ generated_at }}</div>
    </div>

    <div class="kpi-grid">
        {% for kpi in kpis %}
        <div class="kpi-card">
            <div class="label">{{ kpi.label }}</div>
            <div class="value">{{ kpi.value }}</div>
            {% if kpi.delta %}
            <div class="delta {{ kpi.delta_class }}">{{ kpi.delta }}</div>
            {% endif %}
        </div>
        {% endfor %}
    </div>

    <div class="chart-grid">
        {% for chart in charts %}
        <div class="chart-card">
            <h3>{{ chart.title }}</h3>
            <div id="chart-{{ loop.index }}" class="chart"></div>
        </div>
        {% endfor %}
    </div>

    <div class="insights">
        <h3>核心发现与建议</h3>
        <ul>
            {% for insight in insights %}
            <li class="level-{{ insight.level }}">{{ insight.text }}</li>
            {% endfor %}
        </ul>
    </div>

    <div class="footer">
        本看板由 financial-dashboard-builder 自动生成，关键口径需人工确认。
    </div>

    <script>
        const charts = {{ charts_json | safe }};
        charts.forEach((cfg, idx) => {
            const chart = echarts.init(document.getElementById('chart-' + (idx + 1)));
            chart.setOption(cfg);
            window.addEventListener('resize', () => chart.resize());
        });
    </script>
</body>
</html>
"""


def load_data(data_dir: Path, config: dict) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """读取预算表、实际表、主体信息表。"""
    data_dir = Path(data_dir)
    excel_files = list(data_dir.glob("*.xlsx")) + list(data_dir.glob("*.xls"))
    if not excel_files:
        raise FileNotFoundError(f"未在 {data_dir} 找到 Excel 文件")

    file_path = excel_files[0]
    budget = pd.read_excel(file_path, sheet_name=config["budget_sheet"])
    actual = pd.read_excel(file_path, sheet_name=config["actual_sheet"])
    park = pd.read_excel(file_path, sheet_name=config["park_sheet"])

    # 清理表头空格
    for df in [budget, actual, park]:
        df.columns = [str(c).strip() for c in df.columns]

    return budget, actual, park


def merge_budget_actual(budget: pd.DataFrame, actual: pd.DataFrame, join_keys: list) -> pd.DataFrame:
    """合并预算和实际数据。"""
    budget = budget.copy()
    actual = actual.copy()

    # 清理非数字值
    amount_cols = [c for c in budget.columns if c not in join_keys]
    for col in amount_cols:
        budget[col] = pd.to_numeric(budget[col], errors="coerce")

    amount_cols = [c for c in actual.columns if c not in join_keys]
    for col in amount_cols:
        actual[col] = pd.to_numeric(actual[col], errors="coerce")

    # 重命名避免冲突
    budget_cols = {c: f"预算_{c}" for c in budget.columns if c not in join_keys}
    actual_cols = {c: f"实际_{c}" for c in actual.columns if c not in join_keys}

    budget = budget.rename(columns=budget_cols)
    actual = actual.rename(columns=actual_cols)

    merged = pd.merge(budget, actual, on=join_keys, how="outer")
    return merged


def format_number(val: float, unit: str = "元") -> str:
    """格式化数字显示。"""
    if pd.isna(val):
        return "-"
    if abs(val) >= 1_0000_0000:
        return f"{val/1_0000_0000:.2f}亿"
    if abs(val) >= 1_0000:
        return f"{val/1_0000:.2f}万"
    return f"{val:,.0f}"


def build_kpis(merged: pd.DataFrame, thresholds: dict) -> list[dict]:
    """构建顶部 KPI 卡片。"""
    budget_revenue = merged["预算_预算收入"].sum()
    actual_revenue = merged["实际_实际收入"].sum()
    budget_cost = merged["预算_预算成本"].sum()
    actual_cost = merged["实际_实际成本"].sum()

    revenue_achievement = actual_revenue / budget_revenue if budget_revenue else 0
    cost_variance = (actual_cost - budget_cost) / budget_cost if budget_cost else 0

    def delta_class(ratio):
        if ratio < thresholds.get("achievement_critical", 0.8):
            return "up"  # 红色，未达标
        return "down"

    return [
        {"label": "实际收入", "value": format_number(actual_revenue), "delta": "", "delta_class": ""},
        {"label": "预算收入", "value": format_number(budget_revenue), "delta": "", "delta_class": ""},
        {
            "label": "收入达成率",
            "value": f"{revenue_achievement:.1%}",
            "delta": f"预算 {format_number(budget_revenue)}",
            "delta_class": delta_class(revenue_achievement),
        },
        {
            "label": "成本偏差率",
            "value": f"{cost_variance:.1%}",
            "delta": "实际 vs 预算",
            "delta_class": "up" if cost_variance > 0 else "down",
        },
    ]


def build_monthly_trend(merged: pd.DataFrame) -> dict:
    """月度收入/成本趋势图。"""
    grouped = merged.groupby("年月").agg({
        "预算_预算收入": "sum",
        "实际_实际收入": "sum",
        "预算_预算成本": "sum",
        "实际_实际成本": "sum",
    }).reset_index()
    grouped = grouped.sort_values("年月")

    return {
        "title": {"text": "月度收入/成本趋势", "left": "center", "textStyle": {"fontSize": 14}},
        "tooltip": {"trigger": "axis"},
        "legend": {"data": ["预算收入", "实际收入", "预算成本", "实际成本"], "bottom": 0},
        "grid": {"left": "3%", "right": "4%", "bottom": "15%", "containLabel": True},
        "xAxis": {"type": "category", "data": grouped["年月"].astype(str).tolist()},
        "yAxis": {"type": "value", "axisLabel": {"formatter": "{value}"}},
        "series": [
            {"name": "预算收入", "type": "line", "data": grouped["预算_预算收入"].round(0).tolist(), "smooth": True},
            {"name": "实际收入", "type": "line", "data": grouped["实际_实际收入"].round(0).tolist(), "smooth": True},
            {"name": "预算成本", "type": "line", "data": grouped["预算_预算成本"].round(0).tolist(), "smooth": True, "lineStyle": {"type": "dashed"}},
            {"name": "实际成本", "type": "line", "data": grouped["实际_实际成本"].round(0).tolist(), "smooth": True, "lineStyle": {"type": "dashed"}},
        ],
    }


def build_achievement_chart(merged: pd.DataFrame) -> dict:
    """月度预算达成率图。"""
    grouped = merged.groupby("年月").agg({
        "预算_预算收入": "sum",
        "实际_实际收入": "sum",
    }).reset_index()
    grouped["达成率"] = grouped["实际_实际收入"] / grouped["预算_预算收入"]
    grouped = grouped.sort_values("年月")

    return {
        "title": {"text": "月度收入预算达成率", "left": "center", "textStyle": {"fontSize": 14}},
        "tooltip": {"trigger": "axis", "formatter": "{b}: {c:.1%}"},
        "grid": {"left": "3%", "right": "4%", "bottom": "10%", "containLabel": True},
        "xAxis": {"type": "category", "data": grouped["年月"].astype(str).tolist()},
        "yAxis": {"type": "value", "axisLabel": {"formatter": "{value}%"}, "max": 1.2},
        "series": [{
            "name": "达成率",
            "type": "bar",
            "data": grouped["达成率"].round(4).tolist(),
            "itemStyle": {"color": "#3b82f6"},
            "markLine": {"data": [{"yAxis": 1.0, "name": "目标线", "lineStyle": {"color": "#dc2626"}}]},
        }],
    }


def build_city_comparison(merged: pd.DataFrame, park: pd.DataFrame, entity_key: str) -> dict:
    """城市经营对比图。"""
    joined = pd.merge(merged, park, on=entity_key, how="left")
    grouped = joined.groupby("城市").agg({
        "实际_实际收入": "sum",
        "实际_实际成本": "sum",
        "预算_预算收入": "sum",
    }).reset_index()
    grouped["利润"] = grouped["实际_实际收入"] - grouped["实际_实际成本"]
    grouped = grouped.sort_values("利润", ascending=False)

    return {
        "title": {"text": "城市利润对比", "left": "center", "textStyle": {"fontSize": 14}},
        "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
        "legend": {"data": ["实际收入", "实际成本", "利润"], "bottom": 0},
        "grid": {"left": "3%", "right": "4%", "bottom": "15%", "containLabel": True},
        "xAxis": {"type": "category", "data": grouped["城市"].tolist()},
        "yAxis": {"type": "value"},
        "series": [
            {"name": "实际收入", "type": "bar", "data": grouped["实际_实际收入"].round(0).tolist()},
            {"name": "实际成本", "type": "bar", "data": grouped["实际_实际成本"].round(0).tolist()},
            {"name": "利润", "type": "line", "data": grouped["利润"].round(0).tolist(), "smooth": True},
        ],
    }


def build_park_ranking(merged: pd.DataFrame, entity_key: str) -> dict:
    """主体利润排行图。"""
    grouped = merged.groupby(entity_key).agg({
        "实际_实际收入": "sum",
        "实际_实际成本": "sum",
    }).reset_index()
    grouped["利润"] = grouped["实际_实际收入"] - grouped["实际_实际成本"]
    grouped = grouped.sort_values("利润", ascending=True).tail(10)

    return {
        "title": {"text": "主体利润排行（Top 10）", "left": "center", "textStyle": {"fontSize": 14}},
        "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
        "grid": {"left": "3%", "right": "8%", "bottom": "10%", "containLabel": True},
        "xAxis": {"type": "value"},
        "yAxis": {"type": "category", "data": grouped[entity_key].tolist()},
        "series": [{
            "name": "利润",
            "type": "bar",
            "data": grouped["利润"].round(0).tolist(),
            "itemStyle": {"color": "#10b981"},
        }],
    }


def generate_insights(merged: pd.DataFrame, thresholds: dict) -> list[dict]:
    """基于规则生成经营建议。"""
    insights = []

    # 整体达成率
    total_budget = merged["预算_预算收入"].sum()
    total_actual = merged["实际_实际收入"].sum()
    overall_achievement = total_actual / total_budget if total_budget else 0

    if overall_achievement < thresholds.get("achievement_critical", 0.8):
        insights.append({
            "level": "high",
            "text": f"整体收入达成率仅 {overall_achievement:.1%}，低于警戒线 {thresholds.get('achievement_critical', 0.8):.0%}，需重点关注营收缺口。",
        })
    elif overall_achievement < thresholds.get("achievement_warning", 0.9):
        insights.append({
            "level": "medium",
            "text": f"整体收入达成率 {overall_achievement:.1%}，低于预警线 {thresholds.get('achievement_warning', 0.9):.0%}，建议分析各主体/月份差异。",
        })
    else:
        insights.append({
            "level": "low",
            "text": f"整体收入达成率 {overall_achievement:.1%}，预算完成情况良好。",
        })

    # 月度达成率异常
    monthly = merged.groupby("年月").agg({
        "预算_预算收入": "sum",
        "实际_实际收入": "sum",
    }).reset_index()
    monthly["达成率"] = monthly["实际_实际收入"] / monthly["预算_预算收入"]
    low_months = monthly[monthly["达成率"] < thresholds.get("achievement_critical", 0.8)]
    if not low_months.empty:
        months = ", ".join(low_months["年月"].astype(str).tolist())
        insights.append({
            "level": "high",
            "text": f"月份 {months} 的收入达成率低于警戒线，建议排查对应月份的招生/运营情况。",
        })

    # 成本偏差
    total_budget_cost = merged["预算_预算成本"].sum()
    total_actual_cost = merged["实际_实际成本"].sum()
    cost_variance = (total_actual_cost - total_budget_cost) / total_budget_cost if total_budget_cost else 0
    if cost_variance > thresholds.get("variance_critical", 0.1):
        insights.append({
            "level": "high",
            "text": f"实际成本超出预算 {cost_variance:.1%}，超出容忍阈值 {thresholds.get('variance_critical', 0.1):.0%}，建议核查成本科目。",
        })
    elif cost_variance > thresholds.get("variance_warning", 0.05):
        insights.append({
            "level": "medium",
            "text": f"实际成本超出预算 {cost_variance:.1%}，建议关注主要成本项变动。",
        })

    # 连续下滑检测
    monthly_sorted = monthly.sort_values("年月")
    declines = 0
    for i in range(1, len(monthly_sorted)):
        if monthly_sorted.iloc[i]["实际_实际收入"] < monthly_sorted.iloc[i - 1]["实际_实际收入"]:
            declines += 1
        else:
            declines = 0
        if declines >= 2:
            insights.append({
                "level": "high",
                "text": f"实际收入已连续 {declines + 1} 个月下滑，需立即介入分析原因。",
            })
            break

    return insights


def main():
    parser = argparse.ArgumentParser(description="财务经营看板设计师：生成 ECharts HTML 看板")
    parser.add_argument("--config", "-c", required=True, help="看板配置 JSON 文件路径")
    parser.add_argument("--output", "-o", required=True, help="输出 HTML 看板路径")
    args = parser.parse_args()

    config = json.loads(Path(args.config).read_text(encoding="utf-8"))

    budget, actual, park = load_data(config["data_dir"], config)
    merged = merge_budget_actual(budget, actual, config["join_keys"])

    thresholds = config.get("thresholds", {
        "achievement_warning": 0.9,
        "achievement_critical": 0.8,
        "variance_warning": 0.05,
        "variance_critical": 0.1,
    })

    kpis = build_kpis(merged, thresholds)

    entity_key = config.get("entity_key", "公司编码")

    charts = []
    dimensions = config.get("dimensions", ["月度趋势", "预算达成率", "城市对比", "主体排行"])
    if "月度趋势" in dimensions:
        charts.append(build_monthly_trend(merged))
    if "预算达成率" in dimensions:
        charts.append(build_achievement_chart(merged))
    if "城市对比" in dimensions:
        charts.append(build_city_comparison(merged, park, entity_key))
    if "主体排行" in dimensions or "单园排行" in dimensions:
        charts.append(build_park_ranking(merged, entity_key))

    insights = generate_insights(merged, thresholds)

    from datetime import datetime

    template = Template(HTML_TEMPLATE)
    html = template.render(
        kpis=kpis,
        charts_json=json.dumps(charts, ensure_ascii=False),
        charts=[{"title": c["title"]["text"]} for c in charts],
        insights=insights,
        period=f"{merged['年月'].min()} - {merged['年月'].max()}",
        unit=config.get("unit", "元"),
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
    )

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")
    print(f"看板已生成: {output_path}")


if __name__ == "__main__":
    main()

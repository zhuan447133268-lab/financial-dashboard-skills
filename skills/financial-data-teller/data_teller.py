"""
financial-data-teller: 读取财务 Excel，输出《数据资产说明书》。

Usage:
    python data_teller.py --input ./data/ --output ./report.md
"""

import argparse
import re
from pathlib import Path
from typing import Any

import pandas as pd


def clean_header(header: str) -> str:
    """清理表头：去首尾空格、合并连续空格。"""
    if pd.isna(header):
        return ""
    return re.sub(r"\s+", " ", str(header).strip())


def detect_unit(column_names: list[str], sample_values: pd.Series) -> str:
    """基于表头关键词和数据量级推断单位。"""
    name_text = " ".join(column_names).lower()
    if any(k in name_text for k in ["万元", "万"]):
        return "万元"
    if any(k in name_text for k in ["千元", "千"]):
        return "千元"
    if any(k in name_text for k in ["元"]):
        return "元"

    # 基于数据量级推断
    numeric_values = pd.to_numeric(sample_values, errors="coerce").dropna()
    if not numeric_values.empty:
        max_val = numeric_values.abs().max()
        if max_val >= 1_0000_0000:
            return "疑似万元或更大单位（建议确认）"
        if max_val >= 100_0000:
            return "疑似元或万元（建议确认）"
    return "待确认"


def detect_table_type(column_names: list[str]) -> str:
    """根据表头关键词识别表类型。"""
    text = " ".join(column_names).lower()

    patterns = {
        "科目余额表": ["科目编码", "科目名称", "期初余额", "本期发生", "期末余额"],
        "预算表": ["预算", "年度预算", "预算金额", "预算收入", "预算成本"],
        "实际发生表": ["实际", "发生额", "入账", "实际金额", "实际收入"],
        "组织基础信息表": ["公司", "分公司", "校区", "园区", "门店", "城市", "区域", "成立日期", "主体名称", "主体编码"],
        "部门信息表": ["部门", "部门名称", "部门编码"],
        "时间维度表": ["年份", "月份", "季度", "财年", "年月"],
    }

    scores = {}
    for table_type, keywords in patterns.items():
        scores[table_type] = sum(1 for kw in keywords if kw in text)

    best = max(scores, key=scores.get)
    if scores[best] >= 2:
        return best
    return "未知类型（需要用户确认）"


def detect_field_semantics(column_name: str) -> tuple[str, str]:
    """识别字段语义，返回（语义，置信度）。"""
    name = column_name.lower()

    semantic_patterns = [
        ("时间-年月", ["年月", "月份", "日期", "时间", "年份", "月份"], 0.9),
        ("主体-组织", ["公司", "分公司", "子公司", "校区", "园区", "园所", "门店", "主体"], 0.9),
        ("主体-城市", ["城市", "地区"], 0.9),
        ("主体-区域", ["区域", "大区", "片区"], 0.8),
        # 预算/实际必须优先于收入/成本/费用，因为存在"预算收入"、"实际成本"等复合字段
        ("金额-预算", ["预算", "预算金额"], 0.9),
        ("金额-实际", ["实际", "发生额", "实际金额"], 0.9),
        ("金额-收入", ["收入", "营收", "营业额", "主营业务收入"], 0.8),
        ("金额-成本", ["成本", "主营业务成本"], 0.8),
        ("金额-费用", ["费用", "管理费用", "销售费用", "财务费用"], 0.8),
        ("金额-利润", ["利润", "净利润", "毛利", "营业利润"], 0.8),
        ("科目-编码", ["科目编码", "科目代码"], 0.9),
        ("科目-名称", ["科目名称", "科目"], 0.8),
    ]

    for semantic, keywords, confidence in semantic_patterns:
        if any(kw in name for kw in keywords):
            return semantic, f"{confidence:.0%}"

    return "未识别", "-"


def find_merged_cells(file_path: Path, sheet_name: str) -> list[str]:
    """检测工作表中的合并单元格。"""
    try:
        import openpyxl

        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb[sheet_name]
        merged = []
        for rng in ws.merged_cells.ranges:
            merged.append(f"{rng}")
        return merged
    except Exception:
        return []


def analyze_sheet(file_path: Path, sheet_name: str) -> dict[str, Any]:
    """分析单个工作表。"""
    # 读取前 20 行用于检测表头
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, nrows=20, header=None)

    # 自动检测表头行：找第一个非空行数较多且包含中文/英文关键词的行
    header_row = 0
    for i in range(min(5, len(df_raw))):
        row = df_raw.iloc[i]
        non_empty = row.notna().sum()
        if non_empty >= 2:
            header_row = i
            break

    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

    # 清理表头
    original_columns = list(df.columns)
    df.columns = [clean_header(c) for c in df.columns]

    # 检测异常
    anomalies = []

    # 表头空格异常
    for orig, cleaned in zip(original_columns, df.columns):
        if str(orig) != cleaned and str(orig).strip() == cleaned:
            anomalies.append(f"[高] 表头 '{orig}' 存在首尾空格，已自动清理")

    # 合并单元格
    merged = find_merged_cells(file_path, sheet_name)
    if merged:
        anomalies.append(f"[高] 存在合并单元格: {', '.join(merged[:3])}{' 等' if len(merged) > 3 else ''}")

    # 分析每一列
    fields = []
    for col in df.columns:
        if not col:
            continue
        series = df[col]
        null_rate = series.isna().mean()
        sample_values = series.dropna().head(3).tolist()

        # 尝试转换数字
        numeric_series = pd.to_numeric(series, errors="coerce")
        is_numeric = numeric_series.notna().sum() / len(series) >= 0.5

        dtype = "数值" if is_numeric else "文本"
        semantic, confidence = detect_field_semantics(col)

        if is_numeric and "金额" in semantic:
            unit = detect_unit([col], series)
            if unit in ["待确认", "疑似万元或更大单位（建议确认）", "疑似元或万元（建议确认）"]:
                anomalies.append(f"[中] 字段 '{col}' 单位{unit}")

        # 检查非数字异常
        if is_numeric and (series.astype(str).isin(["-", "", "暂估", "None"]).any()):
            anomalies.append(f"[中] 字段 '{col}' 存在非数字占位值")

        fields.append(
            {
                "字段名": col,
                "类型": dtype,
                "空值率": f"{null_rate:.1%}",
                "示例值": str(sample_values[:3]),
                "语义猜测": semantic,
                "置信度": confidence,
            }
        )

    table_type = detect_table_type([str(c) for c in df.columns])

    return {
        "工作表名": sheet_name,
        "表头行": header_row + 1,
        "列数": len(df.columns),
        "行数": len(df),
        "识别表类型": table_type,
        "字段清单": fields,
        "异常": anomalies,
    }


def analyze_file(file_path: Path) -> dict[str, Any]:
    """分析单个 Excel 文件。"""
    xl = pd.ExcelFile(file_path)
    sheets = []
    for sheet_name in xl.sheet_names:
        try:
            sheets.append(analyze_sheet(file_path, sheet_name))
        except Exception as e:
            sheets.append(
                {
                    "工作表名": sheet_name,
                    "错误": f"读取失败: {e}",
                }
            )

    return {
        "文件名": file_path.name,
        "工作表数": len(xl.sheet_names),
        "工作表": sheets,
    }


def generate_candidate_dimensions(files_analysis: list[dict]) -> list[dict]:
    """基于分析结果生成候选看板维度。"""
    candidates = []

    # 收集所有字段语义
    all_semantics = set()
    has_budget = False
    has_actual = False
    has_time = False
    has_subject = False
    has_region = False
    has_park = False

    for file in files_analysis:
        for sheet in file.get("工作表", []):
            for field in sheet.get("字段清单", []):
                semantic = field.get("语义猜测", "")
                all_semantics.add(semantic)
                if semantic == "金额-预算":
                    has_budget = True
                if semantic == "金额-实际":
                    has_actual = True
                if semantic.startswith("时间"):
                    has_time = True
                if semantic.startswith("科目"):
                    has_subject = True
                if semantic == "主体-区域":
                    has_region = True
                if semantic == "主体-组织":
                    has_park = True

    if has_time and any(s.startswith("金额-") for s in all_semantics):
        candidates.append(
            {
                "维度名称": "月度营收/成本趋势",
                "所需字段": "时间字段 + 金额字段（收入/成本/费用）",
                "前提条件": "确认收入/成本口径和科目范围",
                "风险等级": "低",
            }
        )

    if has_budget and has_actual and has_time:
        candidates.append(
            {
                "维度名称": "预算达成率",
                "所需字段": "预算金额 + 实际金额 + 时间字段",
                "前提条件": "确认预算表和实际表的关联字段（主体/月份/科目）",
                "风险等级": "中",
            }
        )

    if has_region and any(s.startswith("金额-") for s in all_semantics):
        candidates.append(
            {
                "维度名称": "区域/城市经营对比",
                "所需字段": "区域/城市字段 + 金额字段",
                "前提条件": "确认区域层级划分",
                "风险等级": "低",
            }
        )

    if has_park and any(s.startswith("金额-") for s in all_semantics):
        candidates.append(
            {
                "维度名称": "单主体经营画像",
                "所需字段": "主体字段 + 金额字段",
                "前提条件": "确认主体口径和关园/关停识别规则",
                "风险等级": "中",
            }
        )

    if has_subject:
        candidates.append(
            {
                "维度名称": "科目结构分析",
                "所需字段": "科目编码/名称 + 金额字段",
                "前提条件": "确认科目层级（一级/二级）",
                "风险等级": "低",
            }
        )

    if not candidates:
        candidates.append(
            {
                "维度名称": "通用数据透视",
                "所需字段": "待识别",
                "前提条件": "需要用户补充业务说明",
                "风险等级": "高",
            }
        )

    return candidates


def build_report(files_analysis: list[dict]) -> str:
    """生成 Markdown 报告。"""
    lines = ["# 数据资产说明书\n"]

    # 总览
    lines.append("## 数据资产总览\n")
    lines.append("| 文件名 | 工作表数 | 识别到的表类型 | 异常数 |")
    lines.append("|---|---|---|---|")
    for file in files_analysis:
        sheets = file.get("工作表", [])
        types = ", ".join(set(s.get("识别表类型", "未知") for s in sheets if "识别表类型" in s))
        anomaly_count = sum(len(s.get("异常", [])) for s in sheets)
        lines.append(f"| {file['文件名']} | {file['工作表数']} | {types} | {anomaly_count} |")
    lines.append("")

    # 每张表详情
    for file in files_analysis:
        lines.append(f"## {file['文件名']}\n")
        for sheet in file.get("工作表", []):
            if "错误" in sheet:
                lines.append(f"### 工作表: {sheet['工作表名']}")
                lines.append(f"⚠️ {sheet['错误']}\n")
                continue

            lines.append(f"### 工作表: {sheet['工作表名']}")
            lines.append(f"- 表头行: 第 {sheet['表头行']} 行")
            lines.append(f"- 维度: {sheet['列数']} 列 × {sheet['行数']} 行")
            lines.append(f"- 识别表类型: **{sheet['识别表类型']}**\n")

            lines.append("#### 字段清单")
            lines.append("| 字段名 | 类型 | 空值率 | 示例值 | 语义猜测 | 置信度 |")
            lines.append("|---|---|---|---|---|---|")
            for field in sheet["字段清单"]:
                lines.append(
                    f"| {field['字段名']} | {field['类型']} | {field['空值率']} | "
                    f"{field['示例值']} | {field['语义猜测']} | {field['置信度']} |"
                )
            lines.append("")

    # 异常清单
    lines.append("## 异常清单\n")
    has_anomaly = False
    for file in files_analysis:
        for sheet in file.get("工作表", []):
            if "异常" in sheet and sheet["异常"]:
                has_anomaly = True
                lines.append(f"### {file['文件名']} / {sheet['工作表名']}")
                for anomaly in sheet["异常"]:
                    lines.append(f"- {anomaly}")
                lines.append("")
    if not has_anomaly:
        lines.append("✅ 未检测到明显异常。\n")

    # 候选维度
    candidates = generate_candidate_dimensions(files_analysis)
    lines.append("## 可分析维度候选\n")
    lines.append("| 维度名称 | 所需字段 | 前提条件 | 风险等级 |")
    lines.append("|---|---|---|---|")
    for c in candidates:
        lines.append(f"| {c['维度名称']} | {c['所需字段']} | {c['前提条件']} | {c['风险等级']} |")
    lines.append("")

    lines.append("---\n")
    lines.append("*本报告由 financial-data-teller 自动生成，所有语义猜测仅供参考，关键口径需人工确认。*")

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="财务数据识岗员：读取 Excel 生成数据资产说明书")
    parser.add_argument("--input", "-i", required=True, help="输入 Excel 文件或目录")
    parser.add_argument("--output", "-o", required=True, help="输出 Markdown 报告路径")
    args = parser.parse_args()

    input_path = Path(args.input)
    if input_path.is_dir():
        excel_files = sorted(input_path.glob("*.xlsx")) + sorted(input_path.glob("*.xls"))
    else:
        excel_files = [input_path]

    if not excel_files:
        print(f"未在 {input_path} 找到 Excel 文件")
        return

    files_analysis = []
    for file_path in excel_files:
        print(f"分析中: {file_path.name}")
        files_analysis.append(analyze_file(file_path))

    report = build_report(files_analysis)

    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(report, encoding="utf-8")
    print(f"报告已生成: {output_path}")


if __name__ == "__main__":
    main()

"""
检查生成的 Excel 文件或 Python 脚本是否存在常见质量问题。

Usage:
    python scripts/excel_lint.py --file output.xlsx
    python scripts/excel_lint.py --code script.py
"""

import argparse
import re
from pathlib import Path


def lint_code(code_path: Path) -> list[str]:
    """检查 openpyxl/pandas Excel 代码常见坑。"""
    code = code_path.read_text(encoding="utf-8")
    issues = []

    if "data_only=True" in code and "save" in code:
        issues.append("[XL001] 使用 data_only=True 加载后又 save，会丢失公式")

    if re.search(r"ws\[.+\]\s*=\s*['\"].*=", code):
        issues.append("[XL002] 公式赋值给单元格时缺少前导等号检查")

    if "R1C1" in code:
        issues.append("[XL003] 建议使用 A1 而非 R1C1 公式风格")

    if re.search(r"to_excel.*\.xlsx", code) and "engine" not in code:
        issues.append("[XL004] 建议显式指定 to_excel 的 engine 参数")

    return issues


def main():
    parser = argparse.ArgumentParser(description="Excel 代码/文件质量检查")
    parser.add_argument("--file", help="要检查的 .xlsx 文件")
    parser.add_argument("--code", help="要检查的 Python 脚本")
    args = parser.parse_args()

    if args.code:
        issues = lint_code(Path(args.code))
        if issues:
            print("\n".join(issues))
        else:
            print("未发现问题")

    if args.file:
        print("文件检查功能待实现")


if __name__ == "__main__":
    main()

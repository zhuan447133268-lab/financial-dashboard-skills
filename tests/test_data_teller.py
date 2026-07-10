import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "skills" / "financial-data-teller"))
from data_teller import analyze_file, clean_header, detect_field_semantics, detect_table_type


def test_clean_header():
    assert clean_header("  收入  ") == "收入"
    assert clean_header("预算  收入") == "预算 收入"


def test_detect_table_type():
    assert detect_table_type(["科目编码", "科目名称", "期初余额"]) == "科目余额表"
    assert detect_table_type(["公司编码", "公司名称", "城市"]) == "组织基础信息表"
    assert detect_table_type(["校区编码", "校区名称", "城市"]) == "组织基础信息表"


def test_detect_field_semantics():
    assert detect_field_semantics("预算收入")[0] == "金额-预算"
    assert detect_field_semantics("实际成本")[0] == "金额-实际"
    assert detect_field_semantics("年月")[0] == "时间-年月"


def test_analyze_file(tmp_path):
    # 创建测试 Excel
    df = pd.DataFrame({
        " 年月 ": [202501, 202502],
        "公司编码": ["C001", "C001"],
        "实际收入": [100000, 120000],
    })
    file_path = tmp_path / "test.xlsx"
    df.to_excel(file_path, index=False)

    result = analyze_file(file_path)
    assert result["文件名"] == "test.xlsx"
    assert result["工作表数"] == 1
    assert result["工作表"][0]["工作表名"] == "Sheet1"

    # 检查表头空格异常
    anomalies = result["工作表"][0]["异常"]
    assert any("表头" in a and "空格" in a for a in anomalies)

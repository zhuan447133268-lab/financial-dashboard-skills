import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "skills" / "financial-dashboard-builder"))
from dashboard_builder import load_data, merge_budget_actual, build_kpis


def test_merge_budget_actual():
    budget = pd.DataFrame({
        "年月": [202501, 202502],
        "公司编码": ["C001", "C001"],
        "预算收入": [100000, 120000],
    })
    actual = pd.DataFrame({
        "年月": [202501, 202502],
        "公司编码": ["C001", "C001"],
        "实际收入": [90000, 130000],
    })

    merged = merge_budget_actual(budget, actual, ["年月", "公司编码"])
    assert "预算_预算收入" in merged.columns
    assert "实际_实际收入" in merged.columns
    assert len(merged) == 2


def test_build_kpis():
    merged = pd.DataFrame({
        "预算_预算收入": [100000, 120000],
        "实际_实际收入": [90000, 130000],
        "预算_预算成本": [60000, 70000],
        "实际_实际成本": [65000, 72000],
    })
    thresholds = {
        "achievement_warning": 0.9,
        "achievement_critical": 0.8,
    }
    kpis = build_kpis(merged, thresholds)
    assert len(kpis) == 4
    assert any(k["label"] == "实际收入" for k in kpis)
    assert any(k["label"] == "收入达成率" for k in kpis)

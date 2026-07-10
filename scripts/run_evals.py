"""
运行所有 Skill 的 evals，验证输出是否符合预期。

Usage:
    python scripts/run_evals.py
"""

import json
import subprocess
import sys
from pathlib import Path


def run_eval(eval_path: Path) -> dict:
    """运行单个 eval 文件并返回结果。"""
    results = []
    evals = json.loads(eval_path.read_text(encoding="utf-8"))

    for case in evals:
        results.append({
            "name": case["name"],
            "status": "manual",
            "expected_outputs": case.get("expected_outputs", []),
            "note": "需人工检查或补充自动化断言",
        })

    return {
        "skill": eval_path.parent.parent.name,
        "eval_file": eval_path.name,
        "results": results,
    }


def main():
    repo_root = Path(__file__).parent.parent
    eval_files = sorted(repo_root.rglob("evals/evals.json"))

    if not eval_files:
        print("未找到 evals.json 文件")
        return

    all_results = []
    for eval_path in eval_files:
        print(f"运行: {eval_path.relative_to(repo_root)}")
        all_results.append(run_eval(eval_path))

    summary_path = repo_root / "evals_report.json"
    summary_path.write_text(json.dumps(all_results, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n评估报告已生成: {summary_path}")

    # 简单统计
    total = sum(len(r["results"]) for r in all_results)
    print(f"共 {len(all_results)} 个 Skill，{total} 个 eval 用例")


if __name__ == "__main__":
    main()

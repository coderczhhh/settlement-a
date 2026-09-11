from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from a.a_utils import progress, read_json, status, write_json


def a_5_build_final_table(
    market_summary: dict[str, Any],
    user_volume_summary: dict[str, Any],
    output_path: str | Path | None = None,
) -> dict[str, Any]:
    """把左侧分组汇总与右侧两组用户电量合成统一 JSON 总表。"""
    status("A5", "开始组装最终统一数据结构")
    sections: dict[str, Any] = {}
    items = [
        ("市场主体类型与电压等级汇总", market_summary["rows"]),
        ("不含永强", user_volume_summary["without_yongqiang"]),
        ("含永强", user_volume_summary["with_yongqiang"]),
    ]
    for name, rows in progress(items, desc="A5 组装结果", unit="section"):
        sections[name] = rows

    result = {
        "schema_version": "1.0",
        "format": "JSON object containing lists of row dictionaries",
        "units": {"售电电量": "千瓦时", "电能电费": "元", "用电量": "千瓦时"},
        "sections": sections,
        "audit": {
            "market_summary": market_summary["audit"],
            "user_volume_summary": user_volume_summary["audit"],
            "status": "PASS"
            if market_summary["audit"]["status"] == "PASS"
            and user_volume_summary["audit"]["status"] == "PASS"
            else "FAIL",
        },
    }
    if result["audit"]["status"] != "PASS":
        raise AssertionError("最终总表存在未通过的勾稽检查")

    if output_path is not None:
        output = write_json(result, output_path)
        status("A5", f"已输出最终总表：{output}")
    status("A5", "最终总表组装完成")
    return result


def _main() -> None:
    parser = argparse.ArgumentParser(description="合并 A3 与 A4 结果为最终 JSON 总表")
    parser.add_argument("--market-summary", required=True)
    parser.add_argument("--user-volume-summary", required=True)
    parser.add_argument("--output", default="outputs/a_group/a_5_final_table.json")
    args = parser.parse_args()
    a_5_build_final_table(
        read_json(args.market_summary),
        read_json(args.user_volume_summary),
        args.output,
    )


if __name__ == "__main__":
    _main()

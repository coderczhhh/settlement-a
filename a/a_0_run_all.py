from __future__ import annotations

import argparse
import platform
import sys
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from time import perf_counter

from a.a_1_read_excel_to_dict import a_1_read_excel_to_dict
from a.a_2_standardize_records import a_2_standardize_records
from a.a_3_aggregate_market_subjects import a_3_aggregate_market_subjects
from a.a_4_build_user_totals import a_4_build_user_totals
from a.a_5_build_final_table import a_5_build_final_table
from a.a_6_export_json_to_excel import a_6_export_json_to_excel
from a.a_utils import status, value_by_label, write_json


def a_0_run_all(
    user_file: str | Path,
    yongqiang_file: str | Path,
    output_dir: str | Path,
    pumped_storage_kwh: Decimal | int | float | str = Decimal("0"),
) -> dict:
    """串行执行 A1-A6，保留JSON中间结果、Excel核对表及运行清单。"""
    started = perf_counter()
    started_at = datetime.now().astimezone()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    user_file = Path(user_file)
    yongqiang_file = Path(yongqiang_file)

    status("A0", f"流程开始，输出目录：{output_dir.resolve()}")
    raw_data = a_1_read_excel_to_dict(
        user_file,
        yongqiang_file,
        output_dir / "a_1_all_workbooks.json",
    )
    standardized = a_2_standardize_records(
        raw_data,
        output_dir / "a_2_standardized_records.json",
    )
    market_summary = a_3_aggregate_market_subjects(
        standardized,
        output_dir / "a_3_market_subject_voltage_summary.json",
    )
    user_volume_summary = a_4_build_user_totals(
        standardized,
        pumped_storage_kwh,
        output_dir / "a_4_user_volume_summary.json",
    )
    final_table = a_5_build_final_table(
        market_summary,
        user_volume_summary,
        output_dir / "a_5_final_table.json",
    )
    excel_output = a_6_export_json_to_excel(
        final_table,
        standardized,
        output_dir / "a_6_statistics_comparison.xlsx",
    )

    output_files = sorted(output_dir.glob("a_[1-6]_*"))
    without_total = value_by_label(user_volume_summary["without_yongqiang"], "总计")
    with_total = value_by_label(user_volume_summary["with_yongqiang"], "总计")
    with_pumped = value_by_label(user_volume_summary["with_yongqiang"], "合计（含抽水蓄能）")
    manifest = {
        "status": "PASS",
        "started_at": started_at.isoformat(),
        "finished_at": datetime.now().astimezone().isoformat(),
        "elapsed_seconds": round(perf_counter() - started, 3),
        "runtime": {
            "python": sys.version,
            "platform": platform.platform(),
        },
        "inputs": [
            {"path": str(user_file.resolve())},
            {"path": str(yongqiang_file.resolve())},
        ],
        "parameters": {"pumped_storage_kwh": pumped_storage_kwh},
        "record_counts": standardized["statistics"],
        "key_results": {
            "without_yongqiang_total": without_total,
            "with_yongqiang_total": with_total,
            "with_pumped_storage_total": with_pumped,
        },
        "outputs": [
            {
                "path": str(path.resolve()),
                "size_bytes": path.stat().st_size,
            }
            for path in output_files
        ],
        "final_audit_status": final_table["audit"]["status"],
    }
    manifest_path = write_json(manifest, output_dir / "a_0_run_manifest.json")
    status("A0", f"全流程完成，不含永强总计 {without_total}，含永强总计 {with_total}")
    status("A0", f"运行清单：{manifest_path}")
    return {
        "manifest": manifest,
        "final_table": final_table,
        "excel_output": excel_output,
    }


def _main() -> None:
    parser = argparse.ArgumentParser(description="运行 A 组用户侧分类汇总全流程")
    parser.add_argument("--user-file", required=True, help="用户侧清算数据 Excel")
    parser.add_argument("--yongqiang-file", required=True, help="永强供电公司月度数据 Excel")
    parser.add_argument("--output-dir", default="outputs/a_group")
    parser.add_argument(
        "--pumped-storage-kwh",
        default="0",
        help="抽水蓄能用电量（千瓦时）；两个测试文件不包含该来源，默认 0",
    )
    args = parser.parse_args()
    a_0_run_all(
        args.user_file,
        args.yongqiang_file,
        args.output_dir,
        args.pumped_storage_kwh,
    )


if __name__ == "__main__":
    _main()

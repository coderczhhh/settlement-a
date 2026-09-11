from __future__ import annotations

import argparse
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from typing import Any

from a.a_utils import clean_text, progress, read_json, status, to_decimal, write_json


MARKET_ORDER = ["-", "1批发市场用户", "2兜底用户", "3零售用户", "4代理用户"]

RESIDENT_VOLTAGE_ORDER = [
    "1,不满1千伏",
    "1,不满1千伏年用电2760千瓦时及以下部分",
    "1,不满1千伏年用电2761-4800千瓦时部分",
    "1,不满1千伏年用电4801千瓦时及以上部分",
    "2,1-10千伏",
    "2,1-10千伏年用电2760千瓦时及以下部分",
    "2,1-10千伏年用电2761-4800千瓦时部分",
    "2,1-10千伏年用电4801千瓦时及以上部分",
    "3,20千伏",
    "4,35千伏",
    "5,110千伏",
    "6,220千伏及以上",
]

MARKET_VOLTAGE_ORDER = [
    "1,不满1千伏",
    "2,1-10(20)千伏",
    "4,35千伏",
    "5,110千伏",
    "6,220千伏及以上",
]


def a_3_aggregate_market_subjects(
    standardized: dict[str, Any], output_path: str | Path | None = None
) -> dict[str, Any]:
    """按市场主体类型、电压等级汇总售电电量与电能电费。"""
    status("A3", "开始按市场主体类型和电压等级汇总")
    grouped: dict[tuple[str, str], list[Decimal]] = defaultdict(
        lambda: [Decimal("0"), Decimal("0")]
    )
    records = standardized["records"]["user_sales"]

    for record in progress(records, desc="A3 分组汇总", unit="row"):
        market_type = clean_text(record["市场主体类型"])
        voltage = clean_text(record["电压等级"])
        if market_type not in MARKET_ORDER:
            raise ValueError(f"未识别市场主体类型：{market_type!r}，来源 {record['来源']}")
        allowed_voltages = RESIDENT_VOLTAGE_ORDER if market_type == "-" else MARKET_VOLTAGE_ORDER
        if voltage not in allowed_voltages:
            raise ValueError(
                f"市场主体 {market_type!r} 出现未配置电压等级 {voltage!r}，来源 {record['来源']}"
            )
        grouped[(market_type, voltage)][0] += to_decimal(
            record["售电电量"], field="售电电量", source=str(record["来源"])
        )
        grouped[(market_type, voltage)][1] += to_decimal(
            record["电能电费"], field="电能电费", source=str(record["来源"])
        )

    rows: list[dict[str, Any]] = []
    for market_type in MARKET_ORDER:
        voltage_order = RESIDENT_VOLTAGE_ORDER if market_type == "-" else MARKET_VOLTAGE_ORDER
        subtotal_volume = sum((grouped[(market_type, v)][0] for v in voltage_order), Decimal("0"))
        subtotal_fee = sum((grouped[(market_type, v)][1] for v in voltage_order), Decimal("0"))
        rows.append(
            {
                "市场主体类型": market_type,
                "电压等级": "",
                "求和项:售电电量": subtotal_volume,
                "求和项:电能电费": subtotal_fee,
            }
        )
        for voltage in voltage_order:
            rows.append(
                {
                    "市场主体类型": market_type,
                    "电压等级": voltage,
                    "求和项:售电电量": grouped[(market_type, voltage)][0],
                    "求和项:电能电费": grouped[(market_type, voltage)][1],
                }
            )

    grand_volume = sum((item["求和项:售电电量"] for item in rows if item["电压等级"] == ""), Decimal("0"))
    grand_fee = sum((item["求和项:电能电费"] for item in rows if item["电压等级"] == ""), Decimal("0"))
    rows.append(
        {
            "市场主体类型": "总计",
            "电压等级": "",
            "求和项:售电电量": grand_volume,
            "求和项:电能电费": grand_fee,
        }
    )

    source_volume = sum(
        (to_decimal(item["售电电量"], field="售电电量", source="user_sales") for item in records),
        Decimal("0"),
    )
    source_fee = sum(
        (to_decimal(item["电能电费"], field="电能电费", source="user_sales") for item in records),
        Decimal("0"),
    )
    audit = {
        "source_volume": source_volume,
        "summary_volume": grand_volume,
        "volume_difference": grand_volume - source_volume,
        "source_fee": source_fee,
        "summary_fee": grand_fee,
        "fee_difference": grand_fee - source_fee,
        "status": "PASS" if grand_volume == source_volume and grand_fee == source_fee else "FAIL",
    }
    if audit["status"] != "PASS":
        raise AssertionError(f"A3 汇总与来源不平：{audit}")

    result = {"unit": {"售电电量": "千瓦时", "电能电费": "元"}, "rows": rows, "audit": audit}
    if output_path is not None:
        output = write_json(result, output_path)
        status("A3", f"已输出市场主体分压汇总：{output}")
    status("A3", f"汇总完成，售电电量总计 {grand_volume}，电能电费总计 {grand_fee}")
    return result


def _main() -> None:
    parser = argparse.ArgumentParser(description="按市场主体类型、电压等级汇总")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", default="outputs/a_group/a_3_market_subject_voltage_summary.json")
    args = parser.parse_args()
    a_3_aggregate_market_subjects(read_json(args.input), args.output)


if __name__ == "__main__":
    _main()

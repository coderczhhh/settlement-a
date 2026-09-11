from __future__ import annotations

import argparse
import re
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable

from a.a_utils import clean_text, progress, read_json, status, to_decimal, value_by_label, write_json


def _voltage_code(voltage: str, source: Any) -> int:
    match = re.match(r"^(\d+),", clean_text(voltage))
    if not match:
        raise ValueError(f"无法识别电压编码：{voltage!r}，来源 {source}")
    return int(match.group(1))


def _proxy_voltage_band(voltage: str, source: Any) -> str:
    normalized = clean_text(voltage).upper().replace(" ", "")
    if "380V" in normalized or "0.38KV" in normalized or normalized.endswith("220V"):
        return "low"
    high_markers = ("1-10", "20KV", "35KV", "110KV", "220KV")
    if any(marker in normalized for marker in high_markers):
        return "high"
    raise ValueError(f"1.5倍代理购电出现未识别电压：{voltage!r}，来源 {source}")


def _sum_records(records: Iterable[dict[str, Any]]) -> Decimal:
    return sum(
        (
            to_decimal(record["售电电量"], field="售电电量", source=str(record["来源"]))
            for record in records
        ),
        Decimal("0"),
    )


def _row(label: str, value: Decimal) -> dict[str, Any]:
    return {"用户类型": label, "用电量": value}


def _validate_yongqiang_record(record: dict[str, Any]) -> str:
    market_type = clean_text(record["市场主体类型"])
    use_category = clean_text(record["用电分类"])
    if use_category == "五、打水用电":
        raise ValueError(f"永强打水用电不应进入有效明细：{record}")
    if not market_type:
        if not use_category.startswith(("一、", "四、")):
            raise ValueError(f"永强空市场主体行不属于居民农业：{record}")
        return "resident"
    if market_type == "代理用户":
        return "proxy"
    if market_type == "零售用户":
        return "market"
    raise ValueError(f"永强数据出现未识别市场主体类型：{market_type!r}")


def a_4_build_user_totals(
    standardized: dict[str, Any],
    pumped_storage_kwh: Decimal | int | float | str = Decimal("0"),
    output_path: str | Path | None = None,
) -> dict[str, Any]:
    """生成不含永强、含永强两组用户侧电量汇总。"""
    status("A4", "开始计算不含永强与含永强汇总")
    user_sales = standardized["records"]["user_sales"]
    proxy_15x = standardized["records"]["proxy_15x"]
    yongqiang = standardized["records"]["yongqiang"]

    buckets = {
        "resident_low": Decimal("0"),
        "resident_mid": Decimal("0"),
        "resident_high": Decimal("0"),
        "proxy_low": Decimal("0"),
        "proxy_high": Decimal("0"),
        "fallback_low": Decimal("0"),
        "fallback_high": Decimal("0"),
        "market_low": Decimal("0"),
        "market_high": Decimal("0"),
    }

    for record in progress(user_sales, desc="A4 归集国网用户电量", unit="row"):
        market_type = clean_text(record["市场主体类型"])
        code = _voltage_code(record["电压等级"], record["来源"])
        volume = to_decimal(record["售电电量"], field="售电电量", source=str(record["来源"]))
        if market_type == "-":
            band = "low" if code == 1 else "mid" if code in (2, 3) else "high"
            buckets[f"resident_{band}"] += volume
        elif market_type == "4代理用户":
            buckets["proxy_low" if code == 1 else "proxy_high"] += volume
        elif market_type == "2兜底用户":
            buckets["fallback_low" if code == 1 else "fallback_high"] += volume
        elif market_type in ("1批发市场用户", "3零售用户"):
            buckets["market_low" if code == 1 else "market_high"] += volume
        else:
            raise ValueError(f"未识别市场主体类型：{market_type!r}")

    proxy_15x_total = Decimal("0")
    for record in progress(proxy_15x, desc="A4 并入 1.5 倍代理购电", unit="row"):
        volume = to_decimal(record["售电电量"], field="售电电量", source=str(record["来源"]))
        band = _proxy_voltage_band(record["电压等级"], record["来源"])
        buckets[f"proxy_{band}"] += volume
        proxy_15x_total += volume

    resident_total = buckets["resident_low"] + buckets["resident_mid"] + buckets["resident_high"]
    proxy_total = buckets["proxy_low"] + buckets["proxy_high"]
    fallback_total = buckets["fallback_low"] + buckets["fallback_high"]
    market_total = buckets["market_low"] + buckets["market_high"]
    industrial_low = buckets["proxy_low"] + buckets["fallback_low"] + buckets["market_low"]
    industrial_high = buckets["proxy_high"] + buckets["fallback_high"] + buckets["market_high"]
    industrial_total = industrial_low + industrial_high
    without_total = resident_total + industrial_total

    without_yongqiang = [
        _row("1.居民、农业电量", resident_total),
        _row("其中：1kV及以下", buckets["resident_low"]),
        _row("1-10（20）kV", buckets["resident_mid"]),
        _row("35kV及以上", buckets["resident_high"]),
        _row("2.工商业用户电量小计", industrial_total),
        _row("其中：不满1千伏电量", industrial_low),
        _row(" 1-10千伏及以上电量", industrial_high),
        _row("代理购电用户", proxy_total),
        _row("其中：代理购电不满1千伏电量", buckets["proxy_low"]),
        _row("代理1-10千伏及以上", buckets["proxy_high"]),
        _row("以代理价格收费的兜底用户", fallback_total),
        _row("其中：兜底用户不满1千伏电量", buckets["fallback_low"]),
        _row("兜底1-10千伏及以上", buckets["fallback_high"]),
        _row("市场化用户", market_total),
        _row("其中：市场用户不满1千伏电量", buckets["market_low"]),
        _row("市场1-10千伏及以上", buckets["market_high"]),
        _row("总计", without_total),
    ]

    yq = {
        "resident": Decimal("0"),
        "proxy_low": Decimal("0"),
        "proxy_high": Decimal("0"),
        "market_low": Decimal("0"),
        "market_high": Decimal("0"),
    }
    for record in progress(yongqiang, desc="A4 归集永强电量", unit="row"):
        category = _validate_yongqiang_record(record)
        volume = to_decimal(record["售电电量"], field="售电电量", source=str(record["来源"]))
        if category == "resident":
            yq["resident"] += volume
        else:
            code = _voltage_code(record["电压等级"], record["来源"])
            yq[f"{category}_{'low' if code == 1 else 'high'}"] += volume

    with_resident = resident_total + yq["resident"]
    with_proxy_low = buckets["proxy_low"] + yq["proxy_low"]
    with_proxy_high = buckets["proxy_high"] + yq["proxy_high"]
    with_market_low = buckets["market_low"] + yq["market_low"]
    with_market_high = buckets["market_high"] + yq["market_high"]
    with_fallback_low = buckets["fallback_low"]
    with_fallback_high = buckets["fallback_high"]
    with_proxy = with_proxy_low + with_proxy_high
    with_fallback = with_fallback_low + with_fallback_high
    with_market = with_market_low + with_market_high
    with_industrial_low = with_proxy_low + with_fallback_low + with_market_low
    with_industrial_high = with_proxy_high + with_fallback_high + with_market_high
    with_industrial = with_industrial_low + with_industrial_high
    with_total = with_resident + with_industrial
    pumped = to_decimal(pumped_storage_kwh, field="抽水蓄能用电量", source="CLI")

    with_yongqiang = [
        _row("1.居民、农业电量", with_resident),
        _row("2.工商业用户电量小计", with_industrial),
        _row("其中：不满1千伏电量", with_industrial_low),
        _row(" 1-10千伏及以上电量", with_industrial_high),
        _row("代理购电用户", with_proxy),
        _row("其中：代购电不满1千伏电量", with_proxy_low),
        _row("代理1-10千伏及以上", with_proxy_high),
        _row("以代理价格收费的兜底用户", with_fallback),
        _row("其中：兜底用户不满1千伏电量", with_fallback_low),
        _row("兜底1-10千伏及以上", with_fallback_high),
        _row("市场化用户", with_market),
        _row("其中：市场用户不满1千伏电量", with_market_low),
        _row("市场1-10千伏及以上", with_market_high),
        _row("总计", with_total),
        _row("合计（含抽水蓄能）", with_total + pumped),
    ]

    yq_detail_total = _sum_records(yongqiang)
    declared_total_raw = standardized.get("statistics", {}).get("yongqiang_declared_total")
    declared_total = (
        None
        if declared_total_raw is None
        else to_decimal(declared_total_raw, field="永强声明合计", source="standardized.statistics")
    )
    checks = {
        "without_industrial_check": value_by_label(without_yongqiang, "2.工商业用户电量小计")
        - value_by_label(without_yongqiang, "代理购电用户")
        - value_by_label(without_yongqiang, "以代理价格收费的兜底用户")
        - value_by_label(without_yongqiang, "市场化用户"),
        "without_total_check": value_by_label(without_yongqiang, "总计")
        - value_by_label(without_yongqiang, "1.居民、农业电量")
        - value_by_label(without_yongqiang, "2.工商业用户电量小计"),
        "with_industrial_check": value_by_label(with_yongqiang, "2.工商业用户电量小计")
        - value_by_label(with_yongqiang, "代理购电用户")
        - value_by_label(with_yongqiang, "以代理价格收费的兜底用户")
        - value_by_label(with_yongqiang, "市场化用户"),
        "with_total_check": value_by_label(with_yongqiang, "总计")
        - value_by_label(with_yongqiang, "1.居民、农业电量")
        - value_by_label(with_yongqiang, "2.工商业用户电量小计"),
        "yongqiang_increment_check": with_total - without_total - yq_detail_total,
        "pumped_storage_check": value_by_label(with_yongqiang, "合计（含抽水蓄能）") - with_total - pumped,
    }
    if any(value != 0 for value in checks.values()):
        raise AssertionError(f"A4 勾稽失败：{checks}")

    declared_difference = None if declared_total is None else yq_detail_total - declared_total
    result = {
        "unit": "千瓦时",
        "assumptions": {
            "pumped_storage_kwh": pumped,
            "proxy_15x_is_added_to_proxy_volume": True,
            "yongqiang_Sheet3_is_excluded": True,
        },
        "without_yongqiang": without_yongqiang,
        "with_yongqiang": with_yongqiang,
        "audit": {
            "checks": checks,
            "status": "PASS",
            "proxy_15x_added_volume": proxy_15x_total,
            "yongqiang_detail_total": yq_detail_total,
            "yongqiang_declared_total": declared_total,
            "yongqiang_declared_total_difference": declared_difference,
            "yongqiang_declared_total_status": (
                "NOT_AVAILABLE"
                if declared_total is None
                else "PASS" if declared_difference == 0 else "WARNING"
            ),
        },
    }

    if output_path is not None:
        output = write_json(result, output_path)
        status("A4", f"已输出用户侧电量汇总：{output}")
    status("A4", f"汇总完成：不含永强 {without_total}，含永强 {with_total}")
    return result


def _main() -> None:
    parser = argparse.ArgumentParser(description="生成不含永强、含永强用户侧电量汇总")
    parser.add_argument("--input", required=True)
    parser.add_argument("--pumped-storage-kwh", default="0")
    parser.add_argument("--output", default="outputs/a_group/a_4_user_volume_summary.json")
    args = parser.parse_args()
    a_4_build_user_totals(read_json(args.input), args.pumped_storage_kwh, args.output)


if __name__ == "__main__":
    _main()

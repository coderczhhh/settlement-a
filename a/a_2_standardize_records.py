from __future__ import annotations

import argparse
from collections import Counter
from decimal import Decimal
from pathlib import Path
from typing import Any

from a.a_utils import (
    cell,
    clean_text,
    find_workbook,
    get_sheet,
    progress,
    read_json,
    status,
    to_decimal,
    write_json,
)


USER_SHEET = "1.结算_用户属性用户侧售电信息"
PROXY_15X_SHEET = "4.结算_1.5倍代理购电信息"
YONGQIANG_SHEET = "luxi"
EXCLUDED_USE_CATEGORIES = {"五、打水用电"}


def _find_header_row(sheet: dict[str, Any], required: set[str]) -> dict[str, Any]:
    for row in sheet["rows"][:20]:
        labels = {clean_text(value) for value in row["values"]}
        if required.issubset(labels):
            return row
    raise ValueError(f"sheet {sheet['sheet_name']} 前 20 行中未找到表头：{sorted(required)}")


def _header_map(header_row: dict[str, Any]) -> dict[str, int]:
    return {
        clean_text(value): index
        for index, value in enumerate(header_row["values"], start=1)
        if clean_text(value)
    }


def _record_source(workbook: dict[str, Any], sheet_name: str, row_number: int) -> dict[str, Any]:
    return {
        "workbook": workbook["file_name"],
        "sheet": sheet_name,
        "row": row_number,
    }


def _standardize_user_sales(
    workbook: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    sheet = get_sheet(workbook, USER_SHEET)
    required = {"用电分类", "分类明细", "电压等级", "市场主体类型", "分时类型", "售电电量", "电能电费"}
    header_row = _find_header_row(sheet, required)
    columns = _header_map(header_row)
    records: list[dict[str, Any]] = []
    excluded: list[dict[str, Any]] = []

    for row in progress(
        sheet["rows"][header_row["row_number"] :],
        desc="A2 标准化用户侧售电信息",
        unit="row",
    ):
        use_category = clean_text(cell(row, columns["用电分类"]))
        if use_category in EXCLUDED_USE_CATEGORIES:
            excluded.append(
                {
                    "来源行": row["row_number"],
                    "用电分类": use_category,
                    "原因": "业务口径明确排除",
                }
            )
            continue
        voltage = clean_text(cell(row, columns["电压等级"]))
        market_type = clean_text(cell(row, columns["市场主体类型"]))
        if not any((use_category, voltage, market_type)):
            continue
        source_text = f"{workbook['file_name']}!{USER_SHEET}:{row['row_number']}"
        records.append(
            {
                "用电分类": use_category,
                "分类明细": clean_text(cell(row, columns["分类明细"])),
                "电压等级": voltage,
                "市场主体类型": market_type,
                "分时类型": clean_text(cell(row, columns["分时类型"])),
                "售电电量": to_decimal(
                    cell(row, columns["售电电量"]), field="售电电量", source=source_text
                ),
                "电能电费": to_decimal(
                    cell(row, columns["电能电费"]), field="电能电费", source=source_text
                ),
                "来源": _record_source(workbook, USER_SHEET, row["row_number"]),
            }
        )
    return records, excluded


def _standardize_proxy_15x(workbook: dict[str, Any]) -> list[dict[str, Any]]:
    sheet = get_sheet(workbook, PROXY_15X_SHEET)
    header_row = next(
        (
            row
            for row in sheet["rows"][:20]
            if clean_text(cell(row, 1)) == "户号" and clean_text(cell(row, 4)) == "电压等级"
        ),
        None,
    )
    if header_row is None:
        raise ValueError(f"sheet {PROXY_15X_SHEET} 未找到多层表头")

    # 表头占 3 行；F 列为市场化交易-电量-合计，N 列为电能电费-合计。
    data_start_row = header_row["row_number"] + 3
    records: list[dict[str, Any]] = []
    for row in progress(
        [item for item in sheet["rows"] if item["row_number"] >= data_start_row],
        desc="A2 标准化 1.5 倍代理购电",
        unit="row",
    ):
        account = clean_text(cell(row, 1))
        voltage = clean_text(cell(row, 4))
        if not account and not voltage and cell(row, 6) is None:
            continue
        source_text = f"{workbook['file_name']}!{PROXY_15X_SHEET}:{row['row_number']}"
        records.append(
            {
                "户号": account,
                "户名": clean_text(cell(row, 2)),
                "用电类型": clean_text(cell(row, 3)),
                "电压等级": voltage,
                "所属市县": clean_text(cell(row, 5)),
                "售电电量": to_decimal(cell(row, 6), field="电量合计", source=source_text),
                "电能电费": to_decimal(cell(row, 14), field="电能电费合计", source=source_text),
                "来源": _record_source(workbook, PROXY_15X_SHEET, row["row_number"]),
            }
        )
    return records


def _standardize_yongqiang(
    workbook: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], Decimal | None]:
    sheet = get_sheet(workbook, YONGQIANG_SHEET)
    required = {"用电分类", "电压等级", "市场主体类型", "分时类型", "售电电量"}
    header_row = _find_header_row(sheet, required)
    columns = _header_map(header_row)
    records: list[dict[str, Any]] = []
    excluded: list[dict[str, Any]] = []
    declared_total: Decimal | None = None

    for row in progress(
        sheet["rows"][header_row["row_number"] :],
        desc="A2 标准化永强月度数据",
        unit="row",
    ):
        use_category = clean_text(cell(row, columns["用电分类"]))
        if use_category in EXCLUDED_USE_CATEGORIES:
            excluded.append(
                {
                    "来源行": row["row_number"],
                    "用电分类": use_category,
                    "原因": "业务口径明确排除",
                }
            )
            continue
        if use_category == "合计":
            declared_total = to_decimal(
                cell(row, columns["售电电量"]),
                field="永强合计电量",
                source=f"{workbook['file_name']}!{YONGQIANG_SHEET}:{row['row_number']}",
            )
            excluded.append({"来源行": row["row_number"], "原因": "合计行不参与明细再汇总"})
            continue
        if use_category.startswith("编制人"):
            excluded.append({"来源行": row["row_number"], "原因": "签字行不参与汇总"})
            continue

        voltage = clean_text(cell(row, columns["电压等级"]))
        market_type = clean_text(cell(row, columns["市场主体类型"]))
        if not any((use_category, voltage, market_type)):
            continue
        source_text = f"{workbook['file_name']}!{YONGQIANG_SHEET}:{row['row_number']}"
        records.append(
            {
                "用电分类": use_category,
                "栏目说明": clean_text(cell(row, columns.get("栏目说明", 2))),
                "用电类型": clean_text(cell(row, columns.get("用电类型", 3))),
                "电压等级": voltage,
                "市场主体类型": market_type,
                "分时类型": clean_text(cell(row, columns["分时类型"])),
                "售电电量": to_decimal(
                    cell(row, columns["售电电量"]), field="售电电量", source=source_text
                ),
                "来源": _record_source(workbook, YONGQIANG_SHEET, row["row_number"]),
            }
        )
    return records, excluded, declared_total


def a_2_standardize_records(
    raw_data: dict[str, Any], output_path: str | Path | None = None
) -> dict[str, Any]:
    """从全量字典中提取 A 组需要的三类标准明细记录。"""
    status("A2", "开始识别表头并标准化业务记录")
    user_workbook = find_workbook(raw_data, "用户侧清算数据")
    yongqiang_workbook = find_workbook(raw_data, "永强供电公司")

    user_sales, user_sales_excluded = _standardize_user_sales(user_workbook)
    proxy_15x = _standardize_proxy_15x(user_workbook)
    yongqiang, yongqiang_excluded, declared_total = _standardize_yongqiang(yongqiang_workbook)

    result = {
        "schema_version": "1.1",
        "unit": {"售电电量": "千瓦时", "电能电费": "元"},
        "records": {
            "user_sales": user_sales,
            "proxy_15x": proxy_15x,
            "yongqiang": yongqiang,
        },
        "excluded_rows": {
            "user_sales": user_sales_excluded,
            "yongqiang": yongqiang_excluded,
            "yongqiang_Sheet3": [
                {
                    "sheet": "Sheet3",
                    "reason": "辅助计算 sheet；不加入 luxi 主数据，避免重复汇总",
                }
            ],
        },
        "statistics": {
            "user_sales_rows": len(user_sales),
            "user_sales_excluded_rows": len(user_sales_excluded),
            "proxy_15x_rows": len(proxy_15x),
            "yongqiang_rows": len(yongqiang),
            "user_market_types": dict(Counter(item["市场主体类型"] for item in user_sales)),
            "yongqiang_market_types": dict(
                Counter(item["市场主体类型"] for item in yongqiang)
            ),
            "yongqiang_declared_total": declared_total,
        },
    }

    if output_path is not None:
        output = write_json(result, output_path)
        status("A2", f"已输出标准明细：{output}")
    status(
        "A2",
        f"标准化完成：用户侧 {len(user_sales)} 行，1.5倍代理 {len(proxy_15x)} 行，永强 {len(yongqiang)} 行",
    )
    return result


def _main() -> None:
    parser = argparse.ArgumentParser(description="从 A1 JSON 提取标准业务记录")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", default="outputs/a_group/a_2_standardized_records.json")
    args = parser.parse_args()
    a_2_standardize_records(read_json(args.input), args.output)


if __name__ == "__main__":
    _main()

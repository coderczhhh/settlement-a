from __future__ import annotations

import json
import math
import sys
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any


def status(step: str, message: str) -> None:
    """输出带时间和步骤名的状态。"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now}] [{step}] {message}", flush=True)


def progress(iterable: Any, *, desc: str, unit: str = "item") -> Any:
    """零依赖文本进度条，避免因 tqdm 未安装导致业务流程中断。"""
    items = iterable if hasattr(iterable, "__len__") else list(iterable)
    total = len(items)
    if total == 0:
        print(f"{desc}: [##############################] 0/0 {unit}", file=sys.stderr)
        return iter(())

    def _generator() -> Any:
        last_percent = -1
        for index, item in enumerate(items, start=1):
            percent = int(index * 100 / total)
            if percent != last_percent and (percent % 5 == 0 or index == 1 or index == total):
                width = 30
                filled = int(width * index / total)
                bar = "#" * filled + "-" * (width - filled)
                print(
                    f"\r{desc}: [{bar}] {index}/{total} {unit} {percent:3d}%",
                    end="\n" if index == total else "",
                    file=sys.stderr,
                    flush=True,
                )
                last_percent = percent
            yield item

    return _generator()


def to_decimal(value: Any, *, field: str, source: str) -> Decimal:
    """把 Excel/JSON 数值转换为 Decimal，空值按 0 处理。"""
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, bool):
        raise ValueError(f"{source} 字段 {field} 不应为布尔值：{value!r}")
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        raise ValueError(f"{source} 字段 {field} 不是有限数：{value!r}")
    try:
        return Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, AttributeError) as exc:
        raise ValueError(f"{source} 字段 {field} 无法转换为数值：{value!r}") from exc


def json_ready(value: Any) -> Any:
    """递归转换为 JSON 可序列化类型。"""
    if isinstance(value, Decimal):
        return int(value) if value == value.to_integral_value() else float(value)
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    if isinstance(value, dict):
        return {str(key): json_ready(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_ready(item) for item in value]
    return value


def write_json(data: Any, path: str | Path) -> Path:
    """以 UTF-8 及中文可读格式写入 JSON。"""
    output = Path(path)
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(
        json.dumps(json_ready(data), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return output


def read_json(path: str | Path) -> dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def clean_text(value: Any) -> str:
    """业务字段标准化：空值转空字符串，其余去除首尾空白。"""
    return "" if value is None else str(value).strip()


def cell(row: dict[str, Any], column_number: int) -> Any:
    """从 A1 读取结果的行记录中取指定列（1 起始）。"""
    values = row["values"]
    return values[column_number - 1] if column_number <= len(values) else None


def find_workbook(raw_data: dict[str, Any], name_fragment: str) -> dict[str, Any]:
    matches = [
        workbook
        for workbook in raw_data["workbooks"]
        if name_fragment in workbook["file_name"]
    ]
    if len(matches) != 1:
        raise ValueError(
            f"期望唯一匹配工作簿 {name_fragment!r}，实际匹配 {len(matches)} 个"
        )
    return matches[0]


def get_sheet(workbook: dict[str, Any], sheet_name: str) -> dict[str, Any]:
    try:
        return workbook["sheets"][sheet_name]
    except KeyError as exc:
        raise KeyError(
            f"工作簿 {workbook['file_name']} 缺少 sheet：{sheet_name}"
        ) from exc


def value_by_label(rows: list[dict[str, Any]], label: str) -> Decimal:
    for row in rows:
        if row["用户类型"] == label:
            return to_decimal(row["用电量"], field="用电量", source=label)
    raise KeyError(f"汇总结果中缺少行：{label}")

from __future__ import annotations

from collections import OrderedDict
from copy import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


ProgressCallback = Callable[[int, str], None]
OrderProgressCallback = Callable[[int, int], None]


class SplitterError(Exception):
    """An input workbook cannot be processed safely."""


@dataclass
class SplitStats:
    orders: int = 0
    rows: int = 0
    matched_rows: int = 0
    unmatched_rows: int = 0
    exceptional_orders: int = 0
    warnings: list[str] = field(default_factory=list)
    output_path: str = ""


SALES_HEADERS = {
    "订单编号",
    "应收合计",
    "货品编号",
    "数量",
    "单价",
    "优惠",
    "折扣",
    "金额",
    "锁定待发",
}

RATIO_HEADERS = {"单品编号", "金额", "分摊比例", "单价"}
RESULT_HEADERS = {
    27: "组合装单价",  # AA
    28: "占比",        # AB
    29: "单价",        # AC
    30: "金额",        # AD
    34: "摊后金额",    # AH
}


def _text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _code(value: object) -> str:
    return _text(value).upper()


def _number(value: object, label: str) -> float:
    if isinstance(value, bool):
        raise ValueError(f"{label}不是数字")
    if isinstance(value, (int, float)):
        return float(value)
    text = _text(value).replace(",", "")
    if not text:
        raise ValueError(f"{label}为空")
    try:
        return float(text)
    except ValueError as exc:
        raise ValueError(f"{label}不是数字：{value}") from exc


def _find_sheet(workbook, required: set[str], preferred: str | None = None):
    candidates = list(workbook.worksheets)
    if preferred in workbook.sheetnames:
        candidates.remove(workbook[preferred])
        candidates.insert(0, workbook[preferred])

    for sheet in candidates:
        for row_number in range(1, min(sheet.max_row, 5) + 1):
            headers = {_text(sheet.cell(row_number, col).value) for col in range(1, sheet.max_column + 1)}
            if required.issubset(headers):
                return sheet, row_number
    return None, None


def _header_map(sheet, header_row: int) -> dict[str, int]:
    result: dict[str, int] = {}
    for column in range(1, sheet.max_column + 1):
        name = _text(sheet.cell(header_row, column).value)
        if name and name not in result:
            result[name] = column
    return result


def _load_ratio_prices(ratio_path: str | Path) -> tuple[OrderedDict[str, float], int]:
    path = Path(ratio_path)
    if path.suffix.lower() != ".xlsx":
        raise SplitterError("组合装拆分占比表必须是 .xlsx 文件。")

    try:
        values_book = load_workbook(path, data_only=True, read_only=True, keep_links=False)
        formulas_book = load_workbook(path, data_only=False, read_only=True, keep_links=False)
    except Exception as exc:
        raise SplitterError(f"无法打开组合装拆分占比表：{exc}") from exc

    try:
        value_sheet, header_row = _find_sheet(values_book, RATIO_HEADERS, "Sheet2")
        if value_sheet is None:
            raise SplitterError("占比表中未找到包含“单品编号、金额、分摊比例、单价”的工作表。")
        formula_sheet = formulas_book[value_sheet.title]
        headers = _header_map(value_sheet, header_row)
        code_col = headers["单品编号"]
        price_col = headers["单价"]

        prices: OrderedDict[str, float] = OrderedDict()
        missing_cache: list[int] = []
        source_rows = 0
        for row in range(header_row + 1, value_sheet.max_row + 1):
            item_code = _code(value_sheet.cell(row, code_col).value)
            if not item_code:
                continue
            source_rows += 1
            # VLOOKUP returns the first occurrence, so later duplicates are ignored.
            if item_code in prices:
                continue
            cached_value = value_sheet.cell(row, price_col).value
            formula_value = formula_sheet.cell(row, price_col).value
            if cached_value is None and isinstance(formula_value, str) and formula_value.startswith("="):
                missing_cache.append(row)
                continue
            try:
                prices[item_code] = _number(cached_value, f"占比表第 {row} 行单价")
            except ValueError as exc:
                raise SplitterError(str(exc)) from exc

        if missing_cache:
            preview = "、".join(map(str, missing_cache[:8]))
            raise SplitterError(
                f"占比表 E 列公式没有可读取的计算结果（第 {preview} 行）。"
                "请先用 Excel 打开该文件，完成计算并保存后再试。"
            )
        if not prices:
            raise SplitterError("占比表没有可用的单品编号和拆分单价。")
        return prices, source_rows
    finally:
        values_book.close()
        formulas_book.close()


def _copy_column_layout(sheet, old_layout: dict[int, dict[str, object]]) -> None:
    mapping: dict[int, int] = {}
    for old_col in range(1, 27):
        mapping[old_col] = old_col
    for old_col in range(27, 30):
        mapping[old_col] = old_col + 4
    for old_col in range(30, 35):
        mapping[old_col] = old_col + 5

    for old_col, new_col in mapping.items():
        props = old_layout.get(old_col, {})
        dim = sheet.column_dimensions[get_column_letter(new_col)]
        for name, value in props.items():
            setattr(dim, name, value)

    widths = {27: 16, 28: 15, 29: 16, 30: 16, 34: 16}
    for col, width in widths.items():
        sheet.column_dimensions[get_column_letter(col)].width = width


def _capture_column_layout(sheet) -> dict[int, dict[str, object]]:
    layout: dict[int, dict[str, object]] = {}
    for col in range(1, sheet.max_column + 1):
        dim = sheet.column_dimensions[get_column_letter(col)]
        layout[col] = {
            "width": dim.width,
            "hidden": dim.hidden,
            "outlineLevel": dim.outlineLevel,
            "bestFit": dim.bestFit,
            "collapsed": dim.collapsed,
        }
    return layout


def _style_new_columns(sheet, header_row: int) -> None:
    source_style_col = 26  # Z
    amount_style_col = 33  # AG, the original 金额 after the first insertion

    for row in range(1, sheet.max_row + 1):
        source = sheet.cell(row, source_style_col)
        for col in (27, 28, 29, 30):
            target = sheet.cell(row, col)
            if source.has_style:
                target._style = copy(source._style)
            target.number_format = "0.00"
        amount_source = sheet.cell(row, amount_style_col)
        target = sheet.cell(row, 34)
        if amount_source.has_style:
            target._style = copy(amount_source._style)
        target.number_format = "0.00"

    header_fill = PatternFill("solid", fgColor="245C73")
    for col, title in RESULT_HEADERS.items():
        cell = sheet.cell(header_row, col, title)
        cell.fill = header_fill
        cell.font = Font(name="Microsoft YaHei UI", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _extend_auto_filter(sheet) -> None:
    if sheet.auto_filter and sheet.auto_filter.ref:
        ref = sheet.auto_filter.ref
        if ":" in ref:
            start, end = ref.split(":", 1)
            end_row = "".join(ch for ch in end if ch.isdigit()) or str(sheet.max_row)
            sheet.auto_filter.ref = f"{start}:AM{end_row}"


def _prepare_result_sheet(workbook, source_sheet, header_row: int):
    old_layout = _capture_column_layout(source_sheet)
    if "拆分结果" in workbook.sheetnames:
        workbook.remove(workbook["拆分结果"])
    result_sheet = workbook.copy_worksheet(source_sheet)
    result_sheet.title = "拆分结果"
    result_sheet.insert_cols(27, amount=4)
    result_sheet.insert_cols(34, amount=1)
    _copy_column_layout(result_sheet, old_layout)
    _style_new_columns(result_sheet, header_row)
    _extend_auto_filter(result_sheet)
    return result_sheet


def _zero_targets(sheet, row: int) -> None:
    for col in RESULT_HEADERS:
        sheet.cell(row, col, 0.0)


def _calculate_rows(
    sheet,
    header_row: int,
    prices: OrderedDict[str, float],
    order_progress: OrderProgressCallback | None = None,
) -> SplitStats:
    headers = _header_map(sheet, header_row)
    # These columns are unchanged by insertions because they all sit at or before Z.
    required = {"订单编号", "应收合计", "货品编号", "数量"}
    missing = required.difference(headers)
    if missing:
        raise SplitterError(f"销售表缺少字段：{'、'.join(sorted(missing))}")

    order_col = headers["订单编号"]
    total_col = headers["应收合计"]
    item_col = headers["货品编号"]
    quantity_col = headers["数量"]

    groups: OrderedDict[str, list[int]] = OrderedDict()
    for row in range(header_row + 1, sheet.max_row + 1):
        if all(sheet.cell(row, col).value is None for col in range(1, 27)):
            continue
        order_number = _text(sheet.cell(row, order_col).value)
        key = order_number or f"__ROW_{row}"
        groups.setdefault(key, []).append(row)

    stats = SplitStats(orders=len(groups), rows=sum(len(rows) for rows in groups.values()))
    total_orders = len(groups)
    for order_index, (order_number, rows) in enumerate(groups.items(), start=1):
        row_prices: dict[int, float] = {}
        invalid_reason: str | None = None

        for row in rows:
            item_code = _code(sheet.cell(row, item_col).value)
            if item_code in prices:
                row_prices[row] = prices[item_code]
                stats.matched_rows += 1
            else:
                row_prices[row] = 0.0
                stats.unmatched_rows += 1

        total_price = sum(row_prices.values())
        if abs(total_price) < 1e-15:
            invalid_reason = "没有可分摊的匹配单价"

        quantities: dict[int, float] = {}
        order_totals: dict[int, float] = {}
        if invalid_reason is None:
            for row in rows:
                try:
                    quantities[row] = _number(sheet.cell(row, quantity_col).value, f"第 {row} 行数量")
                    order_totals[row] = _number(sheet.cell(row, total_col).value, f"第 {row} 行应收合计")
                    if row_prices[row] != 0 and quantities[row] == 0:
                        invalid_reason = f"第 {row} 行匹配成功但数量为 0"
                        break
                except ValueError as exc:
                    invalid_reason = str(exc)
                    break

        if invalid_reason is not None:
            stats.exceptional_orders += 1
            display_order = order_number.replace("__ROW_", "第 ")
            stats.warnings.append(f"订单 {display_order}：{invalid_reason}，拆分列已填 0。")
            for row in rows:
                _zero_targets(sheet, row)
            if order_progress:
                order_progress(order_index, total_orders)
            continue

        for row in rows:
            aa = row_prices[row]
            if aa == 0:
                _zero_targets(sheet, row)
                continue
            quantity = quantities[row]
            ab = aa / total_price / quantity
            ac = order_totals[row] * ab
            ad = ac * quantity
            sheet.cell(row, 27, aa)
            sheet.cell(row, 28, ab)
            sheet.cell(row, 29, ac)
            sheet.cell(row, 30, ad)
            sheet.cell(row, 34, ad)

        if order_progress:
            order_progress(order_index, total_orders)

    return stats


def process_workbooks(
    ratio_path: str | Path,
    sales_path: str | Path,
    output_path: str | Path,
    progress: ProgressCallback | None = None,
) -> SplitStats:
    """Generate a numeric result workbook that mirrors the reference formulas."""

    def report(percent: int, message: str) -> None:
        if progress:
            progress(percent, message)

    ratio_path = Path(ratio_path)
    sales_path = Path(sales_path)
    output_path = Path(output_path)
    if ratio_path.resolve() == sales_path.resolve():
        raise SplitterError("组合装占比表和销售单不能是同一个文件。")
    if sales_path.suffix.lower() != ".xlsx":
        raise SplitterError("销售单必须是 .xlsx 文件。")
    if output_path.suffix.lower() != ".xlsx":
        raise SplitterError("输出文件必须使用 .xlsx 扩展名。")
    if output_path.resolve() in {ratio_path.resolve(), sales_path.resolve()}:
        raise SplitterError("输出文件不能覆盖上传的源文件。")

    report(10, "读取组合装拆分占比表")
    prices, _ = _load_ratio_prices(ratio_path)
    report(30, f"已载入 {len(prices)} 个首匹配单品")

    try:
        workbook = load_workbook(sales_path, data_only=False, keep_links=False)
    except Exception as exc:
        raise SplitterError(f"无法打开销售单：{exc}") from exc

    try:
        source_sheet, header_row = _find_sheet(workbook, SALES_HEADERS)
        if source_sheet is None:
            raise SplitterError("销售单中未找到包含完整销售字段的明细工作表。")
        report(45, f"复制销售明细表“{source_sheet.title}”")
        result_sheet = _prepare_result_sheet(workbook, source_sheet, header_row)
        report(60, "按订单计算拆分单价")
        stats = _calculate_rows(
            result_sheet,
            header_row,
            prices,
            order_progress=lambda done, total: report(
                60 + int((done / total) * 28) if total else 88,
                f"正在拆分订单 {done}/{total}",
            ),
        )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        report(90, "保存结果工作簿")
        workbook.save(output_path)
        stats.output_path = str(output_path)
        report(100, "拆分完成")
        return stats
    except SplitterError:
        raise
    except Exception as exc:
        raise SplitterError(f"生成结果时发生错误：{exc}") from exc
    finally:
        workbook.close()


def default_output_path(sales_path: str | Path) -> Path:
    from datetime import datetime

    source = Path(sales_path)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return source.with_name(f"{source.stem}_已拆单价_{stamp}.xlsx")

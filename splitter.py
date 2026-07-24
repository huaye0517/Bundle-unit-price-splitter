from __future__ import annotations

from collections import OrderedDict
from copy import copy
from dataclasses import dataclass, field
import os
from pathlib import Path
import re
import shutil
from typing import Callable, Iterable

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties


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


@dataclass
class RatioUpdateStats:
    mode: str
    added_rows: int
    source_rows: int
    unique_items: int


@dataclass(frozen=True)
class RatioEntry:
    price: float
    ratio: float
    row: int


@dataclass
class RatioData:
    prices: OrderedDict[str, float]
    pairs: dict[tuple[str, str], RatioEntry]
    source_rows: int
    has_parent_ratios: bool


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

RATIO_SCHEMAS = (
    (
        {"单品编号", "金额", "分摊比例", "单价"},
        "单品编号",
        "单价",
        "",
        "",
        "Sheet2",
    ),
    (
        {"母件编号", "编号", "执行价格", "分摊金额", "分摊比例"},
        "编号",
        "执行价格",
        "母件编号",
        "分摊比例",
        "sheet1",
    ),
)
RESULT_HEADERS = {
    27: "组合装单价",  # AA
    28: "占比",        # AB
    29: "单价",        # AC
    30: "金额",        # AD
    34: "手续费1%",    # AH
    35: "最终金额",    # AI
    36: "摊后金额",    # AJ
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

    matches = []
    for sheet in candidates:
        for row_number in range(1, min(sheet.max_row, 5) + 1):
            headers = {_text(sheet.cell(row_number, col).value) for col in range(1, sheet.max_column + 1)}
            if required.issubset(headers):
                matches.append((sheet, row_number))
                break
    if matches:
        return max(matches, key=lambda match: (match[0].title == preferred, match[0].max_row))
    return None, None


def _header_map(sheet, header_row: int) -> dict[str, int]:
    result: dict[str, int] = {}
    for column in range(1, sheet.max_column + 1):
        name = _text(sheet.cell(header_row, column).value)
        if name and name not in result:
            result[name] = column
    return result


def _ratio_sheet_info(workbook):
    for (
        required,
        code_header,
        price_header,
        mother_header,
        ratio_header,
        preferred,
    ) in RATIO_SCHEMAS:
        sheet, header_row = _find_sheet(workbook, required, preferred)
        if sheet is not None:
            return (
                sheet,
                header_row,
                code_header,
                price_header,
                mother_header,
                ratio_header,
            )
    return None, None, "", "", "", ""


def _load_ratio_data(ratio_path: str | Path) -> RatioData:
    path = Path(ratio_path)
    if path.suffix.lower() != ".xlsx":
        raise SplitterError("组合装拆分占比表必须是 .xlsx 文件。")

    try:
        values_book = load_workbook(path, data_only=True, read_only=True, keep_links=False)
        formulas_book = load_workbook(path, data_only=False, read_only=True, keep_links=False)
    except Exception as exc:
        raise SplitterError(f"无法打开组合装拆分占比表：{exc}") from exc

    try:
        (
            value_sheet,
            header_row,
            code_header,
            price_header,
            mother_header,
            ratio_header,
        ) = _ratio_sheet_info(values_book)
        if value_sheet is None:
            raise SplitterError(
                "占比表中未找到可识别的工作表。支持旧版“单品编号、金额、分摊比例、单价”，"
                "或新版“母件编号、编号、执行价格、分摊金额、分摊比例”。"
            )
        formula_sheet = formulas_book[value_sheet.title]
        headers = _header_map(value_sheet, header_row)
        code_col = headers[code_header]
        price_col = headers[price_header]
        mother_col = headers[mother_header] if mother_header else None
        ratio_col = headers[ratio_header] if ratio_header else None

        prices: OrderedDict[str, float] = OrderedDict()
        pairs: dict[tuple[str, str], RatioEntry] = {}
        missing_cache: list[tuple[int, str]] = []
        source_rows = 0
        value_rows = value_sheet.iter_rows(min_row=header_row + 1, values_only=True)
        formula_rows = formula_sheet.iter_rows(min_row=header_row + 1, values_only=True)
        for row, (values, formulas) in enumerate(
            zip(value_rows, formula_rows), start=header_row + 1
        ):
            item_code = _code(values[code_col - 1])
            if not item_code:
                continue
            cached_value = values[price_col - 1]
            # 合并多个导出文件时，数据区中可能再次出现完整表头。
            if item_code == _code(code_header) and _text(cached_value) == price_header:
                continue
            source_rows += 1
            formula_value = formulas[price_col - 1]
            if cached_value is None and isinstance(formula_value, str) and formula_value.startswith("="):
                missing_cache.append((row, price_header))
                continue
            try:
                price = _number(cached_value, f"占比表第 {row} 行{price_header}")
            except ValueError as exc:
                raise SplitterError(str(exc)) from exc
            prices.setdefault(item_code, price)

            if mother_col is None or ratio_col is None:
                continue

            mother_code = _code(values[mother_col - 1])
            if not mother_code:
                raise SplitterError(f"占比表第 {row} 行母件编号为空。")
            ratio_value = values[ratio_col - 1]
            ratio_formula = formulas[ratio_col - 1]
            if ratio_value is None and isinstance(ratio_formula, str) and ratio_formula.startswith("="):
                missing_cache.append((row, ratio_header))
                continue
            try:
                ratio = _number(ratio_value, f"占比表第 {row} 行{ratio_header}")
            except ValueError as exc:
                raise SplitterError(str(exc)) from exc

            key = (mother_code, item_code)
            existing = pairs.get(key)
            if existing is not None:
                if (
                    abs(existing.price - price) > 1e-9
                    or abs(existing.ratio - ratio) > 1e-9
                ):
                    raise SplitterError(
                        f"占比表母件 {mother_code}、子件 {item_code} 在第 "
                        f"{existing.row} 行和第 {row} 行存在不同的执行价格或分摊比例。"
                    )
                continue
            pairs[key] = RatioEntry(price, ratio, row)

        if missing_cache:
            preview = "、".join(
                f"{row} 行{header}" for row, header in missing_cache[:8]
            )
            raise SplitterError(
                f"占比表公式没有可读取的计算结果（第 {preview}）。"
                "请先用 Excel 打开该文件，完成计算并保存后再试。"
            )
        if not prices:
            raise SplitterError("占比表没有可用的子件编号和拆分单价。")
        if mother_col is not None and not pairs:
            raise SplitterError("占比表没有可用的母件编号、子件编号和分摊比例。")
        return RatioData(
            prices=prices,
            pairs=pairs,
            source_rows=source_rows,
            has_parent_ratios=mother_col is not None,
        )
    finally:
        values_book.close()
        formulas_book.close()


def ratio_file_info(ratio_path: str | Path) -> tuple[int, int]:
    ratio_data = _load_ratio_data(ratio_path)
    return len(ratio_data.prices), ratio_data.source_rows


def update_ratio_data(
    current_path: str | Path,
    incoming_path: str | Path,
    mode: str,
) -> RatioUpdateStats:
    current = Path(current_path)
    incoming = Path(incoming_path)
    if mode not in {"append", "replace"}:
        raise SplitterError("基础数据更新方式必须是新增或覆盖。")

    incoming_items, incoming_rows = ratio_file_info(incoming)
    current.parent.mkdir(parents=True, exist_ok=True)
    temporary = current.with_name(current.stem + ".updating.xlsx")

    if mode == "replace":
        shutil.copyfile(incoming, temporary)
        os.replace(temporary, current)
        return RatioUpdateStats(mode, incoming_rows, incoming_rows, incoming_items)

    ratio_file_info(current)
    try:
        base_book = load_workbook(current, data_only=False, keep_links=False)
        incoming_book = load_workbook(incoming, data_only=True, keep_links=False)
    except Exception as exc:
        raise SplitterError(f"无法打开基础数据文件：{exc}") from exc

    try:
        (
            base_sheet,
            base_header_row,
            base_code_header,
            base_price_header,
            _,
            _,
        ) = _ratio_sheet_info(base_book)
        (
            new_sheet,
            new_header_row,
            new_code_header,
            new_price_header,
            _,
            _,
        ) = _ratio_sheet_info(incoming_book)
        if base_sheet is None or new_sheet is None:
            raise SplitterError("新增文件中未找到可识别的组合装数据工作表。")

        base_headers = _header_map(base_sheet, base_header_row)
        new_headers = _header_map(new_sheet, new_header_row)
        base_mother_header = "母件编号" if "母件编号" in base_headers else "组合装编号"
        new_mother_header = "母件编号" if "母件编号" in new_headers else "组合装编号"
        base_mother_col = base_headers.get(base_mother_header)
        new_mother_col = new_headers.get(new_mother_header)
        base_code_col = base_headers[base_code_header]
        new_code_col = new_headers[new_code_header]
        new_price_col = new_headers[new_price_header]

        existing: set[tuple[str, str]] = set()
        for values in base_sheet.iter_rows(min_row=base_header_row + 1, values_only=True):
            code = _code(values[base_code_col - 1])
            if not code or code == _code(base_code_header):
                continue
            mother = _code(values[base_mother_col - 1]) if base_mother_col else ""
            existing.add((mother, code))

        style_row = min(base_header_row + 1, base_sheet.max_row)
        added_rows = 0
        for values in new_sheet.iter_rows(min_row=new_header_row + 1, values_only=True):
            code = _code(values[new_code_col - 1])
            if not code or code == _code(new_code_header):
                continue
            mother = _code(values[new_mother_col - 1]) if new_mother_col else ""
            key = (mother, code)
            if key in existing:
                continue

            target_row = base_sheet.max_row + 1
            for header, target_col in base_headers.items():
                source_header = header
                if header == base_code_header:
                    source_header = new_code_header
                elif header == base_price_header:
                    source_header = new_price_header
                elif header == base_mother_header:
                    source_header = new_mother_header
                source_col = new_headers.get(source_header)
                value = values[source_col - 1] if source_col else None
                target = base_sheet.cell(target_row, target_col, value)
                template = base_sheet.cell(style_row, target_col)
                if template.has_style:
                    target._style = copy(template._style)
                target.number_format = template.number_format
                target.alignment = copy(template.alignment)
            existing.add(key)
            added_rows += 1

        if base_sheet.auto_filter.ref:
            start = base_sheet.auto_filter.ref.split(":", 1)[0]
            base_sheet.auto_filter.ref = f"{start}:{get_column_letter(base_sheet.max_column)}{base_sheet.max_row}"
        for table in base_sheet.tables.values():
            start = table.ref.split(":", 1)[0]
            table.ref = f"{start}:{get_column_letter(base_sheet.max_column)}{base_sheet.max_row}"

        base_book.save(temporary)
    finally:
        base_book.close()
        incoming_book.close()

    os.replace(temporary, current)
    unique_items, source_rows = ratio_file_info(current)
    return RatioUpdateStats(mode, added_rows, source_rows, unique_items)


def _copy_column_layout(sheet, old_layout: dict[int, dict[str, object]]) -> None:
    mapping: dict[int, int] = {}
    for old_col in range(1, 27):
        mapping[old_col] = old_col
    for old_col in range(27, 30):
        mapping[old_col] = old_col + 4
    for old_col in range(30, 35):
        mapping[old_col] = old_col + 7

    for old_col, new_col in mapping.items():
        props = old_layout.get(old_col, {})
        dim = sheet.column_dimensions[get_column_letter(new_col)]
        for name, value in props.items():
            setattr(dim, name, value)

    widths = {27: 16, 28: 15, 29: 16, 30: 16, 34: 13, 35: 13, 36: 16}
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
        for col in (34, 35, 36):
            target = sheet.cell(row, col)
            if amount_source.has_style:
                target._style = copy(amount_source._style)
        sheet.cell(row, 36).number_format = "0.00"

    header_fill = PatternFill("solid", fgColor="245C73")
    for col, title in RESULT_HEADERS.items():
        cell = sheet.cell(header_row, col, title)
        cell.fill = header_fill
        cell.font = Font(name="Microsoft YaHei UI", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _extend_auto_filter(sheet, header_row: int) -> None:
    if sheet.auto_filter and sheet.auto_filter.ref:
        ref = sheet.auto_filter.ref
        if ":" in ref:
            start, end = ref.split(":", 1)
            end_row = "".join(ch for ch in end if ch.isdigit()) or str(sheet.max_row)
            sheet.auto_filter.ref = f"{start}:AO{end_row}"
            return
    sheet.auto_filter.ref = f"A{header_row}:AO{sheet.max_row}"


def _refresh_pivots_on_open(workbook) -> None:
    for sheet in workbook.worksheets:
        for pivot in getattr(sheet, "_pivots", ()):
            cache = getattr(pivot, "cache", None)
            if cache is not None:
                cache.enableRefresh = True
                cache.refreshOnLoad = True


def _freeze_cached_amounts(workbook, cached_workbook) -> int:
    frozen = 0
    for sheet in workbook.worksheets:
        if sheet.title not in cached_workbook.sheetnames:
            continue
        for header_row in range(1, min(sheet.max_row, 5) + 1):
            headers = _header_map(sheet, header_row)
            amount_columns = [
                headers[header]
                for header in ("订单金额", "应收合计")
                if header in headers
            ]
            if not amount_columns:
                continue
            cached_sheet = cached_workbook[sheet.title]
            for column in amount_columns:
                cached_values = cached_sheet.iter_rows(
                    min_row=header_row + 1,
                    max_row=sheet.max_row,
                    min_col=column,
                    max_col=column,
                    values_only=True,
                )
                for row, values in enumerate(cached_values, start=header_row + 1):
                    formula = sheet.cell(row, column).value
                    if isinstance(formula, str) and formula.startswith("="):
                        sheet.cell(row, column).value = values[0]
                        if values[0] is not None:
                            frozen += 1
            break
    return frozen


def _load_cached_order_amounts(cached_workbook) -> dict[str, float]:
    order_amounts: dict[str, float] = {}
    for sheet in cached_workbook.worksheets:
        for header_row in range(1, min(sheet.max_row, 5) + 1):
            headers = _header_map(sheet, header_row)
            if not {"网店订单号", "订单金额"}.issubset(headers):
                continue
            web_order_col = headers["网店订单号"]
            order_amount_col = headers["订单金额"]
            for row, values in enumerate(
                sheet.iter_rows(min_row=header_row + 1, values_only=True),
                start=header_row + 1,
            ):
                web_order = _text(values[web_order_col - 1])
                amount_value = values[order_amount_col - 1]
                if not web_order or _text(amount_value) == "":
                    continue
                amount = _number(amount_value, f"{sheet.title}第 {row} 行订单金额")
                if web_order in order_amounts and abs(order_amounts[web_order] - amount) > 1e-9:
                    raise SplitterError(f"Sheet1 网店订单 {web_order} 存在多个不同的订单金额。")
                order_amounts[web_order] = amount
            return order_amounts
    return order_amounts


def _prepare_result_sheet(workbook, source_sheet, header_row: int):
    old_layout = _capture_column_layout(source_sheet)
    if "拆分结果" in workbook.sheetnames:
        workbook.remove(workbook["拆分结果"])
    result_sheet = workbook.copy_worksheet(source_sheet)
    result_sheet.title = "拆分结果"
    result_sheet.insert_cols(27, amount=4)
    result_sheet.insert_cols(34, amount=3)
    _copy_column_layout(result_sheet, old_layout)
    _style_new_columns(result_sheet, header_row)
    _extend_auto_filter(result_sheet, header_row)
    return result_sheet


def _zero_targets(sheet, row: int) -> None:
    for col in RESULT_HEADERS:
        sheet.cell(row, col, 0.0)


def _write_result_formulas(sheet, header_row: int) -> None:
    for row in range(header_row + 1, sheet.max_row + 1):
        if all(sheet.cell(row, col).value is None for col in range(1, 27)):
            continue
        sheet.cell(row, 34, f"=AG{row}*1%")
        sheet.cell(row, 35, f"=N{row}-AH{row}")


def _prepare_summary_sheet(
    workbook,
    result_sheet,
    header_row: int,
    web_targets: OrderedDict[str, float],
):
    if "透视表" in workbook.sheetnames:
        workbook.remove(workbook["透视表"])
    summary = workbook.create_sheet("透视表")
    summary.append(["网店订单号", "平均值项:应收合计", "求和项:金额", "差异"])

    totals: OrderedDict[str, float] = OrderedDict()
    for row in range(header_row + 1, result_sheet.max_row + 1):
        order = _text(result_sheet.cell(row, 11).value)
        if not order:
            continue
        totals.setdefault(order, 0.0)
        allocated = result_sheet.cell(row, 30).value
        if _text(allocated) != "":
            try:
                totals[order] += _number(allocated, f"第 {row} 行拆分金额")
            except ValueError:
                pass

    for order in sorted(set(totals) | set(web_targets)):
        row = summary.max_row + 1
        summary.cell(row, 1, order)
        summary.cell(row, 2, web_targets.get(order, 0.0))
        summary.cell(row, 3, totals.get(order, 0.0))
        summary.cell(row, 4, f"=C{row}-B{row}")

    total_row = summary.max_row + 1
    summary.cell(total_row, 1, "总计")
    if total_row == 2:
        summary.cell(total_row, 2, 0.0)
        summary.cell(total_row, 3, 0.0)
    else:
        summary.cell(total_row, 2, f"=SUM(B2:B{total_row - 1})")
        summary.cell(total_row, 3, f"=SUM(C2:C{total_row - 1})")
    summary.cell(total_row, 4, f"=C{total_row}-B{total_row}")

    header_fill = PatternFill("solid", fgColor="245C73")
    for cell in summary[1]:
        cell.fill = header_fill
        cell.font = Font(name="Microsoft YaHei UI", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for cell in summary[total_row]:
        cell.font = Font(name="Microsoft YaHei UI", size=10, bold=True)
    for row in range(2, total_row + 1):
        for col in range(2, 5):
            summary.cell(row, col).number_format = "0.00_);[Red]\\(0.00\\)"
    summary.column_dimensions["A"].width = 36
    summary.column_dimensions["B"].width = 18.25
    summary.column_dimensions["C"].width = 13
    summary.column_dimensions["D"].width = 12.625
    summary.freeze_panes = "A2"
    if total_row > 2:
        summary.auto_filter.ref = f"A1:D{total_row - 1}"
    return summary


def _enable_formula_recalculation(workbook) -> None:
    if workbook.calculation is None:
        workbook.calculation = CalcProperties()
    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True


def _mother_code(note: object) -> str:
    match = re.search(r"编号\s*[:：]\s*([^;；,，\s]+)", _text(note), re.IGNORECASE)
    return _code(match.group(1)) if match else ""


def _calculate_rows(
    sheet,
    header_row: int,
    ratio_data: RatioData,
    order_amounts: dict[str, float],
    order_progress: OrderProgressCallback | None = None,
) -> tuple[SplitStats, OrderedDict[str, float]]:
    headers = _header_map(sheet, header_row)
    # These columns are unchanged by insertions because they all sit at or before Z.
    required = {
        "订单编号",
        "物流单号",
        "网店订单号",
        "应收合计",
        "货品编号",
        "数量",
        "单价",
        "金额",
        "备注",
    }
    missing = required.difference(headers)
    if missing:
        raise SplitterError(f"销售表缺少字段：{'、'.join(sorted(missing))}")

    order_col = headers["订单编号"]
    logistics_col = headers["物流单号"]
    web_order_col = headers["网店订单号"]
    total_col = headers["应收合计"]
    item_col = headers["货品编号"]
    quantity_col = headers["数量"]
    unit_price_col = headers["单价"]
    note_col = headers["备注"]
    original_amount_col = max(
        col
        for col in range(1, sheet.max_column + 1)
        if _text(sheet.cell(header_row, col).value) == "金额"
    )

    groups: OrderedDict[tuple[str, str, int], list[int]] = OrderedDict()
    for row in range(header_row + 1, sheet.max_row + 1):
        if all(sheet.cell(row, col).value is None for col in range(1, 27)):
            continue
        logistics_number = _text(sheet.cell(row, logistics_col).value)
        order_number = _text(sheet.cell(row, order_col).value)
        web_order_number = _text(sheet.cell(row, web_order_col).value)
        if logistics_number:
            key = ("物流单号", logistics_number, 0)
        elif order_number:
            key = ("订单编号", order_number, 0)
        elif web_order_number:
            key = ("网店订单号", web_order_number, 0)
        else:
            key = ("单行", "", row)
        groups.setdefault(key, []).append(row)

    stats = SplitStats(orders=len(groups), rows=sum(len(rows) for rows in groups.values()))
    group_data: OrderedDict[tuple[str, str, int], dict[str, object]] = OrderedDict()
    web_groups: OrderedDict[str, list[tuple[str, str, int]]] = OrderedDict()

    for group_key, rows in groups.items():
        invalid_reason: str | None = None
        original_amounts: dict[int, float] = {}
        unit_prices: dict[int, float] = {}
        quantities: dict[int, float] = {}
        try:
            for row in rows:
                amount_value = sheet.cell(row, original_amount_col).value
                original_amounts[row] = (
                    0.0
                    if _text(amount_value) == ""
                    else _number(amount_value, f"第 {row} 行原金额")
                )
                price_value = sheet.cell(row, unit_price_col).value
                unit_prices[row] = (
                    0.0
                    if _text(price_value) == ""
                    else _number(price_value, f"第 {row} 行原单价")
                )
                quantity_value = sheet.cell(row, quantity_col).value
                quantities[row] = (
                    0.0
                    if _text(quantity_value) == ""
                    else _number(quantity_value, f"第 {row} 行数量")
                )
        except ValueError as exc:
            invalid_reason = str(exc)

        receivables: list[float] = []
        if invalid_reason is None:
            try:
                for row in rows:
                    value = sheet.cell(row, total_col).value
                    if _text(value) != "":
                        receivables.append(_number(value, f"第 {row} 行应收合计"))
            except ValueError as exc:
                invalid_reason = str(exc)
        if invalid_reason is None:
            if not receivables:
                invalid_reason = (
                    "应收合计为空或公式没有缓存值，请先用 Excel/WPS 打开并保存源文件后再上传"
                )
            elif any(abs(value - receivables[0]) > 1e-9 for value in receivables[1:]):
                invalid_reason = f"同一{group_key[0]}存在多个不同的应收合计"

        web_orders = tuple(
            dict.fromkeys(
                _text(sheet.cell(row, web_order_col).value)
                for row in rows
                if _text(sheet.cell(row, web_order_col).value)
            )
        )
        group_data[group_key] = {
            "rows": rows,
            "web_orders": web_orders,
            "original_amounts": original_amounts,
            "unit_prices": unit_prices,
            "quantities": quantities,
            "receivable": receivables[0] if receivables else 0.0,
            "invalid_reason": invalid_reason,
        }
        for web_order in web_orders:
            web_groups.setdefault(web_order, []).append(group_key)

    group_targets: dict[tuple[str, str, int], float] = {}
    for group_key, data in group_data.items():
        if data["invalid_reason"] is None:
            group_targets[group_key] = float(data["receivable"]) - sum(
                data["original_amounts"].values()
            ) * 0.01

    external_group_targets: dict[tuple[str, str, int], float] = {}
    externally_covered_orders: dict[tuple[str, str, int], set[str]] = {}
    for web_order, external_total in order_amounts.items():
        valid_keys = [
            key
            for key in web_groups.get(web_order, [])
            if group_data[key]["invalid_reason"] is None
        ]
        if not valid_keys:
            continue
        weights = [abs(float(group_data[key]["receivable"])) for key in valid_keys]
        if sum(weights) < 1e-15:
            weights = [
                abs(sum(group_data[key]["original_amounts"].values()))
                for key in valid_keys
            ]
        total_weight = sum(weights)
        if total_weight < 1e-15:
            continue
        allocated = 0.0
        for index, (key, weight) in enumerate(zip(valid_keys, weights)):
            target = (
                external_total - allocated
                if index == len(valid_keys) - 1
                else external_total * weight / total_weight
            )
            external_group_targets[key] = external_group_targets.get(key, 0.0) + target
            externally_covered_orders.setdefault(key, set()).add(web_order)
            allocated += target

    for key, target in external_group_targets.items():
        if set(group_data[key]["web_orders"]).issubset(
            externally_covered_orders.get(key, set())
        ):
            group_targets[key] = target

    total_orders = len(groups)
    for order_index, (group_key, rows) in enumerate(groups.items(), start=1):
        data = group_data[group_key]
        invalid_reason = data["invalid_reason"]
        display_order = group_key[1] or f"第 {group_key[2]} 行"
        if invalid_reason:
            stats.exceptional_orders += 1
            stats.warnings.append(f"订单 {display_order}：{invalid_reason}，拆分列已填 0。")
            for row in rows:
                _zero_targets(sheet, row)
            if order_progress:
                order_progress(order_index, total_orders)
            continue

        original_amounts: dict[int, float] = data["original_amounts"]
        unit_prices: dict[int, float] = data["unit_prices"]
        quantities: dict[int, float] = data["quantities"]
        target_total = group_targets[group_key]

        eligible_rows: list[int] = []
        for row in rows:
            if (
                abs(original_amounts[row]) >= 1e-15
                and abs(unit_prices[row]) >= 1e-15
            ):
                if quantities[row] == 0:
                    invalid_reason = f"第 {row} 行单价和金额不为 0，但数量为 0"
                    break
                eligible_rows.append(row)
            else:
                _zero_targets(sheet, row)
                stats.unmatched_rows += 1

        if invalid_reason:
            stats.exceptional_orders += 1
            stats.warnings.append(f"订单 {display_order}：{invalid_reason}，拆分列已填 0。")
            for row in rows:
                _zero_targets(sheet, row)
            if order_progress:
                order_progress(order_index, total_orders)
            continue

        if not eligible_rows:
            original_total = sum(original_amounts.values())
            if abs(original_total) < 1e-15:
                invalid_reason = "整单单价或原金额合计为 0，无法分摊订单金额"
            else:
                invalid_reason = "非零金额行的单价均为 0，无法按比例拆分"
        if invalid_reason:
            stats.exceptional_orders += 1
            stats.warnings.append(f"订单 {display_order}：{invalid_reason}，拆分列已填 0。")
            for row in rows:
                _zero_targets(sheet, row)
            if order_progress:
                order_progress(order_index, total_orders)
            continue

        if not ratio_data.has_parent_ratios:
            row_weights: dict[int, float] = {}
            for row in eligible_rows:
                item_code = _code(sheet.cell(row, item_col).value)
                price = ratio_data.prices.get(item_code, 0.0)
                if price:
                    row_weights[row] = price
                    stats.matched_rows += 1
                else:
                    stats.unmatched_rows += 1
            if not row_weights:
                row_weights = {row: original_amounts[row] for row in eligible_rows}
            weight_total = sum(row_weights.values())
            allocated = 0.0
            weighted_rows = list(row_weights)
            for row in eligible_rows:
                if row not in row_weights:
                    _zero_targets(sheet, row)
                    continue
                ad = (
                    target_total - allocated
                    if row == weighted_rows[-1]
                    else target_total * row_weights[row] / weight_total
                )
                allocated += ad
                quantity = quantities[row]
                aa = ratio_data.prices.get(
                    _code(sheet.cell(row, item_col).value), 0.0
                )
                sheet.cell(row, 27, aa)
                sheet.cell(row, 28, ad / target_total / quantity if target_total else 0.0)
                sheet.cell(row, 29, ad / quantity)
                sheet.cell(row, 30, ad)
                sheet.cell(row, 36, ad)
        else:
            allocation_groups: OrderedDict[str, list[int]] = OrderedDict()
            parents: dict[int, str] = {}
            for row in eligible_rows:
                parent = _mother_code(sheet.cell(row, note_col).value)
                parents[row] = parent
                key = f"PARENT:{parent}" if parent else f"ROW:{row}"
                allocation_groups.setdefault(key, []).append(row)

            group_weights = {
                key: sum(original_amounts[row] for row in group_rows)
                for key, group_rows in allocation_groups.items()
            }
            group_weight_total = sum(group_weights.values())
            if abs(group_weight_total) < 1e-15:
                group_weights = {
                    key: sum(
                        abs(unit_prices[row] * quantities[row])
                        for row in group_rows
                    )
                    for key, group_rows in allocation_groups.items()
                }
                group_weight_total = sum(group_weights.values())

            order_allocated = 0.0
            group_items = list(allocation_groups.items())
            for group_index, (key, group_rows) in enumerate(group_items):
                group_target = (
                    target_total - order_allocated
                    if group_index == len(group_items) - 1
                    else target_total * group_weights[key] / group_weight_total
                )
                order_allocated += group_target
                parent = parents[group_rows[0]]
                entries = {
                    row: ratio_data.pairs.get(
                        (parent, _code(sheet.cell(row, item_col).value))
                    )
                    for row in group_rows
                }
                use_ratio = bool(parent) and all(
                    entry is not None
                    and abs(entry.price) >= 1e-15
                    and abs(entry.ratio) >= 1e-15
                    for entry in entries.values()
                )

                row_weights: dict[int, float] = {}
                if use_ratio:
                    child_rows: OrderedDict[str, list[int]] = OrderedDict()
                    for row in group_rows:
                        child_rows.setdefault(
                            _code(sheet.cell(row, item_col).value), []
                        ).append(row)
                    for child, duplicate_rows in child_rows.items():
                        entry = ratio_data.pairs[(parent, child)]
                        duplicate_weights = {
                            row: abs(unit_prices[row] * quantities[row])
                            for row in duplicate_rows
                        }
                        duplicate_total = sum(duplicate_weights.values())
                        if duplicate_total < 1e-15:
                            duplicate_weights = {row: 1.0 for row in duplicate_rows}
                            duplicate_total = float(len(duplicate_rows))
                        for row in duplicate_rows:
                            row_weights[row] = (
                                entry.ratio
                                * duplicate_weights[row]
                                / duplicate_total
                            )
                    if abs(sum(row_weights.values())) < 1e-15:
                        use_ratio = False

                if use_ratio:
                    stats.matched_rows += len(group_rows)
                else:
                    row_weights = {
                        row: abs(unit_prices[row] * quantities[row])
                        for row in group_rows
                    }
                    if abs(sum(row_weights.values())) < 1e-15:
                        row_weights = {
                            row: abs(original_amounts[row]) for row in group_rows
                    }
                    for row in group_rows:
                        entry = entries[row]
                        if (
                            entry is None
                            or abs(entry.price) < 1e-15
                            or abs(entry.ratio) < 1e-15
                        ):
                            stats.unmatched_rows += 1
                        else:
                            stats.matched_rows += 1

                row_weight_total = sum(row_weights.values())
                group_allocated = 0.0
                for row_index, row in enumerate(group_rows):
                    ad = (
                        group_target - group_allocated
                        if row_index == len(group_rows) - 1
                        else group_target * row_weights[row] / row_weight_total
                    )
                    group_allocated += ad
                    quantity = quantities[row]
                    aa = (
                        entries[row].price
                        if use_ratio and entries[row] is not None
                        else unit_prices[row]
                    )
                    sheet.cell(row, 27, aa)
                    sheet.cell(
                        row,
                        28,
                        ad / target_total / quantity if target_total else 0.0,
                    )
                    sheet.cell(row, 29, ad / quantity)
                    sheet.cell(row, 30, ad)
                    sheet.cell(row, 36, ad)

        if order_progress:
            order_progress(order_index, total_orders)

    web_targets: OrderedDict[str, float] = OrderedDict()
    for row in range(header_row + 1, sheet.max_row + 1):
        web_order = _text(sheet.cell(row, web_order_col).value)
        if web_order:
            web_targets[web_order] = web_targets.get(web_order, 0.0) + float(
                sheet.cell(row, 30).value or 0.0
            )

    return stats, web_targets


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
    ratio_data = _load_ratio_data(ratio_path)
    report(30, f"已载入 {len(ratio_data.prices)} 个子件")

    try:
        workbook = load_workbook(sales_path, data_only=False, keep_links=False)
        cached_workbook = load_workbook(
            sales_path, data_only=True, read_only=True, keep_links=False
        )
    except Exception as exc:
        raise SplitterError(f"无法打开销售单：{exc}") from exc

    try:
        source_sheet, header_row = _find_sheet(workbook, SALES_HEADERS)
        if source_sheet is None:
            raise SplitterError("销售单中未找到包含完整销售字段的明细工作表。")
        _freeze_cached_amounts(workbook, cached_workbook)
        report(45, f"复制销售明细表“{source_sheet.title}”")
        result_sheet = _prepare_result_sheet(workbook, source_sheet, header_row)
        order_amounts = _load_cached_order_amounts(cached_workbook)
        report(60, "按订单计算拆分单价")
        stats, web_targets = _calculate_rows(
            result_sheet,
            header_row,
            ratio_data,
            order_amounts,
            order_progress=lambda done, total: report(
                60 + int((done / total) * 28) if total else 88,
                f"正在拆分订单 {done}/{total}",
            ),
        )
        _write_result_formulas(result_sheet, header_row)
        _prepare_summary_sheet(workbook, result_sheet, header_row, web_targets)
        _enable_formula_recalculation(workbook)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        report(90, "保存结果工作簿")
        _refresh_pivots_on_open(workbook)
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
        cached_workbook.close()


def default_output_path(sales_path: str | Path) -> Path:
    from datetime import datetime

    source = Path(sales_path)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return source.with_name(f"{source.stem}_已拆单价_{stamp}.xlsx")

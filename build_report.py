"""
Сводный отчёт по продажам из нескольких Excel-файлов.

Берёт все .xlsx из папки input/, находит в каждом листе таблицу продаж
(шапка может быть не в первой строке, колонки — в любом порядке и
с разными названиями), чистит данные и сохраняет report.xlsx с листами:

    Все продажи    — единая чистая таблица (фильтр, закреплённая шапка)
    По менеджерам  — выручка, число продаж, средний чек, доля
    По месяцам     — то же по месяцам + столбчатая диаграмма
    Топ товаров    — рейтинг товаров по выручке
    Проблемы       — строки, которые не удалось разобрать, и причина

Запуск:  python build_report.py
         python build_report.py --input "D:\\Выгрузки\\Октябрь" --output "D:\\Отчёт.xlsx"

Нужен только openpyxl (pip install openpyxl).
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

BASE_DIR = Path(__file__).resolve().parent

# ------------------------------------------------------------ настройки ---
# Как могут называться колонки в выгрузках. Сравнение без учёта регистра,
# пробелов и знаков препинания. Если у вас колонка называется иначе —
# просто допишите название в нужный список.
COLUMN_ALIASES = {
    "date": ["дата", "дата продажи", "дата сделки", "дата отгрузки",
             "дата документа", "дата заказа"],
    "manager": ["менеджер", "сотрудник", "продавец", "ответственный",
                "фио менеджера", "фио", "менеджер фио"],
    "city": ["город", "филиал", "точка", "офис", "подразделение", "магазин"],
    "product": ["товар", "наименование", "наименование товара", "продукт",
                "номенклатура", "товар услуга", "позиция"],
    "qty": ["количество", "кол во", "кол", "шт", "кол во шт", "количество шт"],
    "amount": ["сумма", "сумма руб", "сумма продажи", "выручка", "итого руб",
               "сумма с ндс", "стоимость", "сумма р"],
}
REQUIRED = ("date", "manager", "product", "amount")
FIELD_TITLES = {"date": "Дата", "manager": "Менеджер", "city": "Город",
                "product": "Товар", "qty": "Кол-во", "amount": "Сумма"}
HEADER_SCAN_ROWS = 15          # в скольких первых строках искать шапку
NO_CITY = "Не указан"
# Строки «Итого/Всего», которые менеджеры дописывают внизу таблицы
TOTAL_ROW_RE = re.compile(r"^(итог|всего)")

MONTHS_NOM = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль",
              "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
MONTH_STEMS = {"янв": 1, "фев": 2, "мар": 3, "апр": 4, "мая": 5, "май": 5,
               "июн": 6, "июл": 7, "авг": 8, "сен": 9, "окт": 10, "ноя": 11,
               "дек": 12}

# ------------------------------------------------------------ оформление ---
RUB = '#,##0.00 "₽"'
RUB0 = '#,##0 "₽"'
INT = "#,##0"
PCT = "0.0%"
DATE_FMT = "DD.MM.YYYY"

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
TOTAL_FONT = Font(bold=True)
PROBLEM_FILL = PatternFill("solid", fgColor="C00000")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOTAL_BORDER = Border(top=Side(style="medium", color="1F4E78"))


# ================================================================ модели ===
@dataclass
class Sale:
    date: date
    manager: str
    city: str
    product: str
    qty: int
    amount: Decimal
    source: str          # «файл / лист / строка N»
    raw: dict = field(default_factory=dict)


@dataclass
class Problem:
    file: str
    sheet: str
    row: int | None
    reason: str
    raw: dict = field(default_factory=dict)


# ============================================================ разбор ячеек ==
class ParseError(ValueError):
    pass


def norm_header(value) -> str:
    text = str(value or "").lower().replace("ё", "е")
    text = re.sub(r"[^\w]+", " ", text)
    return " ".join(text.split())


ALIAS_INDEX = {norm_header(a): fld for fld, names in COLUMN_ALIASES.items()
               for a in names}


def clean_text(value) -> str:
    """Убирает лишние пробелы (в т.ч. неразрывные) по краям и внутри."""
    if value is None:
        return ""
    text = str(value).replace("\u00a0", " ").replace("\u202f", " ")
    return " ".join(text.split())


def to_text(value) -> str:
    """Исходное значение ячейки — для листа «Проблемы»."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y %H:%M") if value.time() != datetime.min.time() \
            else value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    return str(value)


DATE_FORMATS = ["%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y",
                "%d-%m-%Y", "%d-%m-%y", "%Y.%m.%d", "%Y/%m/%d"]


def parse_date(value) -> date:
    if value is None or (isinstance(value, str) and not value.strip()):
        raise ParseError("не указана дата")
    if isinstance(value, datetime):
        result = value.date()
    elif isinstance(value, date):
        result = value
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        # Дата как число Excel (например, 45870 = 01.08.2025)
        if not 20000 < value < 80000:
            raise ParseError(f"дата похожа на число, а не на дату: «{value}»")
        result = from_excel(value).date()
    else:
        text = clean_text(value).lower().rstrip(" г.").strip()
        text = re.sub(r"[ t]\d{1,2}:\d{2}(:\d{2})?$", "", text).strip()  # время
        result = None
        if text.isdigit() and 20000 < int(text) < 80000:
            result = from_excel(int(text)).date()
        for fmt in DATE_FORMATS:
            if result:
                break
            try:
                result = datetime.strptime(text, fmt).date()
            except ValueError:
                pass
        if result is None:
            m = re.fullmatch(r"(\d{1,2})\s+([а-я]+)\.?\s+(\d{4})", text)
            if m and m.group(2)[:3] in MONTH_STEMS:
                try:
                    result = date(int(m.group(3)), MONTH_STEMS[m.group(2)[:3]],
                                  int(m.group(1)))
                except ValueError:
                    raise ParseError(f"такой даты не существует: «{clean_text(value)}»")
        if result is None:
            if re.fullmatch(r"\d{1,4}[./-]\d{1,2}[./-]\d{1,4}", text):
                raise ParseError(f"такой даты не существует: «{clean_text(value)}»")
            raise ParseError(f"не удалось распознать дату: «{clean_text(value)}»")
    if not 2000 <= result.year <= 2100:
        raise ParseError(f"странный год в дате: «{to_text(value)}»")
    return result


def parse_amount(value) -> Decimal:
    if value is None or (isinstance(value, str) and not value.strip()):
        raise ParseError("не указана сумма")
    if isinstance(value, bool):
        raise ParseError(f"не удалось распознать сумму: «{value}»")
    if isinstance(value, (int, float, Decimal)):
        amount = Decimal(str(value))
    else:
        text = str(value).lower()
        text = re.sub(r"(руб\.?|рублей|р\.|₽|rub)", "", text)
        text = re.sub(r"[\s\u00a0\u202f']", "", text)
        if "," in text and "." in text:
            # Десятичный разделитель — тот, что правее: 1.250,50 или 1,250.50
            if text.rfind(",") > text.rfind("."):
                text = text.replace(".", "").replace(",", ".")
            else:
                text = text.replace(",", "")
        elif text.count(",") == 1:
            text = text.replace(",", ".")
        elif text.count(",") > 1:
            text = text.replace(",", "")
        elif text.count(".") > 1:
            text = text.replace(".", "")
        if not re.fullmatch(r"-?\d+(\.\d+)?", text):
            raise ParseError(f"не удалось распознать сумму: «{clean_text(value)}»")
        try:
            amount = Decimal(text)
        except InvalidOperation:
            raise ParseError(f"не удалось распознать сумму: «{clean_text(value)}»")
    if amount <= 0:
        raise ParseError(f"сумма меньше или равна нулю ({to_text(value)}) — "
                         "возврат или ошибка?")
    return amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def parse_qty(value) -> int:
    if value is None or (isinstance(value, str) and not value.strip()):
        raise ParseError("не указано количество")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        qty = value
    else:
        text = clean_text(value).lower()
        text = re.sub(r"\s*(шт|штук|штуки|ед)\.?$", "", text).replace(",", ".")
        try:
            qty = float(text)
        except ValueError:
            raise ParseError(f"не удалось распознать количество: «{clean_text(value)}»")
    if qty != int(qty) or qty <= 0:
        raise ParseError(f"количество должно быть целым и больше нуля: «{value}»")
    return int(qty)


# ======================================================== чтение файлов ===
def find_header(ws) -> tuple[int, dict[str, int]] | None:
    """Ищет строку шапки: в ней должны найтись все обязательные колонки."""
    for row_idx, row in enumerate(
            ws.iter_rows(min_row=1, max_row=HEADER_SCAN_ROWS, values_only=True), 1):
        mapping = {}
        for col_idx, value in enumerate(row):
            fld = ALIAS_INDEX.get(norm_header(value))
            if fld and fld not in mapping:
                mapping[fld] = col_idx
        if all(f in mapping for f in REQUIRED):
            return row_idx, mapping
    return None


def read_file(path: Path, sales: list[Sale], problems: list[Problem], stats: Counter):
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:  # битый файл, пароль, не тот формат
        problems.append(Problem(path.name, "", None, f"файл не открылся: {exc}"))
        return
    try:
        for ws in wb.worksheets:
            header = find_header(ws)
            if header is None:
                if ws.max_row and ws.max_row > 1:
                    problems.append(Problem(
                        path.name, ws.title, None,
                        "лист пропущен: не найдена шапка с колонками "
                        + ", ".join(FIELD_TITLES[f] for f in REQUIRED)))
                continue
            stats["sheets"] += 1
            header_row, mapping = header
            for row_idx, row in enumerate(
                    ws.iter_rows(min_row=header_row + 1, values_only=True),
                    header_row + 1):
                raw = {fld: (row[col] if col < len(row) else None)
                       for fld, col in mapping.items()}
                if all(clean_text(v) == "" for v in row):
                    stats["empty"] += 1
                    continue
                stats["rows"] += 1
                where = (path.name, ws.title, row_idx)
                if any(isinstance(v, str) and TOTAL_ROW_RE.match(norm_header(v)) for v in row):
                    problems.append(Problem(*where, "строка «Итого» из исходного "
                                            "файла — не учитывается", raw))
                    continue
                errors = []
                values = {}
                for fld, parser in (("date", parse_date), ("amount", parse_amount)):
                    try:
                        values[fld] = parser(raw.get(fld))
                    except ParseError as exc:
                        errors.append(str(exc))
                if "qty" in mapping:
                    try:
                        values["qty"] = parse_qty(raw.get("qty"))
                    except ParseError as exc:
                        errors.append(str(exc))
                else:
                    values["qty"] = 1
                for fld, msg in (("manager", "не указан менеджер"),
                                 ("product", "не указан товар")):
                    values[fld] = clean_text(raw.get(fld))
                    if not values[fld]:
                        errors.append(msg)
                values["city"] = clean_text(raw.get("city")) or NO_CITY
                if errors:
                    problems.append(Problem(*where, "; ".join(errors), raw))
                    continue
                sales.append(Sale(source=f"{path.name} / {ws.title} / строка {row_idx}",
                                  raw=raw, **values))
    finally:
        wb.close()


# ============================================================== очистка ===
def unify_names(sales: list[Sale]):
    """«иванов сергей», «Иванов  Сергей » -> одно написание (самое частое)."""
    for attr in ("manager", "city", "product"):
        variants = defaultdict(Counter)
        for s in sales:
            variants[getattr(s, attr).lower().replace("ё", "е")][getattr(s, attr)] += 1
        best = {}
        for key, counter in variants.items():
            name = counter.most_common(1)[0][0]
            if attr != "product" and (name.islower() or name.isupper()):
                name = name.title()
            best[key] = name
        for s in sales:
            setattr(s, attr, best[getattr(s, attr).lower().replace("ё", "е")])


def remove_duplicates(sales: list[Sale], problems: list[Problem]) -> list[Sale]:
    """Полностью совпадающие продажи (дата, менеджер, город, товар, кол-во,
    сумма) оставляем один раз, повторы — в «Проблемы» со ссылкой на оригинал."""
    seen: dict[tuple, Sale] = {}
    unique = []
    for s in sales:
        key = (s.date, s.manager, s.city, s.product, s.qty, s.amount)
        if key in seen:
            file, sheet, row = s.source.split(" / ")
            problems.append(Problem(
                file, sheet, int(row.split()[-1]),
                f"дубль: такая же продажа уже есть ({seen[key].source}) — "
                "учтена один раз", s.raw))
        else:
            seen[key] = s
            unique.append(s)
    return unique


# ========================================================== запись Excel ===
def style_header(ws, row=1, ncols=None):
    ncols = ncols or ws.max_column
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 32


def style_body(ws, first, last, formats: dict[int, str]):
    for r in range(first, last + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            if c in formats:
                cell.number_format = formats[c]


def style_total(ws, row, formats: dict[int, str]):
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = TOTAL_FILL
        cell.font = TOTAL_FONT
        cell.border = Border(top=TOTAL_BORDER.top, left=THIN, right=THIN, bottom=THIN)
        if c in formats:
            cell.number_format = formats[c]


def set_widths(ws, widths: dict[str, float]):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def page_setup(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"


def color_scale(ws, ref):
    ws.conditional_formatting.add(ref, ColorScaleRule(
        start_type="min", start_color="F8696B",
        mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="63BE7B"))


def month_label(d: date) -> str:
    return f"{MONTHS_NOM[d.month - 1]} {d.year}"


def money_float(x: Decimal) -> float:
    return float(x.quantize(Decimal("0.01")))


def sheet_all_sales(wb, sales: list[Sale]):
    ws = wb.active
    ws.title = "Все продажи"
    ws.append(["Дата", "Месяц", "Город", "Менеджер", "Товар", "Кол-во",
               "Сумма", "Откуда строка (файл / лист / строка)"])
    for s in sales:
        ws.append([s.date, month_label(s.date), s.city, s.manager, s.product,
                   s.qty, money_float(s.amount), s.source])
    last = ws.max_row
    style_header(ws)
    style_body(ws, 2, last, {1: DATE_FMT, 6: INT, 7: RUB})
    for r in range(2, last + 1):
        ws.cell(row=r, column=8).font = Font(color="808080", size=9)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{last}"
    if last >= 2:
        color_scale(ws, f"G2:G{last}")
    set_widths(ws, {"A": 12, "B": 15, "C": 18, "D": 22, "E": 42, "F": 9,
                    "G": 16, "H": 52})
    page_setup(ws)


def sheet_managers(wb, sales: list[Sale], grand: Decimal):
    ws = wb.create_sheet("По менеджерам")
    ws.append(["Менеджер", "Город", "Продаж", "Товаров, шт.", "Выручка",
               "Средний чек", "Доля в выручке"])
    groups = defaultdict(list)
    for s in sales:
        groups[s.manager].append(s)
    rows = sorted(groups.items(), key=lambda kv: -sum(x.amount for x in kv[1]))
    for name, items in rows:
        total = sum((x.amount for x in items), Decimal(0))
        ws.append([name, ", ".join(sorted({x.city for x in items})), len(items),
                   sum(x.qty for x in items), money_float(total),
                   money_float(total / len(items)), float(total / grand) if grand else 0])
    last = ws.max_row
    ws.append(["ИТОГО", "", len(sales), sum(s.qty for s in sales), money_float(grand),
               money_float(grand / len(sales)) if sales else 0, 1 if sales else 0])
    fmts = {3: INT, 4: INT, 5: RUB, 6: RUB, 7: PCT}
    style_header(ws)
    style_body(ws, 2, last, fmts)
    style_total(ws, last + 1, fmts)
    if last >= 2:
        color_scale(ws, f"E2:E{last}")
        ws.conditional_formatting.add(f"G2:G{last}", DataBarRule(
            start_type="num", start_value=0, end_type="num", end_value=1,
            color="5B9BD5"))
    ws.freeze_panes = "A2"
    set_widths(ws, {"A": 24, "B": 18, "C": 10, "D": 13, "E": 18, "F": 16, "G": 14})
    page_setup(ws)


def sheet_months(wb, sales: list[Sale], grand: Decimal):
    ws = wb.create_sheet("По месяцам")
    ws.append(["Месяц", "Продаж", "Товаров, шт.", "Выручка", "Средний чек",
               "К прошлому месяцу"])
    groups = defaultdict(list)
    for s in sales:
        groups[(s.date.year, s.date.month)].append(s)
    prev = None
    for (y, m) in sorted(groups):
        items = groups[(y, m)]
        total = sum((x.amount for x in items), Decimal(0))
        growth = float(total / prev - 1) if prev else None
        ws.append([month_label(date(y, m, 1)), len(items), sum(x.qty for x in items),
                   money_float(total), money_float(total / len(items)), growth])
        prev = total
    last = ws.max_row
    ws.append(["ИТОГО", len(sales), sum(s.qty for s in sales), money_float(grand),
               money_float(grand / len(sales)) if sales else 0, None])
    fmts = {2: INT, 3: INT, 4: RUB, 5: RUB, 6: '+0.0%;[Red]-0.0%;0.0%'}
    style_header(ws)
    style_body(ws, 2, last, fmts)
    style_total(ws, last + 1, fmts)
    if last >= 2:
        color_scale(ws, f"D2:D{last}")
    ws.freeze_panes = "A2"
    set_widths(ws, {"A": 16, "B": 10, "C": 13, "D": 18, "E": 16, "F": 14})
    page_setup(ws)

    if last >= 2:
        chart = BarChart()
        chart.type = "col"
        chart.title = "Выручка по месяцам, ₽"
        chart.style = 10
        chart.legend = None
        chart.height, chart.width = 9, 18
        chart.add_data(Reference(ws, min_col=4, min_row=1, max_row=last),
                       titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=last))
        chart.y_axis.numFmt = '#,##0'
        chart.y_axis.majorGridlines = None
        chart.x_axis.delete = False   # без этого новые Excel прячут оси
        chart.y_axis.delete = False
        chart.series[0].graphicalProperties.solidFill = "2E75B6"
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.dataLabels.numFmt = '#,##0'
        ws.add_chart(chart, "H2")


def sheet_products(wb, sales: list[Sale], grand: Decimal):
    ws = wb.create_sheet("Топ товаров")
    ws.append(["Место", "Товар", "Продано, шт.", "Продаж", "Выручка",
               "Средняя цена за шт.", "Доля в выручке"])
    groups = defaultdict(list)
    for s in sales:
        groups[s.product].append(s)
    rows = sorted(groups.items(), key=lambda kv: -sum(x.amount for x in kv[1]))
    for place, (name, items) in enumerate(rows, 1):
        total = sum((x.amount for x in items), Decimal(0))
        qty = sum(x.qty for x in items)
        ws.append([place, name, qty, len(items), money_float(total),
                   money_float(total / qty), float(total / grand) if grand else 0])
    last = ws.max_row
    ws.append(["", "ИТОГО", sum(s.qty for s in sales), len(sales), money_float(grand),
               None, 1 if sales else 0])
    fmts = {3: INT, 4: INT, 5: RUB, 6: RUB, 7: PCT}
    style_header(ws)
    style_body(ws, 2, last, fmts)
    style_total(ws, last + 1, fmts)
    for r in range(2, min(last, 4) + 1):          # тройка лидеров
        ws.cell(row=r, column=1).font = Font(bold=True, color="C55A11")
        ws.cell(row=r, column=2).font = Font(bold=True)
    for r in range(2, last + 1):
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
    if last >= 2:
        ws.conditional_formatting.add(f"E2:E{last}", DataBarRule(
            start_type="num", start_value=0, end_type="max", color="63BE7B"))
    ws.freeze_panes = "A2"
    set_widths(ws, {"A": 8, "B": 44, "C": 13, "D": 10, "E": 18, "F": 18, "G": 14})
    page_setup(ws)


def sheet_problems(wb, problems: list[Problem]):
    ws = wb.create_sheet("Проблемы")
    fields = ["date", "manager", "city", "product", "qty", "amount"]
    ws.append(["Файл", "Лист", "Строка", "Причина"]
              + [f"{FIELD_TITLES[f]} (как в файле)" for f in fields])
    for p in problems:
        ws.append([p.file, p.sheet, p.row, p.reason[:1].upper() + p.reason[1:]]
                  + [to_text(p.raw.get(f)) for f in fields])
    if not problems:
        ws.append(["Проблем не найдено — все строки разобраны."])
    last = ws.max_row
    style_header(ws)
    for c in range(1, 5):
        ws.cell(row=1, column=c).fill = PROBLEM_FILL
    style_body(ws, 2, last, {})
    for r in range(2, last + 1):
        ws.cell(row=r, column=4).alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{last}"
    set_widths(ws, {"A": 30, "B": 18, "C": 8, "D": 60, "E": 16, "F": 20, "G": 16,
                    "H": 36, "I": 10, "J": 18})
    page_setup(ws)
    if problems:
        ws.sheet_properties.tabColor = "C00000"


def build_report(sales: list[Sale], problems: list[Problem], out: Path) -> Decimal:
    grand = sum((s.amount for s in sales), Decimal(0))
    wb = Workbook()
    sheet_all_sales(wb, sales)
    sheet_managers(wb, sales, grand)
    sheet_months(wb, sales, grand)
    sheet_products(wb, sales, grand)
    sheet_problems(wb, problems)
    wb.save(out)
    return grand


def fmt_rub(x: Decimal) -> str:
    return f"{x:,.2f}".replace(",", " ").replace(".", ",") + " руб."


# ================================================================= main ===
def main(argv=None) -> int:
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Сводный отчёт по продажам из Excel-файлов")
    ap.add_argument("--input", default=str(BASE_DIR / "input"),
                    help="папка с выгрузками .xlsx (по умолчанию input рядом со скриптом)")
    ap.add_argument("--output", default=str(BASE_DIR / "report.xlsx"),
                    help="куда сохранить отчёт (по умолчанию report.xlsx рядом со скриптом)")
    args = ap.parse_args(argv)

    in_dir, out = Path(args.input), Path(args.output)
    if not in_dir.is_dir():
        print(f"Папка не найдена: {in_dir}")
        return 1
    files = sorted(p for p in in_dir.glob("*.xlsx")
                   if not p.name.startswith("~$") and p.resolve() != out.resolve())
    if not files:
        print(f"В папке {in_dir} нет файлов .xlsx. Положите туда выгрузки и запустите снова.")
        return 1

    sales: list[Sale] = []
    problems: list[Problem] = []
    stats = Counter()
    for path in files:
        before_s, before_p = len(sales), len(problems)
        read_file(path, sales, problems, stats)
        print(f"  {path.name}: разобрано строк {len(sales) - before_s}, "
              f"не разобрано {len(problems) - before_p}")

    unify_names(sales)
    sales = remove_duplicates(sales, problems)
    sales.sort(key=lambda s: (s.date, s.city, s.manager))
    problems.sort(key=lambda p: (p.file, p.sheet, p.row or 0))

    try:
        grand = build_report(sales, problems, out)
    except PermissionError:
        print(f"\nНе удалось сохранить {out.name}: файл открыт в Excel. "
              "Закройте его и запустите снова.")
        return 1

    dups = sum(1 for p in problems if p.reason.startswith("дубль"))
    print()
    print(f"Файлов: {len(files)}, листов с данными: {stats['sheets']}, "
          f"строк прочитано: {stats['rows']} (пустых пропущено: {stats['empty']})")
    print(f"В отчёт попало продаж: {len(sales)} на сумму {fmt_rub(grand)}")
    print(f"Убрано дублей: {dups}. Всего строк на листе «Проблемы»: {len(problems)}")
    if sales:
        print(f"Период: {min(s.date for s in sales):%d.%m.%Y} – "
              f"{max(s.date for s in sales):%d.%m.%Y}")
    print(f"Готово: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

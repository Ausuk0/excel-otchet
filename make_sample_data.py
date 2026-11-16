"""
Генератор тестовых данных для демо.

Создаёт в папке input/ четыре Excel-файла «как от разных филиалов» —
у каждого своя шапка, свой порядок колонок и свои «грязные» привычки:
лишние пробелы, даты текстом в разных форматах, суммы текстом
с пробелами и запятой, дубли, строки «Итого», пустые строки и т.п.

В конце печатает контрольные цифры: сколько чистых продаж и на какую
сумму должно попасть в отчёт и сколько строк должно уйти в «Проблемы».

Запуск:  python make_sample_data.py
"""

from __future__ import annotations

import random
import sys
from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"

SEED = 2026
PERIOD_START = date(2026, 7, 1)
PERIOD_END = date(2026, 10, 31)

# Товар -> базовая цена, руб.
PRODUCTS = {
    "Кофемашина DeLonghi Magnifica S": 42990,
    "Чайник электрический Redmond RK-G178": 3490,
    "Пылесос Dyson V12 Detect Slim": 54990,
    "Робот-пылесос Xiaomi Robot Vacuum S10": 21990,
    "Микроволновая печь Samsung ME88SUG": 11490,
    "Блендер Philips HR2600": 4290,
    "Утюг Tefal FV6872": 7990,
    "Мультиварка Redmond RMC-M90": 6490,
    "Холодильник Атлант ХМ 4624-101": 46990,
    "Стиральная машина LG F2J3NS0W": 38990,
    "Телевизор Haier 55 Smart TV S1": 44990,
    "Наушники Sony WH-1000XM5": 29990,
    "Фен Polaris PHD 2077i": 2990,
    "Электрогриль Tefal Optigrill+ GC712D34": 15990,
    "Увлажнитель воздуха Ballu UHB-400": 3190,
}
# «Ходовые» товары продаются чаще.
PRODUCT_WEIGHTS = [6, 10, 3, 5, 6, 8, 7, 6, 2, 3, 3, 4, 9, 4, 7]

BRANCHES = [
    {"file": "Москва_продажи.xlsx", "city": "Москва",
     "managers": ["Иванов Сергей", "Петрова Анна"], "n": (45, 55)},
    {"file": "СПб_выгрузка_Смирнов.xlsx", "city": "Санкт-Петербург",
     "managers": ["Смирнов Алексей"], "n": (50, 60)},
    {"file": "Казань_отчёт менеджеров.xlsx", "city": "Казань",
     "managers": ["Галиева Алсу", "Хабибуллин Рустам"], "n": (30, 40)},
    {"file": "Екатеринбург_Кузнецов.xlsx", "city": "Екатеринбург",
     "managers": ["Кузнецов Дмитрий"], "n": (45, 55)},
]

MONTHS_GEN = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
              "августа", "сентября", "октября", "ноября", "декабря"]

NBSP = " "


def money(x) -> Decimal:
    return Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def rub_text(amount: Decimal, sep: str = " ", with_kop: bool = True) -> str:
    """12500.5 -> '12 500,50' (как в русском Excel, скопированном в текст)."""
    whole, kop = f"{amount:.2f}".split(".")
    groups = []
    while whole:
        groups.insert(0, whole[-3:])
        whole = whole[:-3]
    text = sep.join(groups)
    if with_kop:
        text += "," + kop
    return text


def make_sales(rng: random.Random) -> list[dict]:
    """Чистые (эталонные) продажи по всем филиалам — без дублей."""
    sales = []
    seen = set()
    days = (PERIOD_END - PERIOD_START).days
    names = list(PRODUCTS)
    for br in BRANCHES:
        for manager in br["managers"]:
            n = rng.randint(*br["n"])
            made = 0
            while made < n:
                d = PERIOD_START + timedelta(days=rng.randint(0, days))
                product = rng.choices(names, weights=PRODUCT_WEIGHTS)[0]
                qty = rng.choices([1, 2, 3, 4], weights=[70, 20, 7, 3])[0]
                discount = rng.choice([0, 0, 0, 0, 3, 5, 7, 10])
                amount = money(PRODUCTS[product] * qty * (100 - discount) / 100)
                key = (d, manager, product, qty, amount)
                if key in seen:  # не допускаем случайных «настоящих» дублей
                    continue
                seen.add(key)
                sales.append({"date": d, "manager": manager, "city": br["city"],
                              "product": product, "qty": qty, "amount": amount,
                              "file": br["file"]})
                made += 1
    sales.sort(key=lambda s: (s["date"], s["manager"]))
    return sales


class Stats:
    def __init__(self):
        self.clean = 0
        self.total = Decimal("0")
        self.problems = 0
        self.notes = []

    def ok(self, s):
        self.clean += 1
        self.total += s["amount"]

    def problem(self, note):
        self.problems += 1
        self.notes.append(note)


def bold_header(ws, row=1):
    for c in ws[row]:
        c.font = Font(bold=True)


def noisy_name(rng, name, chance=0.25):
    """Лишние пробелы / другой регистр — как вводят руками."""
    if rng.random() > chance:
        return name
    variant = rng.choice(["trail", "lead", "double", "lower"])
    if variant == "trail":
        return name + "  "
    if variant == "lead":
        return " " + name
    if variant == "double":
        return name.replace(" ", "   ", 1)
    return name.lower()


# ---------------------------------------------------------------- Москва ---
def write_moscow(rng, sales, st: Stats, path: Path):
    """Аккуратная выгрузка: настоящие даты, числа. Но: пробелы в ФИО,
    часть дат текстом, 2 дубля, строка «Итого», одна сумма «уточнить»."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Продажи"
    ws.append(["Дата", "Город", "Менеджер", "Товар", "Количество", "Сумма"])
    bold_header(ws)
    dup_idx = set(rng.sample(range(len(sales)), 2))
    total = Decimal("0")
    for i, s in enumerate(sales):
        d = s["date"].strftime("%d.%m.%Y") if rng.random() < 0.2 else s["date"]
        row = [d, s["city"], noisy_name(rng, s["manager"]), s["product"],
               s["qty"], float(s["amount"])]
        ws.append(row)
        st.ok(s)
        total += s["amount"]
        if i in dup_idx:  # менеджер случайно вставил строку второй раз
            ws.append(list(row))
            st.problem("Москва: дубль строки")
            total += s["amount"]
        if i == len(sales) // 2:
            ws.append([s["date"], s["city"], s["manager"],
                       "Пылесос Dyson V12 Detect Slim", 1, "уточнить у клиента"])
            st.problem("Москва: сумма «уточнить у клиента»")
    ws.append([None, None, "Итого", None, None, float(total)])
    st.problem("Москва: строка «Итого»")
    for r in ws.iter_rows(min_row=2, max_col=1):
        if isinstance(r[0].value, date):
            r[0].number_format = "DD.MM.YYYY"
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 40
    wb.save(path)


# ------------------------------------------------------- Санкт-Петербург ---
def write_spb(rng, sales, st: Stats, path: Path):
    """Шапка не в первой строке, другие названия колонок, даты в ISO-тексте,
    суммы текстом «12 500,00», ФИО в разном регистре, двойные пробелы
    в названиях товаров. Проблемы: дата «вчера», пустой товар, 1 дубль."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Лист1"
    ws.append(["Отчёт о продажах филиала Санкт-Петербург, июль–октябрь 2026"])
    ws.append([])
    ws.append(["№", "Дата продажи", "Филиал", "Сотрудник",
               "Наименование товара", "Кол-во", "Сумма, руб."])
    bold_header(ws, 3)
    dup_at = rng.randrange(len(sales))
    num = 0
    for i, s in enumerate(sales):
        num += 1
        sep = rng.choice([" ", NBSP, ""])
        amount = rub_text(s["amount"], sep=sep)
        if rng.random() < 0.15:
            amount = " " + amount + " "
        product = s["product"].replace(" ", "  ", 1) if rng.random() < 0.2 else s["product"]
        row = [num, s["date"].strftime("%Y-%m-%d"), s["city"],
               noisy_name(rng, s["manager"], 0.4), product, s["qty"], amount]
        ws.append(row)
        st.ok(s)
        if i == dup_at:
            num += 1
            ws.append([num] + row[1:])
            st.problem("СПб: дубль строки")
        if i == 10:
            num += 1
            ws.append([num, "вчера", s["city"], s["manager"],
                       "Фен Polaris PHD 2077i", 1, "2 990,00"])
            st.problem("СПб: дата «вчера»")
        if i == 30:
            num += 1
            ws.append([num, s["date"].strftime("%Y-%m-%d"), s["city"],
                       s["manager"], "   ", 2, "6 980,00"])
            st.problem("СПб: пустой товар")
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 40
    wb.save(path)


# ---------------------------------------------------------------- Казань ---
def write_kazan(rng, sales, st: Stats, path: Path):
    """Другой порядок колонок; даты — то числом Excel (45870), то
    «5 августа 2026», то нормальной датой; суммы «12 500 руб.» и числа;
    пустые строки посередине. Проблемы: 31.09.2026, нет менеджера, минус."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Продажи Казань"
    ws.append(["Товар", "Кол-во", "Сумма", "Менеджер", "Дата", "Город"])
    bold_header(ws)
    for i, s in enumerate(sales):
        kind = rng.choice(["serial", "words", "date"])
        if kind == "serial":
            d = (s["date"] - date(1899, 12, 30)).days
        elif kind == "words":
            d = f"{s['date'].day} {MONTHS_GEN[s['date'].month - 1]} {s['date'].year}"
        else:
            d = s["date"]
        if rng.random() < 0.5:
            amount = rub_text(s["amount"], sep=NBSP, with_kop=s["amount"] % 1 != 0) + " руб."
        else:
            amount = float(s["amount"])
        ws.append([s["product"], s["qty"], amount, noisy_name(rng, s["manager"]),
                   d, " " + s["city"] if rng.random() < 0.3 else s["city"]])
        st.ok(s)
        if i in (7, 21):
            ws.append([])  # пустые строки посередине
        if i == 12:
            ws.append(["Утюг Tefal FV6872", 1, "7 990 руб.", s["manager"],
                       "31.09.2026", s["city"]])
            st.problem("Казань: несуществующая дата 31.09.2026")
        if i == 18:
            ws.append(["Блендер Philips HR2600", 1, 4290, "", s["date"], s["city"]])
            st.problem("Казань: не указан менеджер")
        if i == 25:
            ws.append(["Мультиварка Redmond RMC-M90", 1, -6490, s["manager"],
                       s["date"], s["city"]])
            st.problem("Казань: отрицательная сумма (возврат)")
    for r in ws.iter_rows(min_row=2, min_col=5, max_col=5):
        if isinstance(r[0].value, date):
            r[0].number_format = "DD.MM.YYYY"
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["D"].width = 20
    wb.save(path)


# ---------------------------------------------------------- Екатеринбург ---
def write_ekb(rng, sales, st: Stats, path: Path):
    """Два листа (по два месяца), даты «05/08/26» и «05.08.2026 14:35»,
    количество «2 шт.», суммы «12500.50» текстом. Проблемы: строка
    «ИТОГО:» на каждом листе, количество «два», строка скопирована
    с первого листа на второй."""
    wb = Workbook()
    sheets = {
        "Июль-Август": [s for s in sales if s["date"].month in (7, 8)],
        "Сентябрь-Октябрь": [s for s in sales if s["date"].month in (9, 10)],
    }
    copied = rng.choice(sheets["Июль-Август"])
    first = True
    for title, rows in sheets.items():
        ws = wb.active if first else wb.create_sheet()
        ws.title = title
        ws.append(["Дата", "Менеджер", "Товар", "Кол-во", "Сумма продажи", "Город"])
        bold_header(ws)
        total = Decimal("0")
        for i, s in enumerate(rows):
            if rng.random() < 0.5:
                d = s["date"].strftime("%d/%m/%y")
            else:
                t = f"{rng.randint(10, 19)}:{rng.randint(0, 59):02d}"
                d = s["date"].strftime("%d.%m.%Y") + " " + t
            qty = f"{s['qty']} шт." if rng.random() < 0.6 else s["qty"]
            amount = f"{s['amount']:.2f}" if rng.random() < 0.6 else float(s["amount"])
            ws.append([d, noisy_name(rng, s["manager"]), s["product"], qty,
                       amount, s["city"]])
            st.ok(s)
            total += s["amount"]
            if not first and i == 5:
                ws.append([copied["date"].strftime("%d/%m/%y"), copied["manager"],
                           copied["product"], copied["qty"],
                           f"{copied['amount']:.2f}", copied["city"]])
                st.problem("Екатеринбург: строка продублирована на втором листе")
                total += copied["amount"]
            if first and i == 15:
                ws.append([s["date"].strftime("%d/%m/%y"), s["manager"],
                           "Чайник электрический Redmond RK-G178", "два",
                           "6980.00", s["city"]])
                st.problem("Екатеринбург: количество «два»")
        ws.append([])
        ws.append(["ИТОГО:", None, None, None, float(total), None])
        st.problem(f"Екатеринбург: строка «ИТОГО:» на листе {title}")
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 40
        first = False
    wb.save(path)


WRITERS = [write_moscow, write_spb, write_kazan, write_ekb]


def main():
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    rng = random.Random(SEED)
    INPUT_DIR.mkdir(exist_ok=True)
    sales = make_sales(rng)
    st = Stats()
    for br, writer in zip(BRANCHES, WRITERS):
        part = [s for s in sales if s["file"] == br["file"]]
        path = INPUT_DIR / br["file"]
        writer(rng, part, st, path)
        print(f"Создан файл: input/{br['file']}  ({len(part)} продаж)")

    print()
    print("Контрольные цифры (что должен показать отчёт):")
    print(f"  чистых продаж:          {st.clean}")
    print(f"  общая сумма:            {rub_text(st.total)} руб.")
    print(f"  строк в «Проблемах»:    {st.problems}")
    for note in st.notes:
        print(f"    - {note}")


if __name__ == "__main__":
    main()

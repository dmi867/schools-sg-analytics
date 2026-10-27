#!/usr/bin/env python3
"""Создаёт Excel-шаблон (одна таблица) для план/факт СГ и платежей.

Выход: data/templates/шаблон_СГ_и_платежи.xlsx

В одной строке можно заполнить СГ и/или платёж на дату:
  • есть Факт % → точка стройготовности
  • есть Сумма → платёж (Тип: Предоплата / Основной платеж)

  python3 scripts/make_sg_pay_template.py
  python3 scripts/import_sg_pay_template.py data/templates/шаблон_СГ_и_платежи.xlsx
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).parent.parent
OUT = ROOT / "data" / "templates" / "шаблон_СГ_и_платежи.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="143260")
HEADER_FONT = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
EXAMPLE_FILL = PatternFill("solid", fgColor="F3FBFC")
THIN = Border(
    left=Side(style="thin", color="C5D0DC"),
    right=Side(style="thin", color="C5D0DC"),
    top=Side(style="thin", color="C5D0DC"),
    bottom=Side(style="thin", color="C5D0DC"),
)

HEADERS = [
    "УИН",
    "Дата",
    "План %",
    "Факт %",
    "Сумма платежа, руб.",
    "Тип платежа",
    "Комментарий",
]


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Данные"

    for i, h in enumerate(HEADERS, 1):
        cell = ws.cell(1, i, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = THIN
    ws.row_dimensions[1].height = 32

    # примеры: СГ-строки, платёж, и строка где и то и другое
    examples = [
        ("1000001228.1000001021", "03.10.2025", None, None, 195594900, "Предоплата", "пример: только платёж — удалить"),
        ("1000001228.1000001021", "01.12.2025", 22.5, 10.2, None, None, "пример: только СГ — удалить"),
        ("1000001228.1000001021", "15.01.2026", 35.0, 28.0, 45000000, "Основной платеж", "пример: СГ + платёж в один день — удалить"),
        ("1000001228.1000001021", "01.03.2026", 50.0, 44.5, None, None, "пример — удалить"),
    ]
    for r, row in enumerate(examples, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.fill = EXAMPLE_FILL
            cell.border = THIN
            if c in (3, 4) and v is not None:
                cell.number_format = "0.0"
            if c == 5 and v is not None:
                cell.number_format = "#,##0.00"

    for r in range(6, 60):
        for c in range(1, len(HEADERS) + 1):
            ws.cell(r, c).border = THIN

    dv = DataValidation(
        type="list",
        formula1='"Предоплата,Основной платеж"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Тип платежа",
        error="Укажите «Предоплата» или «Основной платеж» (если есть сумма).",
    )
    dv.add("F2:F1000")
    ws.add_data_validation(dv)

    widths = [24, 12, 10, 10, 18, 18, 36]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = "A1:G59"
    ws.freeze_panes = "A2"

    # краткая подсказка второй строкой на отдельном узком листе не нужна —
    # комментарий в шапке через note на A1 неудобен; добавим лист «Как заполнять» в 5 строк
    ws2 = wb.create_sheet("Как заполнять", 0)
    ws2["A1"] = "Одна таблица: СГ и платежи"
    ws2["A1"].font = Font(bold=True, size=13, color="143260")
    tips = [
        "Заполняйте лист «Данные». В каждой строке — УИН + Дата.",
        "Стройготовность: укажите Факт % (План % желателен). Можно без платежа.",
        "Платёж: укажите Сумму и Тип (Предоплата / Основной платеж). Можно без СГ.",
        "В один день можно заполнить и СГ, и платёж в одной строке.",
        "УИН — полный код, как в Акценте. Дата — ДД.ММ.ГГГГ. Сумма — в рублях.",
        "Строки «пример — удалить» перед загрузкой уберите.",
        "",
        "Импорт: python3 scripts/import_sg_pay_template.py data/templates/шаблон_СГ_и_платежи.xlsx",
        "Затем: python3 build_html.py",
    ]
    for i, t in enumerate(tips, 3):
        ws2.cell(i, 1, t)
    ws2.column_dimensions["A"].width = 95

    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"OK: {OUT}")


if __name__ == "__main__":
    main()

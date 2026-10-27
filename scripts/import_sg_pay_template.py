#!/usr/bin/env python3
"""Импорт шаблона (одна таблица) в data/per_object.json.

Лист «Данные»: УИН | Дата | План % | Факт % | Сумма платежа, руб. | Тип платежа | …

  python3 scripts/import_sg_pay_template.py data/templates/шаблон_СГ_и_платежи.xlsx
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent.parent
OUT_PATH = ROOT / "data" / "per_object.json"

DATA_SHEETS = ("Данные", "СГ")  # «СГ» — на случай старого двухлистного файла


def parse_date(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if not v:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y"):
        try:
            return datetime.strptime(str(v).strip(), fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None


def parse_num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace("\xa0", "").replace(" ", "").replace(",", "."))
    except ValueError:
        return None


def colmap(ws):
    hdr = [str(c.value).strip() if c.value is not None else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
    return {h: i for i, h in enumerate(hdr) if h}, hdr


def pick_sheet(wb):
    for name in DATA_SHEETS:
        if name in wb.sheetnames:
            return wb[name]
    # fallback: first sheet that looks like the template
    for name in wb.sheetnames:
        ws = wb[name]
        col, _ = colmap(ws)
        if "УИН" in col and "Дата" in col:
            return ws
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", type=Path)
    ap.add_argument("-o", "--out", type=Path, default=OUT_PATH)
    args = ap.parse_args()

    if not args.xlsx.exists():
        sys.exit(f"File not found: {args.xlsx}")

    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    ws = pick_sheet(wb)
    if ws is None:
        sys.exit("Не найден лист с колонками УИН и Дата")

    col, hdr = colmap(ws)
    if "УИН" not in col or "Дата" not in col:
        sys.exit(f"Нужны колонки УИН и Дата. Есть: {hdr}")

    fact_i = col.get("Факт %")
    plan_i = col.get("План %")
    # new single-table name, or old «Сумма, руб.» / «Дата платежа» layouts
    sum_i = col.get("Сумма платежа, руб.") or col.get("Сумма, руб.")
    typ_i = col.get("Тип платежа")
    comment_i = col.get("Комментарий") or col.get("Комментарий (необязательно)")

    by_uin = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        uin = str(row[col["УИН"]] or "").strip()
        if not uin:
            continue
        comment = str(row[comment_i] or "") if comment_i is not None else ""
        if "пример" in comment.lower():
            continue

        d = parse_date(row[col["Дата"]])
        if not d:
            continue

        fact = parse_num(row[fact_i]) if fact_i is not None else None
        plan = parse_num(row[plan_i]) if plan_i is not None else None
        amt = parse_num(row[sum_i]) if sum_i is not None else None
        typ = str(row[typ_i] or "").strip() if typ_i is not None else ""

        if fact is None and amt is None:
            continue

        rec = by_uin.setdefault(uin, {"series": [], "pays": []})
        if fact is not None:
            rec["series"].append({"d": d, "plan": plan, "fact": fact})
        if amt is not None:
            if typ not in ("Предоплата", "Основной платеж"):
                print(f"WARN {uin} {d}: тип «{typ or 'пусто'}» → Основной платеж")
                typ = "Основной платеж"
            rec["pays"].append({"d": d, "amt": amt, "advance": typ == "Предоплата"})

    # support legacy second sheet «Платежи» if present and single-table had no pays
    if "Платежи" in wb.sheetnames:
        ws2 = wb["Платежи"]
        col2, _ = colmap(ws2)
        if "УИН" in col2 and ("Дата платежа" in col2 or "Дата" in col2):
            d_key = "Дата платежа" if "Дата платежа" in col2 else "Дата"
            s_key = "Сумма, руб." if "Сумма, руб." in col2 else "Сумма платежа, руб."
            if s_key in col2 and "Тип платежа" in col2:
                for row in ws2.iter_rows(min_row=2, values_only=True):
                    uin = str(row[col2["УИН"]] or "").strip()
                    if not uin:
                        continue
                    d = parse_date(row[col2[d_key]])
                    amt = parse_num(row[col2[s_key]])
                    typ = str(row[col2["Тип платежа"]] or "").strip()
                    if not d or amt is None:
                        continue
                    rec = by_uin.setdefault(uin, {"series": [], "pays": []})
                    rec["pays"].append({"d": d, "amt": amt, "advance": typ == "Предоплата"})

    for uin, rec in by_uin.items():
        # merge series by date (last wins)
        series_by_d = {}
        for p in rec["series"]:
            series_by_d[p["d"]] = p
        rec["series"] = sorted(series_by_d.values(), key=lambda x: x["d"])
        rec["pays"].sort(key=lambda x: x["d"])

    by_uin = {u: r for u, r in by_uin.items() if r["series"]}
    if not by_uin:
        sys.exit("Нет данных СГ (нужен хотя бы Факт % по УИН). Удалите примеры и заполните строки.")

    existing = {}
    if args.out.exists():
        existing = json.loads(args.out.read_text(encoding="utf-8"))

    for uin, rec in by_uin.items():
        existing[uin] = rec
        print(f"OK {uin}: series={len(rec['series'])} pays={len(rec['pays'])}")

    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(existing, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Wrote {args.out} — всего {len(existing)} объектов (обновлено {len(by_uin)})")


if __name__ == "__main__":
    main()

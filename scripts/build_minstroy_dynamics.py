#!/usr/bin/env python3
"""Собрать динамику объектов Минстроя из папки Simple_List/ (срезы Акцента).

Исключает снос. На выходе data/minstroy_dynamics.json для dashboard.html
(см. scripts/build_minstroy_dashboard.py). Добавляет funds, marks, exp_kind.

  python3 scripts/build_minstroy_dynamics.py
  python3 scripts/build_minstroy_dashboard.py
"""
from __future__ import annotations

import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent.parent
SL_DIR = ROOT / "Simple_List"
RAW_DIR = ROOT / "data" / "raw"
OUT = ROOT / "data" / "minstroy_dynamics.json"
KSG_CANDIDATES = [
    RAW_DIR / "КСГ+Экспертиза_минстрой.xlsx",
    RAW_DIR / "1708_КСГ+Экспертиза.xlsx",
]

CORRIDOR = 10.0  # ± п.п. СГ vs оплата


def uin_key(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        # Excel часто даёт 1000002390.1 как float — сохраняем как есть через repr-safe
        s = f"{v:.10f}".rstrip("0").rstrip(".")
        return s
    return str(v).strip()


def stamp(name: str):
    m = re.match(r"(\d{2})(\d{2})", name.replace(" ", ""))
    if not m:
        return None
    d, mo = int(m.group(1)), int(m.group(2))
    y = 2026 if mo >= 5 else 2025
    try:
        return datetime(y, mo, d).date()
    except ValueError:
        return None


def fnum(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace("\xa0", "").replace(" ", "").replace(",", "."))
    except ValueError:
        return None


def parse_date(v):
    """Дата из Simple List → date | None."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if hasattr(v, "year") and hasattr(v, "month") and hasattr(v, "day"):
        return v  # already date
    s = str(v).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def fmt_date(d):
    return d.strftime("%d.%m.%Y") if d else None


def short(n, nmax=48):
    n = (n or "").replace("МБОУ ", "").replace("МАОУ ", "").replace("МОУ ", "").strip()
    return (n[: nmax - 1] + "…") if len(n) > nmax else n


def short_contractor(c):
    if not c:
        return None
    c = re.sub(
        r"(ООО|ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ|ЗАО|АО|"
        r"Закрытое акционерное общество|Акционерное общество)\s*",
        "",
        str(c),
        flags=re.I,
    ).strip(" «»\"'")
    return (c[:40] + "…") if len(c) > 40 else c or None


def work_bucket(work: str) -> str:
    w = (work or "").lower()
    if "снос" in w:
        return "снос"
    if "строитель" in w:
        return "строительство"
    if "реконструк" in w:
        return "реконструкция"
    if "капитальн" in w or "ремонт" in w:
        return "капремонт"
    return "прочее"


def is_snos(work: str, name: str = "") -> bool:
    blob = f"{work or ''} {name or ''}".lower()
    return "снос" in blob


def money_status(sg, pct):
    if sg is None or pct is None:
        return "unknown"
    gap = pct - sg
    if gap > CORRIDOR:
        return "over"
    if gap < -CORRIDOR:
        return "credit"
    return "balanced"


def classify_exp(exp_done, egrz, exp_agreed) -> str:
    """ТЧ — техчасть (есть заключение/ЕГРЗ), СД — смета (согласованная стоимость)."""
    tch = bool(exp_done or egrz)
    sd = bool(exp_agreed)
    if tch and sd:
        return "both"
    if tch:
        return "tch"
    if sd:
        return "sd"
    return ""


def fund_bucket(src: str):
    s = (src or "").lower()
    if "кик" in s or "казначейск" in s:
        return "kik"
    if "ибк" in s:
        return "ibk"
    if "федераль" in s:
        return "fed"
    if "муницип" in s:
        return "mun"
    if "московской области" in s:
        return "mo"
    return None


def axis_date(d):
    if not d:
        return None
    if isinstance(d, datetime):
        return d.strftime("%d.%m.%y")
    if hasattr(d, "year") and hasattr(d, "month") and hasattr(d, "day"):
        return f"{d.day:02d}.{d.month:02d}.{str(d.year)[-2:]}"
    s = parse_date(d)
    return s.strftime("%d.%m.%y") if s else None


def ksg_kind(proc: str):
    p = (proc or "").strip()
    if p == "Экспертиза" or p.startswith("Экспертиза "):
        return "exp"
    if p == "СМР" or p.startswith("СМР "):
        return "smr"
    if p.startswith("Начало функционирования"):
        return "open"
    return None


def load_ksg_marks(uins: set[str]):
    """Окончание плана и факта: экспертиза, СМР, открытие."""
    path = next((p for p in KSG_CANDIDATES if p.exists()), None)
    if not path:
        print("KSG marks: no file, skip")
        return {}
    print(f"KSG marks from {path.name} ...", flush=True)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    acc: dict[str, dict] = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row:
            continue
        uin = uin_key(row[0])
        if uin not in uins:
            continue
        kind = ksg_kind(str(row[6] or ""))
        if not kind:
            continue
        slot = acc.setdefault(uin, {}).setdefault(kind, {"plan": None, "fact": None})
        fact_end = parse_date(row[8] if len(row) > 8 else None)
        plan_end = parse_date(row[10] if len(row) > 10 else None)
        if fact_end and (slot["fact"] is None or fact_end > slot["fact"]):
            slot["fact"] = fact_end
        if plan_end and (slot["plan"] is None or plan_end > slot["plan"]):
            slot["plan"] = plan_end
    wb.close()
    out = {}
    for uin, kinds in acc.items():
        marks = {}
        for key in ("exp", "smr", "open"):
            slot = kinds.get(key) or {}
            plan = axis_date(slot.get("plan"))
            fact = axis_date(slot.get("fact"))
            if plan or fact:
                marks[key] = {"plan": plan, "fact": fact}
        if marks:
            out[uin] = marks
    print(f"  marks for {len(out)} UINs")
    return out


def load_funding(path: Path, uins: set[str]):
    """План и факт по источникам с последнего среза Simple List."""
    if not path or not path.exists():
        return {}
    print(f"funds from {path.name} ...", flush=True)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    headers = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    col = {h: i for i, h in enumerate(headers) if h}
    need = [
        "Код УИН",
        "Наименование источника финансирования",
        "Сумма планового финансирования",
        "Сумма фактического финансирования",
        "Год финансирования",
    ]
    if any(n not in col for n in need):
        wb.close()
        print("  skip funds: missing columns")
        return {}
    seen = set()
    acc: dict[str, dict] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        uin = uin_key(row[col["Код УИН"]])
        if uin not in uins:
            continue
        bucket = fund_bucket(str(row[col["Наименование источника финансирования"]] or ""))
        if not bucket:
            continue
        plan = fnum(row[col["Сумма планового финансирования"]]) or 0.0
        fact = fnum(row[col["Сумма фактического финансирования"]]) or 0.0
        year = row[col["Год финансирования"]]
        key = (uin, bucket, year, round(plan, 2), round(fact, 2))
        if key in seen:
            continue
        seen.add(key)
        slot = acc.setdefault(uin, {}).setdefault(bucket, {"plan": 0.0, "fact": 0.0})
        slot["plan"] += plan
        slot["fact"] += fact
    wb.close()
    out = {}
    for uin, buckets in acc.items():
        funds = []
        for key in ("fed", "mo", "mun", "ibk", "kik"):
            slot = buckets.get(key)
            if not slot:
                continue
            plan = round(slot["plan"] / 1e6, 1)
            fact = round(slot["fact"] / 1e6, 1)
            if plan or fact:
                funds.append({"k": key, "plan": plan, "fact": fact})
        if funds:
            out[uin] = funds
    print(f"  funds for {len(out)} UINs")
    return out


def pay_pct_of(paid, limit, nmck, field_pct=None):
    """% оплаты: поле «Процент выплат» (= paid/предельная), иначе paid/лимит, иначе paid/НМЦК."""
    if field_pct is not None:
        return round(field_pct, 1)
    base = limit if limit and limit > 0 else (nmck if nmck and nmck > 0 else None)
    if paid is not None and base:
        return round(100.0 * paid / base, 1)
    return None


def load_file(path: Path, d):
    """UIN → снимок объекта.

    В Simple List на один УИН несколько строк = годы финансирования (+ иногда
    несколько контрактов). СГ и «Процент выплат» — поля объекта (одинаковы на
    всех строках). Платежи/НМЦК — на контрактной строке. Предельная стоимость —
    лимит объекта (база для % оплаты).
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    hdr = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    col = {h: i for i, h in enumerate(hdr) if h}

    def c(*names):
        for n in names:
            if n in col:
                return col[n]
        return None

    i_uin = c("Код УИН")
    i_grbs = c("ГРБС")
    i_name = c("Название объекта")
    i_work = c("Вид работ")
    i_sg = c("Стройготовность")
    i_pct = c("Процент выплат")
    i_paid = c("Сумма выполненных платежей по контракту")
    i_adv = c("Сумма выполненных предоплат по контракту")
    i_nmck = c("Начальная максимальная цена контракта")
    i_limit = c("Предельная стоимость по объекту, тыc.руб.", "Предельная стоимость по объекту, тыс.руб.")
    i_rn = c("Р/Н контракта")
    i_ctr = c("Наименование подрядчика")
    i_status = c("Статус контракта")
    i_exp_plan = c("Плановая стоимость объекта экспертизы")
    i_exp_agreed = c("Согласованная стоимость объекта экспертизы")
    i_exp_in = c("Дата подачи заявления (захода) на экспертизу")
    i_exp_start = c("Дата начала экспертизы")
    i_exp_done = c("Дата получения заключения (завершения ) экспертизы")
    i_exp_code = c("Код заявления на прохождение экспертизы")
    i_egrz = c("Номер заключения экспертизы")
    if i_uin is None or i_grbs is None:
        wb.close()
        return {}

    # сначала собираем все строки по УИН, потом сворачиваем
    rows_by: dict = defaultdict(list)
    for row in ws.iter_rows(min_row=2, values_only=True):
        uin = uin_key(row[i_uin])
        if not uin:
            continue
        grbs = str(row[i_grbs] or "").strip()
        if "минстрой" not in grbs.lower():
            continue
        name = row[i_name] if i_name is not None else None
        work = str(row[i_work] or "").strip() if i_work is not None else ""
        if is_snos(work, str(name or "")):
            continue
        rows_by[uin].append(row)
    wb.close()

    best = {}
    for uin, rows in rows_by.items():
        sg = None
        field_pct = None
        limit = None
        name = None
        work = ""
        exp_plan = None
        exp_agreed = None
        exp_in = None
        exp_start = None
        exp_done = None
        entered_exp = False
        egrz = None
        # контракты по Р/Н (или anon), берём max nmck на rn
        contracts: dict = {}
        for row in rows:
            if name is None and i_name is not None:
                name = row[i_name]
            if i_work is not None and row[i_work]:
                work = str(row[i_work]).strip()
            v_sg = fnum(row[i_sg]) if i_sg is not None else None
            if v_sg is not None:
                sg = v_sg
            v_pct = fnum(row[i_pct]) if i_pct is not None else None
            if v_pct is not None:
                field_pct = v_pct
            v_lim = fnum(row[i_limit]) if i_limit is not None else None
            if v_lim is not None and v_lim > 0:
                limit = max(limit or 0, v_lim)

            # экспертиза — поля объекта, берём с любой заполненной строки
            v = fnum(row[i_exp_plan]) if i_exp_plan is not None else None
            if v is not None and v > 0:
                exp_plan = v
            v = fnum(row[i_exp_agreed]) if i_exp_agreed is not None else None
            if v is not None and v > 0:
                exp_agreed = v
            dt = parse_date(row[i_exp_in]) if i_exp_in is not None else None
            if dt:
                exp_in = dt
            dt = parse_date(row[i_exp_start]) if i_exp_start is not None else None
            if dt:
                exp_start = dt
            dt = parse_date(row[i_exp_done]) if i_exp_done is not None else None
            if dt:
                exp_done = dt
            if i_exp_code is not None and row[i_exp_code]:
                entered_exp = True
            if i_egrz is not None and row[i_egrz] and not egrz:
                egrz = str(row[i_egrz]).strip()

            paid = fnum(row[i_paid]) if i_paid is not None else None
            adv = fnum(row[i_adv]) if i_adv is not None else None
            nmck = fnum(row[i_nmck]) if i_nmck is not None else None
            rn = row[i_rn] if i_rn is not None else None
            ctr = row[i_ctr] if i_ctr is not None else None
            status = row[i_status] if i_status is not None else None
            if not rn and not (nmck and nmck > 0) and not (paid and paid > 0):
                continue
            key = rn or f"_anon_{id(row)}"
            prev = contracts.get(key)
            score = (nmck or 0, paid or 0)
            if prev is None or score > (prev["nmck"] or 0, prev["paid"] or 0):
                contracts[key] = {
                    "paid": paid,
                    "adv": adv,
                    "nmck": nmck,
                    "contractor": ctr,
                    "status": status,
                }

        active = {
            k: v
            for k, v in contracts.items()
            if "расторг" not in str(v.get("status") or "").lower()
        }
        use = active or contracts
        if use:
            main = max(use.values(), key=lambda v: (v["nmck"] or 0, v["paid"] or 0))
            paid = main["paid"]
            adv = main["adv"]
            nmck = main["nmck"]
            ctr = main["contractor"]
            status = main["status"]
        else:
            paid = adv = nmck = ctr = status = None

        best[uin] = {
            "d": d,
            "sg": sg,
            "field_pct": field_pct,
            "paid": paid,
            "adv": adv,
            "nmck": nmck,
            "limit": limit,
            "name": name,
            "work": work,
            "contractor": ctr,
            "status": status,
            "exp_plan": exp_plan,
            "exp_agreed": exp_agreed,
            "exp_in": exp_in,
            "exp_start": exp_start,
            "exp_done": exp_done,
            "entered_exp": entered_exp,
            "egrz": egrz,
        }
    return best


def main():
    if not SL_DIR.is_dir():
        sys.exit(f"No folder {SL_DIR}")

    files = sorted(
        [f for f in SL_DIR.glob("*.xlsx") if stamp(f.name)],
        key=lambda f: stamp(f.name),
    )
    # пропускаем заведомо пустые/битые срезы
    series = defaultdict(list)
    dates = []
    n_snos_last = 0

    for f in files:
        d = stamp(f.name)
        print(f"load {d} {f.name} ...", flush=True)
        snap = load_file(f, d)
        if not snap:
            print("  skip empty")
            continue
        dates.append(d.isoformat())
        for uin, p in snap.items():
            series[uin].append(p)
        print(f"  objects={len(snap)}")

    if not series:
        sys.exit("No Минстрой data")

    # финальный портфель = объекты из последнего непустого среза (уже без сноса)
    last_date = max(max(p["d"] for p in pts) for pts in series.values())
    last_file = next(f for f in reversed(files) if stamp(f.name) == last_date)

    # предварительный список УИН последнего среза — для funds/marks
    portfolio_uins = {
        uin for uin, pts in series.items() if max(p["d"] for p in pts) == last_date
    }
    funding = load_funding(last_file, portfolio_uins)
    ksg = load_ksg_marks(portfolio_uins)

    objects = []
    for uin, pts in series.items():
        pts = sorted(pts, key=lambda x: x["d"])
        # только если был в последнем срезе
        if pts[-1]["d"] != last_date:
            continue
        last = pts[-1]
        nmck = last.get("nmck") or 0
        limit = last.get("limit") or 0
        exp_plan = last.get("exp_plan")
        exp_agreed = last.get("exp_agreed")
        exp_in = last.get("exp_in")
        exp_start = last.get("exp_start")
        exp_done = last.get("exp_done")
        entered_exp = bool(last.get("entered_exp"))
        egrz = last.get("egrz")
        # протянуть nmck/лимит/экспертизу назад, если в ранних срезах пусто
        for p in pts:
            if (p.get("nmck") or 0) > 0:
                nmck = p["nmck"]
            if (p.get("limit") or 0) > 0:
                limit = p["limit"]
            if p.get("exp_plan"):
                exp_plan = p["exp_plan"]
            if p.get("exp_agreed"):
                exp_agreed = p["exp_agreed"]
            if p.get("exp_in"):
                exp_in = p["exp_in"]
            if p.get("exp_start"):
                exp_start = p["exp_start"]
            if p.get("exp_done"):
                exp_done = p["exp_done"]
            if p.get("entered_exp"):
                entered_exp = True
            if p.get("egrz"):
                egrz = p["egrz"]
        # база % оплаты и рублёвого разрыва — предельная (лимит), иначе НМЦК
        pay_base = limit if limit > 0 else nmck
        # стоимости экспертизы в Simple List — тыс. ₽
        exp_plan_mln = round(exp_plan / 1000, 1) if exp_plan else None
        exp_agreed_mln = round(exp_agreed / 1000, 1) if exp_agreed else None
        exp_delta_mln = (
            round((exp_agreed - exp_plan) / 1000, 1)
            if exp_plan and exp_agreed
            else None
        )
        exp_overrun = (
            round((exp_agreed - exp_plan) / exp_plan * 100, 1)
            if exp_plan and exp_agreed and exp_plan > 0
            else None
        )
        traj = []
        first_sg = None
        for p in pts:
            sg = p["sg"]
            pct = pay_pct_of(
                p.get("paid"),
                p.get("limit") or limit,
                p.get("nmck") or nmck,
                p.get("field_pct"),
            )
            if sg is not None and first_sg is None:
                first_sg = sg
            traj.append(
                {
                    "d": p["d"].strftime("%d.%m.%y"),
                    "sg": round(sg, 1) if sg is not None else None,
                    "pct": pct,
                }
            )
        sg = round(last["sg"], 1) if last["sg"] is not None else None
        pct = pay_pct_of(last.get("paid"), limit, nmck, last.get("field_pct"))
        gap = round(sg - pct, 1) if sg is not None and pct is not None else None
        sg_delta = round(sg - first_sg, 1) if sg is not None and first_sg is not None else None
        objects.append(
            {
                "uin": uin,
                "name": short(last["name"]),
                "full": last["name"] or uin,
                "work": last["work"] or "—",
                "bucket": work_bucket(last["work"]),
                "contractor": short_contractor(last["contractor"]),
                "status": last["status"],
                "sg": sg,
                "pct": pct,
                "contract_value": round(nmck / 1e6, 1) if nmck else None,
                "limit": round(limit / 1e6, 1) if limit else None,
                "gap": gap,
                "gap_rub": round(gap / 100 * pay_base / 1e6, 1) if gap is not None and pay_base else None,
                "money_status": money_status(sg, pct),
                "sg_delta": sg_delta,
                "exp_plan_mln": exp_plan_mln,
                "exp_agreed_mln": exp_agreed_mln,
                "exp_delta_mln": exp_delta_mln,
                "exp_overrun": exp_overrun,
                "exp_in": fmt_date(exp_in),
                "exp_start": fmt_date(exp_start),
                "exp_done": fmt_date(exp_done),
                "exp_kind": classify_exp(exp_done, egrz, exp_agreed),
                "entered_exp": entered_exp,
                "n_pts": len(traj),
                "traj": traj,
                "marks": ksg.get(uin) or {},
                "funds": funding.get(uin) or [],
            }
        )

    objects.sort(key=lambda o: (-(o["sg"] if o["sg"] is not None else -1), o["name"] or ""))

    n_credit = sum(1 for o in objects if o["money_status"] == "credit")
    n_over = sum(1 for o in objects if o["money_status"] == "over")
    sgs = [o["sg"] for o in objects if o["sg"] is not None]
    pcts = [o["pct"] for o in objects if o["pct"] is not None]
    n_with_funds = sum(1 for o in objects if o["funds"])
    n_with_marks = sum(1 for o in objects if o["marks"])

    def median(xs):
        if not xs:
            return None
        xs = sorted(xs)
        return round(xs[len(xs) // 2], 1)

    payload = {
        "meta": {
            "grbs": "Минстрой",
            "exclude": "снос",
            "as_of": last_date.isoformat(),
            "dates": dates,
            "n": len(objects),
            "n_credit": n_credit,
            "n_over": n_over,
            "median_sg": median(sgs),
            "median_pct": median(pcts),
            "n_with_funds": n_with_funds,
            "n_with_marks": n_with_marks,
            "simple_list": last_file.name,
            "note": "СГ и % оплаты — срезы Simple List: факт СГ и «Процент выплат» (= платежи / предельная стоимость). Несколько строк на УИН = годы финансирования; контрактные суммы — со строки с Р/Н.",
        },
        "objects": objects,
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(
        f"OK {OUT} — {len(objects)} объектов, as_of={last_date}, "
        f"median_sg={payload['meta']['median_sg']}, median_pct={payload['meta']['median_pct']}, "
        f"size={OUT.stat().st_size // 1024} kb"
    )


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Generate index.html from Schools xlsx data."""
import json
import math
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent

ADVANCE_PCTS = {30.0, 49.0}  # типовые проценты аванса, а не расчётный факт оплаты

# Ручные исправления опечаток источника (Simple List расходится с более свежими данными).
NAME_OVERRIDES = {
    "1000001282.1000001075": "МБОУ СОШ № 28 ГОЩ",  # источник даёт МАОУ, актуальный тип — МБОУ
}


def short(n, keep_prefix=False):
    n = n or ""
    if not keep_prefix:
        n = n.replace("МБОУ ", "").replace("МАОУ ", "").replace("МОУ ", "")
        n = n.lstrip("-– ").strip()
    return (n[:42] + "…") if len(n) > 42 else n


def short_contractor(c):
    if not c:
        return None
    c = re.sub(
        r'(ООО|ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ|ЗАО|АО|Закрытое акционерное общество|Акционерное общество)\s*',
        "", str(c), flags=re.I,
    ).strip()
    c = c.strip('"\' ')
    return c.title() if c.isupper() else c


def parse_date(v):
    if isinstance(v, datetime):
        return v
    if not v:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(v).strip(), fmt)
        except ValueError:
            pass
    return None


def parse_amt(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace("\xa0", "").replace(" ", "").replace(",", "."))
    except ValueError:
        return 0.0


def two_tailed_p(r, n):
    if r is None or n < 3 or abs(r) >= 1:
        return None
    t = r * math.sqrt(n - 2) / math.sqrt(1 - r**2)
    return round(2 * (1 - 0.5 * (1 + math.erf(abs(t) / math.sqrt(2)))), 3)


def median(xs):
    n = len(xs)
    if n == 0:
        return 0
    s = sorted(xs)
    mid = n // 2
    return (s[mid - 1] + s[mid]) / 2 if n % 2 == 0 else s[mid]


def corr(xs, ys):
    n = len(xs)
    if n < 3:
        return None
    mx, my = sum(xs) / n, sum(ys) / n
    num = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    dx = math.sqrt(sum((x - mx) ** 2 for x in xs))
    dy = math.sqrt(sum((y - my) ** 2 for y in ys))
    if dx == 0 or dy == 0:
        return None
    return num / (dx * dy)


def linreg(xs, ys):
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    num = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    den = sum((x - mx) ** 2 for x in xs)
    b = num / den
    a = my - b * mx
    pred = [a + b * x for x in xs]
    ss_res = sum((y - p) ** 2 for y, p in zip(ys, pred))
    ss_tot = sum((y - my) ** 2 for y in ys)
    return a, b, 1 - ss_res / ss_tot if ss_tot else 0


def fmt_date(d):
    return d.strftime("%d.%m.%y") if d else None


def kt_bounds(items):
    if not items:
        return {}
    return {
        "plan_start": min((x["ps"] for x in items if x["ps"]), default=None),
        "plan_end": max((x["pe"] for x in items if x["pe"]), default=None),
        "fact_start": min((x["fs"] for x in items if x["fs"]), default=None),
        "fact_end": max((x["fe"] for x in items if x["fe"]), default=None),
    }


def load_ksg():
    path = ROOT / "1708_КСГ+Экспертиза.xlsx"
    if not path.exists():
        return {}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    headers = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    col = {h: i for i, h in enumerate(headers)}
    kt = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        uin = row[col["УИН"]]
        if not uin:
            continue
        proc = row[col["Процедура КСГ"]]
        kt.setdefault(uin, {}).setdefault(proc, []).append(
            {
                "ps": parse_date(row[col["Дата начала план"]]),
                "pe": parse_date(row[col["Дата окончания план"]]),
                "fs": parse_date(row[col["Дата начала факт"]]),
                "fe": parse_date(row[col["Дата окончания факт"]]),
            }
        )
    wb.close()
    return kt


SIMPLE_LIST_FILE = "0709_Акцент_Simple List.xlsx"


def load_simple_list():
    wb = openpyxl.load_workbook(ROOT / SIMPLE_LIST_FILE, data_only=True)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    col = {h: i for i, h in enumerate(headers)}
    sl = {}
    years = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        uin = row[col["Код УИН"]]
        if not uin:
            continue
        yr = row[col["Год финансирования"]]

        def num(name):
            v = row[col[name]]
            if v is None:
                return None
            if isinstance(v, (int, float)):
                return float(v)
            try:
                return float(str(v).replace("\xa0", "").replace(" ", "").replace(",", "."))
            except ValueError:
                return None

        if yr and str(yr).isdigit() and int(yr) >= 2020:
            years.setdefault(uin, []).append(
                {
                    "year": int(yr),
                    "plan": num("Сумма планового финансирования") or 0,
                    "obligated": num("Сумма бюджетных обязательств") or 0,
                    "fact": num("Сумма фактического финансирования") or 0,
                }
            )

        if uin in sl and yr and str(yr).isdigit() and int(yr) <= int(sl[uin].get("yr") or 0):
            continue
        pay_pct = None
        v = row[col["Процент выплат"]]
        if v is not None:
            try:
                pay_pct = round(float(str(v).replace(",", ".")), 1)
            except ValueError:
                pass
        prev = sl.get(uin) or {}
        contractor = row[col["Наименование подрядчика"]] or prev.get("contractor")
        inn = row[col["ИНН подрядчика"]] or prev.get("inn")
        rp = row[col["РП"]] or prev.get("rp")
        entered_exp = bool(row[col["Код заявления на прохождение экспертизы"]]) or prev.get("entered_exp", False)

        exp_overrun = prev.get("exp_overrun")
        plan_c, agreed_c = num("Плановая стоимость объекта экспертизы"), num("Согласованная стоимость объекта экспертизы")
        if plan_c and agreed_c and plan_c > 0:
            exp_overrun = round((agreed_c - plan_c) / plan_c * 100, 1)

        contract_value = (
            num("Начальная максимальная цена контракта")
            or num("Предельная стоимость по объекту, тыc.руб.")
            or prev.get("contract_value")
        )

        sl[uin] = {
            "yr": yr,
            "name": row[col["Название объекта"]],
            "pay_pct": pay_pct,
            "contractor": contractor,
            "inn": inn,
            "rp": rp,
            "contract_value": contract_value,
            "entered_exp": entered_exp,
            "exp_overrun": exp_overrun,
            "exp_plan_entry": parse_date(row[col["Плановая дата захода на экспертизу из ДК"]]) or prev.get("exp_plan_entry"),
            "opening_plan": parse_date(row[col["Планируемая дата открытия"]]) or prev.get("opening_plan"),
            "exp_in": parse_date(row[col["Дата подачи заявления (захода) на экспертизу"]]),
            "exp_start": parse_date(row[col["Дата начала экспертизы"]]),
            "exp_done": parse_date(row[col["Дата получения заключения (завершения ) экспертизы"]]),
            "ctr_plan": parse_date(row[col["Заключение контракта начало план КСГ"]]),
            "ctr_fact": parse_date(row[col["Заключение контракта начало факт КСГ"]]),
        }

    for uin, rows in years.items():
        by_year = {}
        for r in rows:
            e = by_year.setdefault(r["year"], {"year": r["year"], "plan": 0, "obligated": 0, "fact": 0})
            e["plan"] += r["plan"]
            e["obligated"] += r["obligated"]
            e["fact"] += r["fact"]
        if uin in sl:
            sl[uin]["program_years"] = sorted(by_year.values(), key=lambda x: x["year"])
    return sl


def load_finance2026():
    path = ROOT / "finance2026.json"
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def load_addresses():
    path = ROOT / "addresses.json"
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def build_flags(last, info, exp, smr, rs, ctr):
    flags = []
    sg_plan_gap = round(last["fact"] - last["plan"], 1) if last["plan"] is not None else None
    pay_pct = info.get("pay_pct")
    sg_pay_gap = round(last["fact"] - pay_pct, 1) if pay_pct is not None else None

    if sg_pay_gap is not None and sg_pay_gap > 55:
        flags.append("готовность выше выплат")
    if sg_plan_gap is not None and sg_plan_gap < -5:
        flags.append("отстаёт от плана")
    if smr.get("fact_start") and exp.get("plan_end") and smr["fact_start"] < exp["plan_end"]:
        flags.append("стройка до экспертизы")
    if info.get("exp_done") and exp.get("plan_end") and info["exp_done"] < exp["plan_end"]:
        flags.append("экспертиза не бьётся с КСГ")
    if smr.get("fact_start") and ctr.get("fact_start") and smr["fact_start"] < ctr["fact_start"]:
        flags.append("стройка до контракта")
    if rs.get("plan_end") and rs.get("fact_end") and rs["fact_end"] > rs["plan_end"]:
        flags.append("РС опоздал")
    elif not rs and smr.get("fact_start") and not info.get("exp_done"):
        flags.append("нет заключения экспертизы")

    return flags, sg_plan_gap, sg_pay_gap


def load_data():
    sl = load_simple_list()
    ksg = load_ksg()
    fin2026 = load_finance2026()
    addresses = load_addresses()

    cross, per, traj, kt_rows, kt_dates = [], [], {}, [], {}
    budget_alert = []
    STAGE_EDGES = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
    stage_gaps = {e: {"plan": [], "pay": []} for e in STAGE_EDGES}

    for f in sorted(ROOT.glob("*.xlsx")):
        if (
            "_платежи" in f.name
            or "Акцент" in f.name
            or "Simple" in f.name
            or "КСГ" in f.name
        ):
            continue
        uin = f.stem
        info = sl.get(uin, {})
        name = NAME_OVERRIDES.get(uin, info.get("name", uin))
        wb = openpyxl.load_workbook(f, data_only=True)
        series = []
        for row in wb.active.iter_rows(min_row=2, values_only=True):
            d = parse_date(row[0])
            if not d:
                continue
            try:
                plan = float(row[1]) if row[1] is not None else None
                fact = float(row[2])
            except (TypeError, ValueError):
                continue
            series.append({"d": d, "plan": plan, "fact": fact})
        series.sort(key=lambda x: x["d"])
        if not series:
            continue

        pays = []
        pf = ROOT / f"{uin}_платежи.xlsx"
        if pf.exists():
            wb2 = openpyxl.load_workbook(pf, data_only=True)
            hdr = [c.value for c in next(wb2.active.iter_rows(min_row=1, max_row=1))]
            if hdr and hdr[0] == "Тип платежа":
                for row in wb2.active.iter_rows(min_row=2, values_only=True):
                    d = parse_date(row[4])
                    amt = parse_amt(row[5])
                    if d and amt:
                        pays.append({"d": d, "amt": amt, "advance": row[0] == "Предоплата"})
        pays.sort(key=lambda x: x["d"])
        total = sum(p["amt"] for p in pays)
        last = series[-1]
        pay_pct = info.get("pay_pct")

        pay_scale = pay_pct / total if (pay_pct is not None and total > 0) else None

        cum = cum_adv = cum_reg = 0
        pi = 0
        pts = []
        for pt in series:
            while pi < len(pays) and pays[pi]["d"] <= pt["d"]:
                amt = pays[pi]["amt"]
                cum += amt
                if pays[pi]["advance"]:
                    cum_adv += amt
                else:
                    cum_reg += amt
                pi += 1
            pts.append(
                {
                    "d": pt["d"].strftime("%d.%m.%y"),
                    "sg": round(pt["fact"], 1),
                    "plan": round(pt["plan"], 1) if pt["plan"] is not None else None,
                    "pay": round(cum / 1e6, 1),
                    "payPct": round(cum * pay_scale, 1) if pay_scale is not None else None,
                    "advPct": round(cum_adv * pay_scale, 1) if pay_scale is not None else None,
                    "regPct": round(cum_reg * pay_scale, 1) if pay_scale is not None else None,
                }
            )
        for edge in STAGE_EDGES:
            hit = next((p for p in pts if p["sg"] >= edge), None)
            if hit is None:
                continue
            if hit["plan"] is not None:
                stage_gaps[edge]["plan"].append(hit["sg"] - hit["plan"])
            if hit["payPct"] is not None:
                stage_gaps[edge]["pay"].append(hit["sg"] - hit["payPct"])

        if len(pts) > 36:
            step = math.ceil(len(pts) / 36)
            pts = pts[::step][:-1] + [pts[-1]]
        traj[uin] = pts

        pays_t = [p["pay"] for p in pts]
        vary = max(pays_t) - min(pays_t) >= 0.01
        r = None
        if vary and len(pts) >= 5:
            r = corr([p["sg"] for p in pts], pays_t)

        uin_kt = ksg.get(uin, {})
        exp = kt_bounds(uin_kt.get("Экспертиза", []))
        smr = kt_bounds(uin_kt.get("СМР", []))
        ctr = kt_bounds(uin_kt.get("Заключение контракта 1", []))
        rs = kt_bounds(uin_kt.get("Получение РС", []))
        flags, sg_plan_gap, sg_pay_gap = build_flags(last, info, exp, smr, rs, ctr)

        fin = fin2026.get(uin)
        if fin and fin.get("osv2026") == 0:
            flags.append("не осваивает бюджет 2026")

        exp_ref = info.get("exp_done") or exp.get("plan_end")
        exp_marker = None
        if exp_ref and pts:
            exp_marker = min(
                pts, key=lambda p: abs((datetime.strptime(p["d"], "%d.%m.%y") - exp_ref).days)
            )["d"]

        kt_dates[uin] = {
            "exp_plan": fmt_date(exp.get("plan_end")),
            "exp_fact": fmt_date(exp.get("fact_end")),
            "exp_sl": fmt_date(info.get("exp_done")),
            "exp_marker": exp_marker,
            "smr_start": fmt_date(smr.get("fact_start")),
            "ctr_fact": fmt_date(ctr.get("fact_start") or info.get("ctr_fact")),
            "rs_plan": fmt_date(rs.get("plan_end")),
            "rs_fact": fmt_date(rs.get("fact_end")),
        }

        contractor = short_contractor(info.get("contractor"))
        addr = addresses.get(uin, {})
        opening_plan = info.get("opening_plan")
        days_to_open = (opening_plan - datetime.now()).days if opening_plan else None

        program_years = info.get("program_years", [])
        program_unbacked = round(
            sum(y["plan"] for y in program_years if y["year"] <= 2026 and y["obligated"] == 0 and y["fact"] == 0) / 1e6,
            1,
        )

        contract_value = info.get("contract_value")
        gap_pct = round(pay_pct - last["fact"], 1) if pay_pct is not None else None
        gap_rub = round(gap_pct / 100 * contract_value / 1e6, 1) if gap_pct is not None and contract_value else None
        if gap_pct is None:
            money_status = "unknown"
        elif gap_pct > 10:
            money_status = "over"
        elif gap_pct < -10:
            money_status = "credit"
        else:
            money_status = "balanced"

        kt_rows.append(
            {
                "uin": uin,
                "name": short(name),
                "full": name,
                "sg": round(last["fact"], 1),
                "plan": round(last["plan"], 1) if last["plan"] is not None else None,
                "sg_plan_gap": sg_plan_gap,
                "pct": pay_pct,
                "sg_pay_gap": sg_pay_gap,
                "flags": flags,
                "flag_n": len(flags),
                "contractor": contractor,
                "inn": info.get("inn"),
                "rp": info.get("rp"),
                "address": addr.get("address"),
                "municipality": addr.get("municipality"),
                "exp_overrun": info.get("exp_overrun"),
                "entered_exp": info.get("entered_exp", False),
                "opening_plan": fmt_date(opening_plan),
                "days_to_open": days_to_open,
                "program_years": program_years,
                "program_unbacked": program_unbacked,
                "contract_value": round(contract_value / 1e6, 1) if contract_value else None,
                "gap_pct": gap_pct,
                "gap_rub": gap_rub,
                "money_status": money_status,
            }
        )

        cross.append(
            {
                "uin": uin,
                "name": short(name),
                "full": name,
                "sg": round(last["fact"], 1),
                "pct": pay_pct,
                "advance": pay_pct in ADVANCE_PCTS,
                "pay": round(total / 1e6, 1),
                "gap": sg_pay_gap,
                "contractor": contractor,
            }
        )
        per.append(
            {
                "uin": uin,
                "name": short(name),
                "sg": round(last["fact"], 1),
                "plan": round(last["plan"], 1) if last["plan"] is not None else None,
                "sg_plan_gap": sg_plan_gap,
                "pct": pay_pct,
                "pay": round(total / 1e6, 1),
                "r": round(r, 3) if r is not None else None,
                "vary": vary,
                "n": len(series),
                "flags": flags,
            }
        )

    name_counts = Counter(c["name"] for c in cross)
    dup_uins = {c["uin"] for c in cross if name_counts[c["name"]] > 1}
    if dup_uins:
        uin_full = {c["uin"]: c["full"] for c in cross}
        for row in (*kt_rows, *cross, *per):
            if row["uin"] in dup_uins:
                row["name"] = short(uin_full[row["uin"]], keep_prefix=True)

    for k in kt_rows:
        fin = fin2026.get(k["uin"])
        if fin and fin.get("osv2026") == 0:
            budget_alert.append(
                {
                    "uin": k["uin"],
                    "name": k["name"],
                    "full": k["full"],
                    "sg": k["sg"],
                    "plan2026": round(fin["plan2026_rub"] / 1e6, 1),
                }
            )
    budget_alert.sort(key=lambda x: -x["plan2026"])

    by_contractor = {}
    for k in kt_rows:
        c = k["contractor"] or "Не указан"
        e = by_contractor.setdefault(c, {"contractor": c, "objects": [], "n_no_budget2026": 0, "n_flags": 0})
        no_bud = "не осваивает бюджет 2026" in k["flags"]
        e["objects"].append({"name": k["name"], "full": k["full"], "sg": k["sg"], "no_budget2026": no_bud})
        if no_bud:
            e["n_no_budget2026"] += 1
        if k["flags"]:
            e["n_flags"] += 1
    contractors = sorted(by_contractor.values(), key=lambda x: (-len(x["objects"]), x["contractor"]))

    stage_analysis = []
    for edge in STAGE_EDGES:
        plan_vals = stage_gaps[edge]["plan"]
        pay_vals = stage_gaps[edge]["pay"]
        if not plan_vals and not pay_vals:
            continue
        stage_analysis.append(
            {
                "stage": edge,
                "n": max(len(plan_vals), len(pay_vals)),
                "plan_gap": round(median(plan_vals), 1) if plan_vals else None,
                "pay_gap": round(median(pay_vals), 1) if pay_vals else None,
            }
        )

    stage_by_edge = {s["stage"]: s["pay_gap"] for s in stage_analysis if s["pay_gap"] is not None}

    red_zone = []
    for k in kt_rows:
        if k["sg_pay_gap"] is None:
            continue
        stage = max((e for e in STAGE_EDGES if e <= k["sg"] and e in stage_by_edge), default=None)
        if stage is None:
            continue
        deviation = round(k["sg_pay_gap"] - stage_by_edge[stage], 1)
        red_zone.append(
            {
                "uin": k["uin"],
                "name": k["name"],
                "full": k["full"],
                "sg": k["sg"],
                "gap": k["sg_pay_gap"],
                "stage": stage,
                "stage_median": stage_by_edge[stage],
                "deviation": deviation,
            }
        )
    red_zone.sort(key=lambda x: -x["deviation"])

    advance_only = []
    for k in kt_rows:
        t = traj.get(k["uin"])
        if not t:
            continue
        last_pt = t[-1]
        adv, reg = last_pt.get("advPct"), last_pt.get("regPct")
        if adv and adv > 0 and not reg:
            advance_only.append({"uin": k["uin"], "name": k["name"], "full": k["full"], "sg": k["sg"], "advPct": adv})
    advance_only.sort(key=lambda x: -x["advPct"])

    valid = [s for s in cross if s["pct"] is not None and s["pay"] > 0]
    facts = [s["sg"] for s in valid]
    pcts = [s["pct"] for s in valid]
    a, b, r2 = linreg(pcts, facts)
    varying = [p for p in per if p["vary"] and p["r"] is not None]
    varying.sort(key=lambda x: -(x["r"] or -1))
    rs_corr = [p["r"] for p in varying]

    gaps = [s["gap"] for s in cross if s["gap"] is not None]
    kt_sorted = sorted(kt_rows, key=lambda x: (-x["flag_n"], -(x["sg_pay_gap"] or 0)))

    no_adv = [s for s in valid if not s["advance"]]
    facts_na = [s["sg"] for s in no_adv]
    pcts_na = [s["pct"] for s in no_adv]
    pearson_na = corr(facts_na, pcts_na) if len(no_adv) >= 3 else None

    return {
        "stats": {
            "n": len(valid),
            "median_r": round(median(rs_corr), 3) if rs_corr else 0,
            "n_varying": len(varying),
            "n_sg_ahead": sum(1 for g in gaps if g > 55),
            "median_gap": round(median(gaps), 1) if gaps else 0,
            "n_sg_behind_plan": sum(
                1 for p in per if p.get("sg_plan_gap") is not None and p["sg_plan_gap"] < -5
            ),
            "n_smr_before_exp": sum(1 for p in per if "стройка до экспертизы" in p.get("flags", [])),
            "n_kt_issues": sum(1 for p in per if any(f in p.get("flags", []) for f in ("стройка до экспертизы", "экспертиза не бьётся с КСГ", "РС опоздал", "нет заключения экспертизы"))),
            "n_no_advance": len(no_adv),
            "pearson_no_advance": round(pearson_na, 3) if pearson_na is not None else None,
            "pearson_no_advance_p": two_tailed_p(pearson_na, len(no_adv)),
            "n_no_budget2026": len(budget_alert),
            "n_red_zone": sum(1 for r in red_zone if r["deviation"] > 15),
        },
        "budget_alert": budget_alert,
        "contractors": contractors,
        "stage_analysis": stage_analysis,
        "red_zone": red_zone,
        "advance_only": advance_only,
        "kt_attention": kt_sorted[:10],
        "objects": kt_rows,
        "reg": {"a": round(a, 1), "b": round(b, 4)},
        "cross": cross,
        "per": sorted(per, key=lambda x: -(x["r"] if x["r"] is not None else -1)),
        "traj": traj,
        "kt_dates": kt_dates,
        "defaultUin": varying[0]["uin"] if varying else per[0]["uin"],
    }


HEAD_STYLE = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>__TITLE__</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Golos+Text:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-annotation@3.0.1/dist/chartjs-plugin-annotation.min.js"></script>
<style>
  :root {
    --bg:#EEF3F8; --surface:#fff; --div:#EDF2F7; --line:#DCE6F0;
    --text:#0D2040; --muted:#5A7189; --faint:#93A8BC;
    --accent:#1B8A9C; --accent-l:#22B0C8; --accent-d:#126880; --accent-dim:rgba(27,138,156,.09);
    --ok:#27AE60; --warn:#E8A020; --err:#D94040; --info:#2E7CC4;
    --ok-bg:#E8F6EE; --warn-bg:#FDF3E2; --err-bg:#FBE9E9; --info-bg:#E9F1FB;
  }
  * { box-sizing:border-box }
  body { margin:0; font:15px/1.55 'Golos Text',Inter,system-ui,sans-serif; background:var(--bg); color:var(--text); letter-spacing:-.01em }
  .wrap { max-width:__WRAP__px; margin:0 auto; padding:28px 18px 56px }
  h1 { font-size:1.65rem; font-weight:700; margin:0 0 4px; letter-spacing:-.02em;
    background:linear-gradient(135deg, hsl(200,60%,23%), hsl(188,70%,30%));
    -webkit-background-clip:text; background-clip:text; -webkit-text-fill-color:transparent; display:inline-block }
  .sub { color:var(--muted); margin:0 0 20px; font-size:.92rem }
  ul.brief { margin:0; padding-left:1.2rem; color:var(--muted) }
  ul.brief li { margin:6px 0 }
  .kpis { display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin:16px 0 }
  .kpi { background:var(--surface); border:1px solid var(--line); border-radius:10px; padding:12px 14px; box-shadow:0 1px 2px rgba(13,32,64,.04) }
  .kpi .n { font-size:1.35rem; font-weight:700; color:var(--accent-d); line-height:1.2; font-variant-numeric:tabular-nums }
  .kpi .l { font-size:.75rem; color:var(--muted); margin-top:4px }
  .chart { position:relative; height:280px; margin-top:10px }
  .chart.tall { height:340px }
  .note { font-size:.82rem; color:var(--faint); margin:6px 0 0 }
  .box { background:var(--surface); border:1px solid var(--line); border-radius:12px; padding:16px 18px; margin:12px 0; box-shadow:0 1px 3px rgba(13,32,64,.05) }
  .box.stub { background:repeating-linear-gradient(135deg, var(--surface) 0 10px, #F7FAFC 10px 20px); border-style:dashed }
  .mini table { width:100%; font-size:.84rem; border-collapse:collapse }
  .mini th,.mini td { padding:7px 8px; border-bottom:1px solid var(--div); text-align:left }
  .mini th { color:var(--faint); font-weight:600; font-size:.72rem; letter-spacing:.04em; text-transform:uppercase; background:#F7FAFC }
  .mini td.r,.mini th.r { text-align:right; font-variant-numeric:tabular-nums }
  details { background:var(--surface); border:1px solid var(--line); border-radius:12px; margin:10px 0; overflow:hidden; box-shadow:0 1px 3px rgba(13,32,64,.05) }
  details > summary { cursor:pointer; padding:14px 18px; font-weight:600; list-style:none; user-select:none }
  details > summary::-webkit-details-marker { display:none }
  details > summary::after { content:'+'; float:right; color:var(--faint); font-weight:400 }
  details[open] > summary::after { content:'−' }
  details > summary:hover { background:#F7FAFC }
  .detail-body { padding:0 18px 18px; border-top:1px solid var(--div) }
  select { font:inherit; padding:6px 10px; border:1px solid var(--line); border-radius:8px; background:#fff; max-width:100% }
  .row { display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin:10px 0 }
  .tag { font-size:.75rem; padding:3px 9px; background:var(--accent-dim); border:1px solid var(--line); border-radius:99px; color:var(--accent-d); font-weight:500 }
  table.full { width:100%; border-collapse:collapse; font-size:.84rem }
  table.full th,table.full td { padding:8px 9px; border-bottom:1px solid var(--div) }
  table.full th { background:#F7FAFC; position:sticky; top:0; text-align:left; color:var(--faint); font-weight:600; font-size:.72rem; letter-spacing:.04em; text-transform:uppercase }
  table.full td.r,table.full th.r { text-align:right; font-variant-numeric:tabular-nums }
  .tbl-wrap { max-height:420px; overflow:auto; border:1px solid var(--line); border-radius:10px; margin-top:10px }
  .flag { font-size:.72rem; padding:2px 7px; margin:1px 2px 1px 0; display:inline-block; background:var(--warn-bg); border:1px solid #EFCB84; border-radius:5px; color:#8A5E10 }
  .flag.warn { background:var(--err-bg); border-color:#F0B3B3; color:#A32E2E }
  .pill { display:inline-block; font-size:.72rem; font-weight:600; padding:2px 9px; border-radius:99px }
  .pill.ok { background:var(--ok-bg); color:#1E8449 }
  .pill.warn { background:var(--warn-bg); color:#8A5E10 }
  .pill.err { background:var(--err-bg); color:#A32E2E }
  .stub-label { display:inline-block; font-size:.7rem; font-weight:600; letter-spacing:.03em; text-transform:uppercase; color:var(--faint); background:#F7FAFC; border:1px solid var(--line); border-radius:5px; padding:2px 8px }
  .kpis.four { grid-template-columns:repeat(4,1fr) }
  @media(max-width:900px) { .kpis.four { grid-template-columns:repeat(2,1fr) } }
  @media(max-width:700px) { .kpis,.kpis.four { grid-template-columns:1fr } }
  .navlink { display:inline-block; margin:0 0 16px; font-size:.85rem; color:var(--accent-d); text-decoration:none; font-weight:600 }
  .navlink:hover { text-decoration:underline }
  .filterbar { display:flex; gap:8px; flex-wrap:wrap; margin:12px 0 18px }
  .fbtn { font:inherit; font-size:.82rem; font-weight:600; padding:7px 14px; border-radius:99px; border:1px solid var(--line); background:var(--surface); color:var(--muted); cursor:pointer }
  .fbtn:hover { border-color:var(--accent) }
  .fbtn.active { background:var(--accent); border-color:var(--accent); color:#fff }
  .money { font-variant-numeric:tabular-nums; font-weight:600 }
  .money.neg { color:#A32E2E }
  .money.pos { color:#1E8449 }
  tr.clickable { cursor:pointer }
  tr.clickable:hover { background:var(--accent-dim) }
  tr.row-active { background:var(--accent-dim) }
  .card-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:10px; margin:14px 0 }
</style>
</head>
<body>
<div class="wrap">
  <h1>СГ и выплаты</h1>
  <p class="sub">__SUB__</p>
  __NAV__
"""

METHOD_BODY = r"""
  <div class="box" id="methodBox">
    <p class="note" style="margin:0 0 10px">Портфель — 48 капремонтов школ. Ниже — риск по деньгам, а не по проценту готовности: пять правил и решения, которые из них следуют. У каждого правила — цифры конкретно по этому портфелю, без общих слов.</p>

    <strong style="display:block;margin-top:16px;font-size:1.05rem">1. Судить объект по отклонению от нормы для его этапа, не по проценту оплаты</strong>
    <p class="note">Разрыв между готовностью и оплатой растёт по ходу стройки у всех объектов одинаково закономерно: на старте аванс идёт впереди работ, после экспертизы оплата по актам отстаёт. Это норма, а не проблема сама по себе. Проблема — когда конкретный объект отклоняется от этой нормы. Кривая ниже — медиана по всем 48 школам на каждом этапе готовности, это и есть точка отсчёта:</p>
    <div class="chart tall"><canvas id="cStage"></canvas></div>
    <p class="note" id="stageNote" style="margin-top:8px"></p>

    <strong style="display:block;margin-top:20px;font-size:1.05rem">2. Красная зона: <span id="mthRedN"></span> школ отклоняются от нормы на 15+ п.п.</strong>
    <p class="note">На этапе ~90% готовности норма — разрыв около <span id="mth90"></span> п.п. Школы в таблице отстают от неё сильнее любого типового лага. Решение: по каждой — запросить у технического заказчика статус актов на этой неделе, не дожидаясь планового отчёта.</p>
    <div class="tbl-wrap" style="max-height:280px">
      <table class="full">
        <thead><tr><th>Школа</th><th class="r">СГ</th><th class="r">Разрыв</th><th class="r">Норма для этапа</th><th class="r">Отклонение</th></tr></thead>
        <tbody id="redZoneTbl"></tbody>
      </table>
    </div>

    <strong style="display:block;margin-top:20px;font-size:1.05rem">3. Нулевое освоение решается на этой неделе, не в декабре</strong>
    <p class="note"><span id="mth0"></span> школ — 0% кассового исполнения бюджета 2026 года при уже утверждённом финансировании (89–418 млн ₽ на объект). Ждать конца года бессмысленно: либо деньги начинают двигаться в ближайший месяц, либо бюджет надо честно переносить на 2027-й — и это решение нужно принять сейчас.</p>
    <div class="box" style="background:var(--err-bg);border-color:#F0B3B3;margin:8px 0 0">
      <strong style="display:block;margin-bottom:6px;color:#A32E2E">⚠ <span id="k6"></span> школ: бюджет утверждён, выплат в 2026 году не было</strong>
      <p class="note" style="margin:0 0 8px">Строительство почти завершено (готовность 80–100%), кассовых выплат за год — 0%.</p>
      <table>
        <thead><tr><th>Школа</th><th class="r">СГ</th><th class="r">План 2026, млн ₽</th></tr></thead>
        <tbody id="budgetAlert"></tbody>
      </table>
    </div>

    <strong style="display:block;margin-top:20px;font-size:1.05rem">4. <span id="advOnlyN"></span> школ из 48 сидят на чистом авансе — новые деньги здесь не помогут</strong>
    <p class="note">Ни одного платежа по актам, только аванс — при этом часть объектов готова на 95–100%. Добавлять финансирование таким объектам бессмысленно: деньги не идут не из-за нехватки бюджета, а потому что акты о выполненных работах не доходят до оплаты. Разбираться нужно в процедуре приёмки — у заказчика или у подрядчика, а не в объёме транша.</p>
    <div class="tbl-wrap" style="max-height:220px">
      <table class="full">
        <thead><tr><th>Школа</th><th class="r">СГ</th><th class="r">Получено (аванс), %</th></tr></thead>
        <tbody id="advOnlyTbl"></tbody>
      </table>
    </div>

    <strong style="display:block;margin-top:20px;font-size:1.05rem">5. Когда у подрядчика стоят все объекты — вопрос к подрядчику, не к объектам</strong>
    <p class="note" id="mthContr">—</p>
    <div class="tbl-wrap">
      <table class="full">
        <thead><tr>
          <th>Подрядчик</th><th class="r">Объектов</th><th class="r">Не осваивают бюджет 2026</th><th>Объекты</th>
        </tr></thead>
        <tbody id="contractorsTbl"></tbody>
      </table>
    </div>

    <strong style="display:block;margin-top:20px;font-size:1.05rem">6. Лимит по госпрограмме — не то же самое, что деньги под контрактом</strong>
    <p class="note"><span id="unbackedN"></span> школ показывают лимит финансирования на 2025–2026 год, под который до сих пор не оформлены бюджетные обязательства (и, соответственно, нет исполнения) — это отдельный вид риска, почти не пересекающийся со списком нулевого освоения выше: деньги формально запланированы в госпрограмме, но не привязаны ни к контракту, ни к платежу.</p>
    <div class="tbl-wrap" style="max-height:220px">
      <table class="full">
        <thead><tr><th>Школа</th><th class="r">СГ</th><th class="r">Лимит без обязательств, млн ₽</th></tr></thead>
        <tbody id="unbackedTbl"></tbody>
      </table>
    </div>

    <strong style="display:block;margin-top:20px;font-size:1.05rem">Данных не хватает — заглушки вместо разделов</strong>
    <div class="box stub">
      <span class="stub-label">Данных нет</span>
      <p class="note" style="margin-top:8px">Ответственный по каждой контрольной точке (кто именно ведёт экспертизу, кто СМР, кто ввод) — в источниках есть только один общий РП на объект, разбивки по этапам нет.</p>
    </div>
    <div class="box stub">
      <span class="stub-label">Данных нет</span>
      <p class="note" style="margin-top:8px">Условия контрактов — ставки пеней, размер банковских гарантий, пороги расторжения. Без этого раздел про рычаги на подрядчика (принцип 5) остаётся на уровне общих принципов, а не конкретных цифр давления по каждому договору.</p>
    </div>

    <strong style="display:block;margin-top:24px;font-size:1.1rem">Выводы и что делать</strong>
    <ul class="brief">
      <li><strong>На этой неделе:</strong> эскалировать все <span id="conclN0"></span> школ с нулевым освоением — по каждой получить причину (акты не поданы / спор / организационная задержка у заказчика) и решение: ускорить или честно перенести бюджет на 2027 год.</li>
      <li><strong>На этой неделе:</strong> запросить статус актов по <span id="conclNRed"></span> школам в красной зоне — они отстают от нормы для своего этапа больше, чем объясняет обычный цикл оплаты.</li>
      <li><strong>В течение месяца:</strong> разбор на уровне договора с подрядчиками, у которых зависли все объекты без исключения (раздел 5 выше) — точечные меры по одной школе тут не сработают.</li>
      <li><strong>Не выделять новый аванс</strong> школам, которые уже сидят только на авансе (раздел 4) — вместо этого требовать акты. Дополнительные деньги не решают проблему с документами.</li>
      <li>Использовать эталонную кривую (раздел 1) как чек-лист перед каждым траншем: этап готовности объекта сверять с нормой, а не только со сроком по контракту.</li>
    </ul>

    <strong style="display:block;margin-top:20px">Чего не хватает для более точного анализа</strong>
    <ul class="brief">
      <li>Причина конкретной задержки (акты не поданы / спор / организационная пробуксовка) в этих данных не видна — по каждой школе из красной зоны и нулевого списка её ещё предстоит выяснить у заказчика и подрядчика.</li>
      <li>Нет истории по годам: неизвестно, типична ли пробуксовка в начале года — возможно, часть объектов обычно нагоняет в четвёртом квартале, и тогда часть «красной зоны» — не риск, а сезонность.</li>
      <li>Нет данных об условиях контрактов (штрафы, порядок расторжения) — непонятно, какие реальные рычаги есть на переговорах с проблемными подрядчиками.</li>
      <li>Список подрядчиков и бюджет 2026 года сверены вручную один раз по состоянию на начало сентября — при обновлении отчёта их нужно сверять заново, автоматически это не пересчитывается.</li>
      <li>Порог красной зоны (отклонение 15+ п.п.) — эвристика по видимому разрыву в данных, а не статистически откалиброванный норматив: выборка в 48 объектов для этого небольшая.</li>
    </ul>
  </div>
"""

DASHBOARD_BODY = r"""
  <p class="note" style="margin:0 0 14px">По каждому объекту: сколько денег получил подрядчик против того, сколько физически построил — в рублях по цене контракта, не в очках готовности. Если оплата обгоняет стройку — избыток (куда делись деньги, непонятно). Если стройка обгоняет оплату — подрядчик кредитует стройку сам, ему должны заплатить.</p>

  <div class="kpis four">
    <div class="kpi"><div class="n" id="dK1"></div><div class="l">объектов в портфеле</div></div>
    <div class="kpi"><div class="n" id="dK2"></div><div class="l">суммарный контракт, млрд ₽</div></div>
    <div class="kpi"><div class="n" id="dK3"></div><div class="l">подрядчики кредитуют, млрд ₽</div></div>
    <div class="kpi"><div class="n" id="dK4"></div><div class="l">0% освоения бюджета 2026</div></div>
  </div>

  <div class="filterbar" id="filterBar">
    <button class="fbtn active" data-f="all">Все объекты</button>
    <button class="fbtn" data-f="credit">Подрядчик кредитует &gt;100 млн ₽</button>
    <button class="fbtn" data-f="balanced">Баланс</button>
    <button class="fbtn" data-f="nobudget">0% освоения 2026</button>
  </div>

  <div class="tbl-wrap" style="max-height:520px">
    <table class="full">
      <thead><tr>
        <th>Школа</th><th>Округ</th><th>Подрядчик</th><th class="r">Контракт, млн ₽</th>
        <th class="r">СГ</th><th class="r">Оплата</th><th class="r">Разница, млн ₽</th><th>Статус</th><th>Ввод</th>
      </tr></thead>
      <tbody id="objTbl"></tbody>
    </table>
  </div>
  <p class="note">Клик по строке — график этого объекта ниже. Разница = оплата% минус СГ% × сумма контракта.</p>

  <strong style="display:block;margin-top:22px;margin-bottom:8px">Подрядчики — сводно по портфелю</strong>
  <p class="note" style="margin-top:0">Клик по подрядчику — отфильтровать таблицу выше только по его объектам.</p>
  <div class="tbl-wrap">
    <table class="full">
      <thead><tr>
        <th>Подрядчик</th><th class="r">Объектов</th><th class="r">Контракт, млн ₽</th>
        <th class="r">Подрядчик кредитует, млн ₽</th><th class="r">Не осваивают бюджет 2026</th>
      </tr></thead>
      <tbody id="contrRollup"></tbody>
    </table>
  </div>

  <strong style="display:block;margin-top:22px;margin-bottom:8px">Один объект</strong>
  <div class="row">
    <select id="selSchool"></select>
    <span class="tag" id="tagR"></span>
    <span class="note" id="metaSchool" style="margin:0"></span>
  </div>
  <div class="chart"><canvas id="cTraj"></canvas></div>
  <p class="note" id="ktLine"></p>
  <p class="note">Синяя — факт, пунктир — план, зелёная — % от суммы контракта, уже выплаченной на эту дату.</p>
"""

DASHBOARD_SCRIPT = r"""</div>
<script>
const DATA = __DATA__;
const blue='#126880', blueL='rgba(18,104,128,.35)', green='#27AE60';

const STATUS_LABEL = { over:'Избыток оплаты', credit:'Кредитует подрядчик', balanced:'Баланс', unknown:'Нет данных' };
const STATUS_PILL = { over:'err', credit:'warn', balanced:'ok', unknown:'' };
function moneyCell(v) {
  if (v==null) return '—';
  const cls = v>0 ? 'pos' : (v<0 ? 'neg' : '');
  return `<span class="money ${cls}">${v>0?'+':''}${v.toLocaleString('ru-RU')}</span>`;
}

document.getElementById('dK1').textContent = DATA.objects.length;
document.getElementById('dK2').textContent = (DATA.objects.reduce((s,o)=>s+(o.contract_value||0),0)/1000).toFixed(1);
document.getElementById('dK3').textContent = (-DATA.objects.filter(o=>o.money_status==='credit').reduce((s,o)=>s+(o.gap_rub||0),0)/1000).toFixed(1);
document.getElementById('dK4').textContent = DATA.objects.filter(o=>o.flags && o.flags.includes('не осваивает бюджет 2026')).length;

let activeFilter = 'all', activeContractor = null;

function passesFilter(o) {
  if (activeContractor && o.contractor !== activeContractor) return false;
  if (activeFilter==='credit') return o.money_status==='credit' && Math.abs(o.gap_rub||0)>100;
  if (activeFilter==='balanced') return o.money_status==='balanced';
  if (activeFilter==='nobudget') return o.flags && o.flags.includes('не осваивает бюджет 2026');
  return true;
}

function renderObjTbl() {
  const rows = DATA.objects.filter(passesFilter).sort((a,b)=>Math.abs(b.gap_rub||0)-Math.abs(a.gap_rub||0));
  document.getElementById('objTbl').innerHTML = rows.map(o => {
    const days = o.days_to_open;
    const overdue = days!=null ? (days<0 ? `<span class="pill err">просрочка ${Math.abs(days)} дн.</span>` : `<span class="pill ok">осталось ${days} дн.</span>`) : '—';
    return `<tr class="clickable" data-uin="${o.uin}"><td title="${o.full}">${o.name}</td><td>${o.municipality||'—'}</td><td>${o.contractor||'—'}</td>` +
      `<td class="r">${o.contract_value??'—'}</td><td class="r">${o.sg}%</td><td class="r">${o.pct??'—'}%</td>` +
      `<td class="r">${moneyCell(o.gap_rub)}</td><td><span class="pill ${STATUS_PILL[o.money_status]}">${STATUS_LABEL[o.money_status]}</span></td><td>${overdue}</td></tr>`;
  }).join('');
  document.querySelectorAll('#objTbl tr.clickable').forEach(tr => tr.onclick = () => { sel.value = tr.dataset.uin; drawTraj(tr.dataset.uin); });
}

document.querySelectorAll('#filterBar .fbtn').forEach(btn => btn.onclick = () => {
  document.querySelectorAll('#filterBar .fbtn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  activeFilter = btn.dataset.f;
  renderObjTbl();
});

function renderRollup() {
  const byC = {};
  DATA.objects.forEach(o => {
    const c = o.contractor || '—';
    const e = byC[c] || (byC[c] = { contractor:c, n:0, contract:0, credit:0, no_budget2026:0 });
    e.n++; e.contract += o.contract_value||0;
    if (o.money_status==='credit') e.credit += -(o.gap_rub||0);
    if (o.flags && o.flags.includes('не осваивает бюджет 2026')) e.no_budget2026++;
  });
  const rows = Object.values(byC).sort((a,b)=>b.credit-a.credit);
  document.getElementById('contrRollup').innerHTML = rows.map(c =>
    `<tr class="clickable${activeContractor===c.contractor?' row-active':''}" data-c="${c.contractor}"><td>${c.contractor}</td><td class="r">${c.n}</td><td class="r">${Math.round(c.contract)}</td><td class="r">${moneyCell(-Math.round(c.credit))}</td><td class="r">${c.no_budget2026||'—'}</td></tr>`
  ).join('');
  document.querySelectorAll('#contrRollup tr.clickable').forEach(tr => tr.onclick = () => {
    activeContractor = activeContractor===tr.dataset.c ? null : tr.dataset.c;
    renderRollup(); renderObjTbl();
  });
}

renderObjTbl();
renderRollup();

const sel = document.getElementById('selSchool');
DATA.objects.forEach(o=>{
  const opt=document.createElement('option'); opt.value=o.uin;
  opt.textContent=o.name; sel.appendChild(opt);
});
sel.value = DATA.defaultUin;
let chartTraj;

function drawTraj(uin) {
  const rows = DATA.traj[uin]||[], m = DATA.per.find(p=>p.uin===uin);
  const k = DATA.kt_dates[uin]||{};
  document.getElementById('tagR').textContent = m&&m.r!=null ? 'корр. '+m.r.toFixed(2) : '';
  document.getElementById('metaSchool').textContent = m ? `${m.name}: ${m.sg}% факт, ${m.plan??'—'}% план, ${m.pct??'—'}% выпл.` : '';
  document.getElementById('ktLine').textContent = `Экспертиза ${k.exp_sl||k.exp_plan||'—'}, СМР с ${k.smr_start||'—'}, контракт ${k.ctr_fact||'—'}` + (m&&m.flags?.length ? '. '+m.flags.join(', ') : '');
  const annotations = {};
  if (k.exp_marker) {
    annotations.exp = { type:'line', xMin:k.exp_marker, xMax:k.exp_marker, borderColor:'#a04ea3', borderWidth:2, borderDash:[3,3],
      label:{ display:true, content:'Экспертиза', position:'start', backgroundColor:'#a04ea3', font:{size:10} } };
  }
  const cfg = { type:'line', data:{ labels:rows.map(r=>r.d), datasets:[
    { label:'Факт', data:rows.map(r=>r.sg), borderColor:blue, tension:.25, pointRadius:2 },
    { label:'План', data:rows.map(r=>r.plan), borderColor:blue, borderDash:[5,4], tension:.25, pointRadius:0 },
    { label:'Основные платежи', data:rows.map(r=>r.regPct), borderColor:green, tension:.25, pointRadius:2 },
    { label:'Аванс', data:rows.map(r=>r.advPct), borderColor:green, borderDash:[2,2], tension:.25, pointRadius:0 }
  ]}, options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom'}, datalabels:{display:false}, annotation:{annotations}}, scales:{ y:{min:0,max:100} } } };
  if(chartTraj) chartTraj.destroy(); chartTraj = new Chart(document.getElementById('cTraj'), cfg);
}
sel.onchange = e => drawTraj(e.target.value);

drawTraj(sel.value);
</script>
</body>
</html>
"""

METHOD_SCRIPT = r"""</div>
<script>
const DATA = __DATA__;
const blue='#126880', blueL='rgba(18,104,128,.35)', green='#27AE60';

document.getElementById('k6').textContent = DATA.stats.n_no_budget2026;

document.getElementById('budgetAlert').innerHTML = DATA.budget_alert.map(s =>
  `<tr><td title="${s.full}">${s.name}</td><td class="r">${s.sg}%</td><td class="r">${s.plan2026}</td></tr>`
).join('');

document.getElementById('contractorsTbl').innerHTML = DATA.contractors.map(c => {
  const objs = c.objects.map(o => o.name + (o.no_budget2026 ? ' *' : '')).join(', ');
  const allStuck = c.objects.length > 1 && c.n_no_budget2026 === c.objects.length;
  return `<tr${allStuck ? ' style="background:var(--err-bg)"' : ''}><td>${c.contractor}</td><td class="r">${c.objects.length}</td><td class="r">${c.n_no_budget2026 || '—'}</td><td>${objs}</td></tr>`;
}).join('');

document.getElementById('redZoneTbl').innerHTML = DATA.red_zone.filter(s=>s.deviation>15).map(s =>
  `<tr><td title="${s.full}">${s.name}</td><td class="r">${s.sg}%</td><td class="r">${s.gap}</td><td class="r">${s.stage_median}</td><td class="r">+${s.deviation}</td></tr>`
).join('');

document.getElementById('advOnlyN').textContent = DATA.advance_only.length;
document.getElementById('advOnlyTbl').innerHTML = [...DATA.advance_only].sort((a,b)=>b.sg-a.sg).map(s =>
  `<tr><td title="${s.full}">${s.name}</td><td class="r">${s.sg}%</td><td class="r">${s.advPct}%</td></tr>`
).join('');

const unbacked = DATA.objects.filter(o=>o.program_unbacked>0).sort((a,b)=>b.program_unbacked-a.program_unbacked);
document.getElementById('unbackedN').textContent = unbacked.length;
document.getElementById('unbackedTbl').innerHTML = unbacked.map(o =>
  `<tr><td title="${o.full}">${o.name}</td><td class="r">${o.sg}%</td><td class="r">${o.program_unbacked}</td></tr>`
).join('');

function initStage() {
  const s = DATA.stage_analysis;
  new Chart(document.getElementById('cStage'), {
    type:'line',
    data:{ labels:s.map(x=>x.stage+'%'), datasets:[
      { label:'СГ обгоняет выплаты, п.п.', data:s.map(x=>x.pay_gap), borderColor:green, backgroundColor:green, tension:.2, pointRadius:4 },
      { label:'Факт отстаёт от плана, п.п.', data:s.map(x=>x.plan_gap), borderColor:blue, backgroundColor:blue, borderDash:[5,4], tension:.2, pointRadius:4 }
    ]},
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom'}, datalabels:{display:false},
      tooltip:{callbacks:{afterLabel:c=>'школ на этом этапе: '+s[c.dataIndex].n}}},
      scales:{ x:{title:{display:true,text:'Этап готовности (СГ)'}}, y:{title:{display:true,text:'Разрыв, п.п.'}} } }
  });
  const first = s.find(x=>x.n>=10), last = [...s].reverse().find(x=>x.n>=10);
  document.getElementById('stageNote').textContent = first && last
    ? `На старте (${first.stage}% готовности) деньги опережают стройку на ${Math.abs(first.pay_gap)} п.п. — это аванс. К ${last.stage}% готовности стройка опережает деньги уже на ${last.pay_gap} п.п., и разрыв растёт быстрее всего после 60–70% готовности. С планом — обратная картина: сильнее всего школы отстают от собственного графика в середине стройки, а к концу почти нагоняют.`
    : '';
}

function initMethod() {
  const s = DATA.stage_analysis;
  const late = [...s].reverse().find(x=>x.n>=10 && x.stage>=80) || [...s].reverse().find(x=>x.n>=10);
  document.getElementById('mth90').textContent = late ? late.pay_gap : '—';
  document.getElementById('mth0').textContent = DATA.stats.n_no_budget2026;
  document.getElementById('mthRedN').textContent = DATA.stats.n_red_zone;
  document.getElementById('conclN0').textContent = DATA.stats.n_no_budget2026;
  document.getElementById('conclNRed').textContent = DATA.stats.n_red_zone;

  const stuck = DATA.contractors.filter(c => c.objects.length > 1 && c.n_no_budget2026 === c.objects.length);
  document.getElementById('mthContr').textContent = stuck.length
    ? `У ${stuck.map(c=>c.contractor+' ('+c.n_no_budget2026+' из '+c.objects.length+' объектов)').join(', ')} без исполнения 2026 года стоит каждый объект. Разбор здесь нужен на уровне договора с подрядчиком, а не по одной школе за раз.`
    : 'На этой выгрузке подрядчиков, у которых стоят все объекты сразу, нет — если появится хотя бы один, это приоритетный сигнал.';
}

initStage();
initMethod();
</script>
</body>
</html>
"""


def main():
    payload = load_data()
    data_json = json.dumps(payload, ensure_ascii=False)
    sub = "48 школ, обновлено 07.09.2026"

    method_html = (
        HEAD_STYLE.replace("__TITLE__", "СГ и выплаты — методика")
        .replace("__SUB__", sub + " — методика управления финансированием портфеля")
        .replace("__NAV__", '<a class="navlink" href="dashboard.html">→ Дашборд по объектам</a>')
        .replace("__WRAP__", "880")
        + METHOD_BODY
        + METHOD_SCRIPT.replace("__DATA__", data_json)
    )
    dashboard_html = (
        HEAD_STYLE.replace("__TITLE__", "СГ и выплаты — дашборд")
        .replace("__SUB__", sub + " — объекты, подрядчики, деньги против готовности")
        .replace("__NAV__", '<a class="navlink" href="index.html">→ Методика и выводы</a>')
        .replace("__WRAP__", "1440")
        + DASHBOARD_BODY
        + DASHBOARD_SCRIPT.replace("__DATA__", data_json)
    )

    out = ROOT / "index.html"
    out.write_text(method_html, encoding="utf-8")
    (ROOT / "sg-pay-analysis.html").write_text(method_html, encoding="utf-8")
    dash_out = ROOT / "dashboard.html"
    dash_out.write_text(dashboard_html, encoding="utf-8")
    print(f"OK: {out} ({out.stat().st_size} bytes) + {dash_out} ({dash_out.stat().st_size} bytes), {payload['stats']['n']} schools")


if __name__ == "__main__":
    main()

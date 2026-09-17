#!/usr/bin/env python3
"""Generate index.html from Schools xlsx data."""
import base64
import json
import math
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent
DATA_DIR = ROOT / "data"
RAW_DIR = DATA_DIR / "raw"
CKS_LOGO_B64 = base64.b64encode((ROOT / "assets" / "cks-logo.png").read_bytes()).decode("ascii")

# Lucide alert-triangle, stroke 1.5px — брендбук ЦКС запрещает emoji в роли иконок.
ICON_WARN = (
    '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
    'stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" '
    'style="vertical-align:-2px;margin-right:2px"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/>'
    '<path d="M12 9v4"/><path d="M12 17h.01"/></svg>'
)

ADVANCE_PCTS = {30.0, 49.0}  # типовые проценты аванса в плане СГ, а не расчётный факт оплаты
# Кластер «типовой аванс, не прогресс»: % выплат застрял около размера аванса по контракту.
ADVANCE_STUCK_BAND = (26.5, 27.5)

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


def percentile(xs, p):
    s = sorted(xs)
    if not s:
        return 0.0
    if len(s) == 1:
        return float(s[0])
    k = (len(s) - 1) * p / 100
    f = math.floor(k)
    c = math.ceil(k)
    if f == c:
        return float(s[int(k)])
    return float(s[f] * (c - k) + s[c] * (k - f))


def iqr(xs):
    if len(xs) < 4:
        return 0.0
    return percentile(xs, 75) - percentile(xs, 25)


def weighted_median(pairs):
    """pairs: list of (value, weight). Returns None if empty."""
    if not pairs:
        return None
    items = sorted(((float(v), float(w)) for v, w in pairs if w and w > 0), key=lambda x: x[0])
    if not items:
        return median([v for v, _ in pairs])
    total_w = sum(w for _, w in items)
    half = total_w / 2
    acc = 0.0
    for v, w in items:
        acc += w
        if acc >= half:
            return v
    return items[-1][0]


def is_advance_stuck(pay_pct, last_pt):
    """True if % выплат — типовой аванс, а не прогресс по актам."""
    if pay_pct is not None and ADVANCE_STUCK_BAND[0] <= pay_pct <= ADVANCE_STUCK_BAND[1]:
        return True
    if last_pt:
        adv, reg = last_pt.get("advPct"), last_pt.get("regPct")
        # Только аванс, платежей по актам нет; % выплат совпадает с долей аванса.
        if adv and adv > 0 and (not reg or reg < 0.5) and pay_pct is not None:
            if abs(pay_pct - adv) < 1.5:
                return True
    return False


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
    path = RAW_DIR / "1708_КСГ+Экспертиза.xlsx"
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


SIMPLE_LIST_FILE = "1409_Акцент_Simple List.xlsx"


def load_simple_list():
    wb = openpyxl.load_workbook(RAW_DIR / SIMPLE_LIST_FILE, data_only=True)
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

        # Раньше строка целиком пропускалась (continue), если её "Год финансирования" был <=
        # уже сохранённого — расчёт на то, что строки идут по возрастанию года. На практике
        # это не так: у объекта бывает несколько строк с ОДНИМ и тем же годом (более полная
        # строка позже) и строка с "Год финансирования" = "0" (сводная/статусная, не привязана
        # к году, может идти в файле после строк с реальными годами) — такие строки отбрасывались
        # целиком вместе с полями, которых больше нигде нет (согласованная стоимость экспертизы,
        # даты экспертизы). Теперь обрабатываем каждую строку: поля объекта (подрядчик, стоимости,
        # даты) берём из ЛЮБОЙ строки, где они заполнены (later wins при повторении, prev — фолбэк
        # при пропуске), а "самый свежий год" (для pay_pct и yr) продвигаем только вперёд.
        yr_num = None
        if isinstance(yr, (int, float)):
            yr_num = int(yr)
        elif yr is not None and str(yr).strip().isdigit():
            yr_num = int(str(yr).strip())
        yr_is_real = yr_num is not None and yr_num >= 1
        prev = sl.get(uin) or {}
        pay_pct = prev.get("pay_pct")
        if not yr_is_real or yr_num >= (prev.get("yr") or 0):
            v = row[col["Процент выплат"]]
            if v is not None:
                try:
                    pay_pct = round(float(str(v).replace(",", ".")), 1)
                except ValueError:
                    pass
        contractor = row[col["Наименование подрядчика"]] or prev.get("contractor")
        inn = row[col["ИНН подрядчика"]] or prev.get("inn")
        rp = row[col["РП"]] or prev.get("rp")
        entered_exp = bool(row[col["Код заявления на прохождение экспертизы"]]) or prev.get("entered_exp", False)

        exp_overrun = prev.get("exp_overrun")
        plan_c = num("Плановая стоимость объекта экспертизы") or prev.get("plan_cost")
        agreed_c = num("Согласованная стоимость объекта экспертизы") or prev.get("agreed_cost")
        if plan_c and agreed_c and plan_c > 0:
            exp_overrun = round((agreed_c - plan_c) / plan_c * 100, 1)

        contract_value = (
            num("Начальная максимальная цена контракта")
            or num("Предельная стоимость по объекту, тыc.руб.")
            or prev.get("contract_value")
        )
        agreed_cost = agreed_c or prev.get("agreed_cost")

        sl[uin] = {
            "yr": yr_num if yr_is_real and yr_num >= (prev.get("yr") or 0) else prev.get("yr"),
            "name": row[col["Название объекта"]] or prev.get("name"),
            "pay_pct": pay_pct,
            "contractor": contractor,
            "inn": inn,
            "rp": rp,
            "contract_value": contract_value,
            "agreed_cost": agreed_cost,
            "plan_cost": plan_c,
            "entered_exp": entered_exp,
            "exp_overrun": exp_overrun,
            "exp_plan_entry": parse_date(row[col["Плановая дата захода на экспертизу из ДК"]]) or prev.get("exp_plan_entry"),
            "opening_plan": parse_date(row[col["Планируемая дата открытия"]]) or prev.get("opening_plan"),
            "exp_in": parse_date(row[col["Дата подачи заявления (захода) на экспертизу"]]) or prev.get("exp_in"),
            "exp_start": parse_date(row[col["Дата начала экспертизы"]]) or prev.get("exp_start"),
            "exp_done": parse_date(row[col["Дата получения заключения (завершения ) экспертизы"]]) or prev.get("exp_done"),
            "ctr_plan": parse_date(row[col["Заключение контракта начало план КСГ"]]) or prev.get("ctr_plan"),
            "ctr_fact": parse_date(row[col["Заключение контракта начало факт КСГ"]]) or prev.get("ctr_fact"),
        }

    for uin, rows in years.items():
        by_year = {}
        for r in rows:
            e = by_year.setdefault(r["year"], {"year": r["year"], "plan": 0, "obligated": 0, "fact": 0, "unbacked": 0})
            e["plan"] += r["plan"]
            e["obligated"] += r["obligated"]
            e["fact"] += r["fact"]
            # Каждая строка своя (разные источники финансирования внутри года) — если считать
            # "без обязательств" по сумме за год, необеспеченная строка маскируется другой
            # строкой с обязательствами за тот же год. Проверяем каждую строку отдельно.
            if r["obligated"] == 0 and r["fact"] == 0:
                e["unbacked"] += r["plan"]
        if uin in sl:
            sl[uin]["program_years"] = sorted(by_year.values(), key=lambda x: x["year"])
    return sl


def load_finance2026():
    path = DATA_DIR / "finance2026.json"
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def load_addresses():
    path = DATA_DIR / "addresses.json"
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def load_per_object():
    path = DATA_DIR / "per_object.json"
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


OPEN_EXPERTISE_STAGES = {
    "Ожидание устранения замечаний", "Рассмотрение ПД", "Подготовка заключения",
    "Ожидание загрузки документации", "Ожидание возврата договора", "Обработка",
}


def load_pir():
    path = RAW_DIR / "1409_Выгрузка_ПИР.xlsx"
    if not path.exists():
        return {}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    col = {h: i for i, h in enumerate(headers)}
    by_uin = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        uin = row[col["УИН"]]
        if not uin:
            continue
        by_uin.setdefault(uin, []).append(
            {
                "stage": row[col["Стадия рассмотрения"]],
                "result": row[col["Результат экспертизы"]],
                "egrz_date": parse_date(row[col["Дата заключения ЕГРЗ. Дата"]]),
                "pir_start_fact": parse_date(row[col["ПИР Дата начала факт"]]),
                "pir_end_fact": parse_date(row[col["ПИР Дата окончания факт"]]),
                "pir_end_plan": parse_date(row[col["ПИР Дата начала план"]]),
            }
        )

    out = {}
    for uin, rows in by_uin.items():
        concluded = [r for r in rows if r["stage"] == "Услуга оказана" and r["egrz_date"]]
        last = max(concluded, key=lambda r: r["egrz_date"]) if concluded else None
        pir_end = max((r["pir_end_fact"] for r in rows if r["pir_end_fact"]), default=None)
        out[uin] = {
            "exp_last_result": last["result"] if last else None,
            "exp_last_date": fmt_date(last["egrz_date"]) if last else None,
            "exp_pending": any(r["stage"] in OPEN_EXPERTISE_STAGES for r in rows),
            "pir_end_fact": fmt_date(pir_end),
        }
    return out


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
    pir = load_pir()

    cross, per, traj, kt_rows, kt_dates = [], [], {}, [], {}
    budget_alert = []
    STAGE_EDGES = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
    stage_gaps = {e: {"plan": [], "pay": [], "pay_w": []} for e in STAGE_EDGES}

    per_object = load_per_object()
    for uin, rec in sorted(per_object.items()):
        info = sl.get(uin, {})
        name = NAME_OVERRIDES.get(uin, info.get("name", uin))
        series = [
            {"d": parse_date(r["d"]), "plan": r["plan"], "fact": r["fact"]} for r in rec["series"]
        ]
        if not series:
            continue

        pays = [
            {"d": parse_date(r["d"]), "amt": r["amt"], "advance": r["advance"]}
            for r in rec["pays"]
        ]
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
                gap = hit["sg"] - hit["payPct"]
                stage_gaps[edge]["pay"].append(gap)
                cv = info.get("contract_value") or 0
                if cv > 0:
                    stage_gaps[edge]["pay_w"].append((gap, cv))

        if len(pts) > 36:
            step = math.ceil(len(pts) / 36)
            pts = pts[::step][:-1] + [pts[-1]]
        traj[uin] = pts

        last_pt = pts[-1] if pts else None
        advance_stuck = is_advance_stuck(pay_pct, last_pt)

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
        days_to_open = (opening_plan.date() - datetime.now().date()).days if opening_plan else None

        program_years = info.get("program_years", [])
        program_unbacked = round(
            sum(y.get("unbacked", 0) for y in program_years if y["year"] <= 2026) / 1e6,
            1,
        )

        contract_value = info.get("contract_value")

        # Лимит по годам из госпрограммы (сумма program_years.plan) может превышать заключённый
        # контракт — это бюджетные обязательства, присвоенные объекту, но не привязанные ни к
        # одному контракту (см. transcript 21: "лимит 150, контракт заключён на 100"). Разница —
        # не ошибка, это может быть будущий этап без контракта, но её стоит видеть отдельно.
        limit_total = sum(y["plan"] for y in program_years)
        limit_gap = round((limit_total - contract_value) / 1e6, 1) if contract_value else None
        if limit_gap is not None and limit_gap > 20:
            flags.append("лимит превышает контракт")

        # Погодовая разбивка платежей на аванс/исполнение (акты) — из истории платежей
        # объекта, а не из program_years (которая берёт план/факт из Simple List и не различает
        # тип платежа).
        pay_years_acc = {}
        for p in pays:
            e = pay_years_acc.setdefault(p["d"].year, {"year": p["d"].year, "advance": 0.0, "act": 0.0})
            if p["advance"]:
                e["advance"] += p["amt"]
            else:
                e["act"] += p["amt"]
        pay_years = [
            {"year": y["year"], "advance": round(y["advance"] / 1e6, 1), "act": round(y["act"] / 1e6, 1)}
            for y in sorted(pay_years_acc.values(), key=lambda x: x["year"])
        ]

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

        # Финансовая аналитика объекта: остаток физической работы (в рублях по цене контракта)
        # против остатка бюджета, ещё не потраченного в 2026+ году. Если денег заметно меньше,
        # чем нужно на остаток работ — недофинансирование, нужно искать откуда добавить.
        # Если заметно больше — план завышен, можно снять и перекинуть на другой объект.
        budget_mismatch = None
        if contract_value:
            remaining_work_rub = (100 - last["fact"]) / 100 * contract_value
            future_years = [y for y in program_years if y["year"] >= 2026]
            if future_years:
                remaining_plan_rub = sum(y["plan"] - y["fact"] for y in future_years)
                budget_mismatch = round((remaining_plan_rub - remaining_work_rub) / 1e6, 1)

        pir_info = pir.get(uin, {})
        if pir_info.get("exp_last_result") == "Отрицательное":
            flags.append("экспертиза отклонена")
        elif pir_info.get("exp_pending"):
            flags.append("экспертиза на пересмотре")

        # Положительное заключение экспертизы подтверждает только техчасть (ОПД) — смета (СД)
        # заключается отдельно и может отставать. Подтверждаем СД по Simple List: если там
        # заполнена "Согласованная стоимость объекта экспертизы", смета согласована; если общий
        # результат "Положительное", а согласованной стоимости нет — СД ещё не закрыта.
        sd_confirmed = bool(info.get("agreed_cost"))
        if pir_info.get("exp_last_result") == "Положительное" and not sd_confirmed:
            flags.append("смета (СД) не подтверждена")
        if advance_stuck:
            flags.append("типовой аванс")

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
                "limit_gap": limit_gap,
                "pay_years": pay_years,
                "contract_value": round(contract_value / 1e6, 1) if contract_value else None,
                "gap_pct": gap_pct,
                "gap_rub": gap_rub,
                "money_status": money_status,
                "budget_mismatch": budget_mismatch,
                "exp_last_result": pir_info.get("exp_last_result"),
                "exp_last_date": pir_info.get("exp_last_date"),
                "exp_pending": pir_info.get("exp_pending", False),
                "pir_end_fact": pir_info.get("pir_end_fact"),
                "sd_confirmed": sd_confirmed,
                "advance_stuck": advance_stuck,
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
                "advance_stuck": advance_stuck,
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
        iqr_val = iqr(pay_vals)
        # 0.75×IQR: адаптивно по этапу (шире разброс → выше порог), но 1.5×IQR
        # на этом портфеле опустошает красную зону (IQR на 90% ≈ 24 п.п.).
        threshold = round(max(8.0, 0.75 * iqr_val), 1) if pay_vals else None
        pay_w = weighted_median(stage_gaps[edge]["pay_w"])
        stage_analysis.append(
            {
                "stage": edge,
                "n": max(len(plan_vals), len(pay_vals)),
                "plan_gap": round(median(plan_vals), 1) if plan_vals else None,
                "pay_gap": round(median(pay_vals), 1) if pay_vals else None,
                "pay_gap_w": round(pay_w, 1) if pay_w is not None else None,
                "iqr": round(iqr_val, 1) if pay_vals else None,
                "threshold": threshold,
            }
        )

    stage_by_edge = {s["stage"]: s["pay_gap"] for s in stage_analysis if s["pay_gap"] is not None}
    stage_threshold = {s["stage"]: s["threshold"] for s in stage_analysis if s["threshold"] is not None}

    red_zone = []
    for k in kt_rows:
        if k["sg_pay_gap"] is None:
            continue
        stage = max((e for e in STAGE_EDGES if e <= k["sg"] and e in stage_by_edge), default=None)
        if stage is None:
            continue
        thr = stage_threshold.get(stage, 15)
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
                "threshold": thr,
                "in_red": deviation > thr,
                "advance_stuck": k.get("advance_stuck", False),
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

    # Единый риск-сигнал "деньги + срок": просто "просрочен" не различает объекты — почти
    # весь портфель уже просрочен одинаково (типовые -9/-26 дней от единой плановой даты
    # ввода). Значимый сигнал — не факт просрочки, а то, что объект просрочен СИЛЬНЕЕ
    # типичного по портфелю И при этом подрядчик уже кредитует стройку деньгами: это
    # компаунд-риск "не успеет достроить, потому что не хватает денег", а не сезонная
    # просрочка, которая есть у всех.
    overdue_vals = [-k["days_to_open"] for k in kt_rows if (k["days_to_open"] or 0) < 0]
    overdue_median = median(overdue_vals) if overdue_vals else 0
    for k in kt_rows:
        d = k["days_to_open"]
        overdue_days = -d if d is not None and d < 0 else 0
        k["urgent_risk"] = bool(
            k["money_status"] == "credit"
            and overdue_days > overdue_median
            and (k["gap_rub"] or 0) < -50
        )

    # Парето по денежному риску: топ объектов по |gap_rub| среди тех, где подрядчик
    # кредитует стройку, с накопленной долей от суммы риска всех "credit"-объектов —
    # чтобы видеть, сколько объектов покрывает большую часть денежного риска портфеля.
    credit_rows = [k for k in kt_rows if k["money_status"] == "credit" and k["gap_rub"]]
    credit_rows.sort(key=lambda x: x["gap_rub"])
    total_credit_risk = sum(-k["gap_rub"] for k in credit_rows)
    pareto = []
    cum = 0.0
    for k in credit_rows[:15]:
        cum += -k["gap_rub"]
        pareto.append(
            {
                "uin": k["uin"],
                "name": k["name"],
                "full": k["full"],
                "gap_rub": k["gap_rub"],
                "cum_pct": round(cum / total_credit_risk * 100, 1) if total_credit_risk else None,
            }
        )

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
            "n_red_zone": sum(1 for r in red_zone if r["in_red"]),
            "n_advance_stuck": sum(1 for o in kt_rows if o.get("advance_stuck")),
            "n_exp_failed": sum(1 for o in kt_rows if o["exp_last_result"] == "Отрицательное" or o["exp_pending"]),
            "n_urgent_risk": sum(1 for o in kt_rows if o["urgent_risk"]),
            "n_limit_gap": sum(1 for o in kt_rows if (o["limit_gap"] or 0) > 20),
        },
        "budget_alert": budget_alert,
        "contractors": contractors,
        "stage_analysis": stage_analysis,
        "red_zone": red_zone,
        "advance_only": advance_only,
        "pareto": pareto,
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
<link href="https://fonts.googleapis.com/css2?family=Golos+Text:wght@400;500;600;700&family=Manrope:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-annotation@3.0.1/dist/chartjs-plugin-annotation.min.js"></script>
<style>
  :root {
    --bg:#F4F7FA; --surface:#fff; --div:#EAF0F5; --line:#D6E2EC;
    --text:#0D2040; --muted:#5A7189; --faint:#8BA4B8;
    --navy:#143260;
    --accent:#1B8A9C; --accent-l:#22B0C8; --accent-d:#126880; --accent-dim:rgba(27,138,156,.09);
    --ok:#27AE60; --warn:#E8A020; --err:#D94040; --info:#2E7CC4;
    --ok-bg:#E8F6EE; --warn-bg:#FDF3E2; --err-bg:#FBE9E9; --info-bg:#E9F1FB;
  }
  * { box-sizing:border-box }
  body { margin:0; font:15px/1.55 'Golos Text','Manrope',system-ui,sans-serif; background:var(--bg); color:var(--text); letter-spacing:-.01em }
  .wrap { max-width:__WRAP__px; margin:0 auto; padding:28px 18px 56px }
  h1 { font-size:1.65rem; font-weight:700; margin:0 0 4px; letter-spacing:-.02em; color:var(--navy) }
  .sub { color:var(--muted); margin:0 0 20px; font-size:.92rem }
  .brand { display:block; height:34px; margin-bottom:14px }
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
  .box.stub { background:repeating-linear-gradient(135deg, var(--surface) 0 10px, #EAF0F5 10px 20px); border-style:dashed }
  .mini table { width:100%; font-size:.84rem; border-collapse:collapse }
  .mini th,.mini td { padding:7px 8px; border-bottom:1px solid var(--div); text-align:left }
  .mini th { color:var(--faint); font-weight:600; font-size:.72rem; letter-spacing:.04em; text-transform:uppercase; background:#EAF0F5 }
  .mini td.r,.mini th.r { text-align:right; font-variant-numeric:tabular-nums }
  details { background:var(--surface); border:1px solid var(--line); border-radius:12px; margin:10px 0; overflow:hidden; box-shadow:0 1px 3px rgba(13,32,64,.05) }
  details > summary { cursor:pointer; padding:14px 18px; font-weight:600; list-style:none; user-select:none }
  details > summary::-webkit-details-marker { display:none }
  details > summary::after { content:'+'; float:right; color:var(--faint); font-weight:400 }
  details[open] > summary::after { content:'−' }
  details > summary:hover { background:#EAF0F5 }
  .detail-body { padding:0 18px 18px; border-top:1px solid var(--div) }
  select { font:inherit; padding:6px 10px; border:1px solid var(--line); border-radius:8px; background:#fff; max-width:100% }
  .row { display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin:10px 0 }
  .tag { font-size:.75rem; padding:3px 9px; background:var(--accent-dim); border:1px solid var(--line); border-radius:99px; color:var(--accent-d); font-weight:500 }
  table.full { width:100%; border-collapse:collapse; font-size:.84rem }
  table.full th,table.full td { padding:8px 9px; border-bottom:1px solid var(--div) }
  table.full th { background:#EAF0F5; position:sticky; top:0; text-align:left; color:var(--faint); font-weight:600; font-size:.72rem; letter-spacing:.04em; text-transform:uppercase }
  table.full td.r,table.full th.r { text-align:right; font-variant-numeric:tabular-nums }
  .tbl-wrap { max-height:420px; overflow:auto; border:1px solid var(--line); border-radius:10px; margin-top:10px }
  .flag { font-size:.72rem; padding:2px 7px; margin:1px 2px 1px 0; display:inline-block; background:var(--warn-bg); border:1px solid #EFCB84; border-radius:5px; color:#8A5E10 }
  .flag.warn { background:var(--err-bg); border-color:#F0B3B3; color:#A32E2E }
  .pill { display:inline-block; font-size:.72rem; font-weight:600; padding:2px 9px; border-radius:99px }
  .pill.ok, .pill.warn, .pill.err { display:inline-flex; align-items:center; gap:5px }
  .pill.ok::before, .pill.warn::before, .pill.err::before { content:''; width:6px; height:6px; border-radius:50%; flex:none }
  .pill.ok { background:var(--ok-bg); color:#1E8449 }
  .pill.ok::before { background:#1E8449 }
  .pill.warn { background:var(--warn-bg); color:#8A5E10 }
  .pill.warn::before { background:#8A5E10 }
  .pill.err { background:var(--err-bg); color:#A32E2E }
  .pill.err::before { background:#A32E2E }
  .stub-label { display:inline-block; font-size:.7rem; font-weight:600; letter-spacing:.03em; text-transform:uppercase; color:var(--faint); background:#EAF0F5; border:1px solid var(--line); border-radius:5px; padding:2px 8px }
  .stub-label.done { color:var(--accent-d); background:var(--accent-dim); border-color:var(--accent-l) }
  .box.done { border-color:var(--accent-l); background:#F3FBFC }
  .kpis.four { grid-template-columns:repeat(4,1fr) }
  @media(max-width:900px) { .kpis.four { grid-template-columns:repeat(2,1fr) } }
  @media(max-width:700px) { .kpis,.kpis.four { grid-template-columns:1fr } }
  .navlink { display:inline-block; margin:0 0 16px; font-size:.85rem; color:var(--accent-d); text-decoration:none; font-weight:600 }
  .navlink:hover { text-decoration:underline }
  .tabbar { display:flex; gap:4px; border-bottom:2px solid var(--line); margin:18px 0 20px }
  .tabbtn { font:inherit; font-size:.95rem; font-weight:700; padding:10px 18px 12px; border:none; border-bottom:3px solid transparent; background:none; color:var(--muted); cursor:pointer; margin-bottom:-2px }
  .tabbtn:hover { color:var(--accent-d) }
  .tabbtn.active { color:var(--accent-d); border-bottom-color:var(--accent) }
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
  <img class="brand" src="data:image/png;base64,__CKS_LOGO__" alt="ЦКС — Центр компетенций строительства">
  <h1>СГ и выплаты</h1>
  <p class="sub">__SUB__</p>
  __NAV__
"""

METHOD_BODY = r"""
  <div class="box" id="methodBox">
    <p class="note" style="margin:0 0 10px">Портфель — 48 капремонтов школ. Ниже — риск по деньгам, а не по проценту готовности: пять правил и решения, которые из них следуют. У каждого правила — цифры конкретно по этому портфелю, без общих слов.</p>

    <strong style="display:block;margin-top:16px;font-size:1.05rem">1. Судить объект по отклонению от нормы для его этапа, не по проценту оплаты</strong>
    <p class="note">Разрыв между готовностью и оплатой растёт по ходу стройки у всех объектов одинаково закономерно: на старте аванс идёт впереди работ, после экспертизы оплата по актам отстаёт. Это норма, а не проблема сама по себе. Проблема — когда конкретный объект отклоняется от этой нормы. На графике две кривые разрыва СГ−выплаты: сплошная — медиана по школам (каждая школа = один голос), пунктир — взвешенная по объёму контракта (типичный рубль портфеля).</p>
    <div class="chart tall"><canvas id="cStage"></canvas></div>
    <p class="note" id="stageNote" style="margin-top:8px"></p>

    <strong style="display:block;margin-top:20px;font-size:1.05rem">2. Красная зона: <span id="mthRedN"></span> школ отклоняются сильнее порога своего этапа</strong>
    <p class="note">Порог считается отдельно на каждом этапе готовности: max(8, 0,75×IQR разрывов). На этапе ~90% норма разрыва — около <span id="mth90"></span> п.п. Школы в таблице отклоняются от нормы сильнее порога этапа. Пометка «типовой аванс» — % выплат застрял около размера аванса, это не прогресс по актам.</p>
    <div class="tbl-wrap" style="max-height:280px">
      <table class="full">
        <thead><tr><th>Школа</th><th class="r">СГ</th><th class="r">Разрыв</th><th class="r">Норма</th><th class="r">Порог</th><th class="r">Отклонение</th><th></th></tr></thead>
        <tbody id="redZoneTbl"></tbody>
      </table>
    </div>

    <strong style="display:block;margin-top:20px;font-size:1.05rem">3. Нулевое освоение решается на этой неделе, не в декабре</strong>
    <p class="note"><span id="mth0"></span> школ — 0% кассового исполнения бюджета 2026 года при уже утверждённом финансировании (89–418 млн ₽ на объект). Ждать конца года бессмысленно: либо деньги начинают двигаться в ближайший месяц, либо бюджет надо честно переносить на 2027-й — и это решение нужно принять сейчас.</p>
    <div class="box" style="background:var(--err-bg);border-color:#F0B3B3;margin:8px 0 0">
      <strong style="display:block;margin-bottom:6px;color:#A32E2E">__ICON_WARN__<span id="k6"></span> школ: бюджет утверждён, выплат в 2026 году не было</strong>
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
        <thead><tr><th>Школа</th><th class="r">СГ</th><th class="r" title="Лимит без обязательств, млн ₽">ЛБО, млн ₽</th></tr></thead>
        <tbody id="unbackedTbl"></tbody>
      </table>
    </div>
    <p class="note">Обратная сторона той же проблемы: у <span id="limitGapN"></span> школ сумма лимита по всем годам госпрограммы больше стоимости заключённого контракта на 20+ млн ₽ — это бюджетные обязательства, присвоенные объекту, но не покрытые ни одним действующим контрактом (не обязательно ошибка — может быть будущий этап, который ещё не законтрактован, но проверить стоит).</p>
    <div class="tbl-wrap" style="max-height:220px">
      <table class="full">
        <thead><tr><th>Школа</th><th class="r">Контракт, млн ₽</th><th class="r">Лимит минус контракт, млн ₽</th></tr></thead>
        <tbody id="limitGapTbl"></tbody>
      </table>
    </div>

    <strong style="display:block;margin-top:20px;font-size:1.05rem">7. У <span id="expFailedN"></span> школ экспертиза не пройдена — вот куда смотреть, почему стоят деньги</strong>
    <p class="note">Появился источник, которого не было раньше (выгрузка ПИР): по каждой заявке на экспертизу видно её реальный результат и историю. У части школ последнее полученное заключение — «Отрицательное», у части сейчас открыта незакрытая заявка на пересмотр. Больше половины этих школ уже были в красной зоне или списке нулевого освоения выше — это, похоже, и есть причина, а не совпадение.</p>
    <div class="tbl-wrap" style="max-height:280px">
      <table class="full">
        <thead><tr><th>Школа</th><th class="r">СГ</th><th>Результат экспертизы</th><th>Дата</th><th>Сейчас на пересмотре</th></tr></thead>
        <tbody id="expFailedTbl"></tbody>
      </table>
    </div>

    <p class="note">«Положительное» заключение закрывает только техническую часть — смета (СД) заключается отдельно и может отставать. Считаем смету подтверждённой, если по объекту в Simple List заполнена «Согласованная стоимость объекта экспертизы»; если её нет при общем результате «Положительное» — деньги по факту не согласованы. <span id="sdGapNote"></span></p>

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
      <li>Причина теперь видна не для всех: у части школ из красной зоны и нулевого списка нашлась конкретная причина — отклонённая экспертиза (раздел 7). Но не для всех — у оставшихся причину (акты не поданы / спор / организационная пробуксовка) всё ещё предстоит выяснить у заказчика и подрядчика напрямую.</li>
      <li>Нет истории по годам: неизвестно, типична ли пробуксовка в начале года — возможно, часть объектов обычно нагоняет в четвёртом квартале, и тогда часть «красной зоны» — не риск, а сезонность.</li>
      <li>Нет данных об условиях контрактов (штрафы, порядок расторжения) — непонятно, какие реальные рычаги есть на переговорах с проблемными подрядчиками.</li>
      <li>Список подрядчиков и бюджет 2026 года сверены вручную один раз по состоянию на начало сентября — при обновлении отчёта их нужно сверять заново, автоматически это не пересчитывается.</li>
      <li>Порог красной зоны — свой на каждый этап (0,75×IQR, не меньше 8 п.п.), а не фиксированные 15 п.п. Эталонная кривая дублируется с весом по контракту. Объекты с «круглым» % выплат около типового аванса помечены отдельно.</li>
    </ul>

    <strong style="display:block;margin-top:24px;font-size:1.1rem">Как считается красная зона и эталон</strong>
    <div class="box done">
      <span class="stub-label done">Сделано</span>
      <p class="note" style="margin-top:8px"><strong>Порог по этапу.</strong> Для каждого этапа готовности берём разброс разрывов СГ−выплаты (IQR) и ставим порог max(8, 0,75×IQR). Красная зона — отклонение от медианы этапа выше этого порога.</p>
    </div>
    <div class="box done">
      <span class="stub-label done">Сделано</span>
      <p class="note" style="margin-top:8px"><strong>Вторая эталонная кривая.</strong> Рядом с медианой «по школам» — взвешенная медиана по объёму контракта («типичный рубль портфеля»).</p>
    </div>
    <div class="box done">
      <span class="stub-label done">Сделано</span>
      <p class="note" style="margin-top:8px"><strong>Флаг «типовой аванс».</strong> % выплат в полосе 26,5–27,5% или совпадение с долей аванса при нулевых платежах по актам. На дашборде — отдельный фильтр.</p>
    </div>
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

  <div class="tabbar" id="tabBar">
    <button class="tabbtn active" data-tab="objects">Объекты</button>
    <button class="tabbtn" data-tab="contractors">Подрядчики и риск</button>
    <button class="tabbtn" data-tab="one">Один объект</button>
  </div>

  <div class="tabpanel" data-tab="objects">
    <div class="filterbar" id="filterBar">
      <button class="fbtn active" data-f="all">Все объекты</button>
      <button class="fbtn" data-f="credit">Подрядчик кредитует &gt;100 млн ₽</button>
      <button class="fbtn" data-f="balanced">Баланс</button>
      <button class="fbtn" data-f="nobudget">0% освоения 2026</button>
      <button class="fbtn" data-f="expfail">Экспертиза отклонена/на пересмотре</button>
      <button class="fbtn" data-f="urgent">Просрочен сильнее типового + кредитует</button>
      <button class="fbtn" data-f="advance">Типовой аванс</button>
    </div>

    <strong style="display:block;margin:18px 0 4px">Матрица риска: готовность vs оплата</strong>
    <p class="note" style="margin-top:0">Каждая точка — школа: по X — стройготовность, по Y — % оплаты по контракту. Пунктирная диагональ — оплата точно по готовности; полоса ±10 п.п. вокруг неё — тот же порог, что делит статусы в таблице ниже. Клик по точке — график этой школы (вкладка «Один объект»).</p>
    <div class="chart" style="height:420px"><canvas id="cMatrix"></canvas></div>

    <div class="tbl-wrap" style="max-height:520px">
      <table class="full">
        <thead><tr>
          <th>Школа</th><th>Округ</th><th>Подрядчик</th><th>РП</th><th class="r">Контракт, млн ₽</th>
          <th class="r">СГ</th><th class="r">Оплата</th><th class="r">Разница, млн ₽</th><th>Статус</th><th>Ввод</th><th class="r">Экспертиза</th>
        </tr></thead>
        <tbody id="objTbl"></tbody>
      </table>
    </div>
    <p class="note">Клик по строке — график этого объекта на вкладке «Один объект». Разница = оплата% минус СГ% × сумма контракта.</p>
  </div>

  <div class="tabpanel" data-tab="contractors" hidden>
    <strong style="display:block;margin-top:4px;margin-bottom:8px">Подрядчики — сводно по портфелю</strong>
    <p class="note" style="margin-top:0">Клик по подрядчику — отфильтровать вкладку «Объекты» только по его объектам.</p>
    <div class="tbl-wrap">
      <table class="full">
        <thead><tr>
          <th>Подрядчик</th><th class="r">Объектов</th><th class="r">Контракт, млн ₽</th>
          <th class="r">Подрядчик кредитует, млн ₽</th><th class="r">Не осваивают бюджет 2026</th>
        </tr></thead>
        <tbody id="contrRollup"></tbody>
      </table>
    </div>

    <strong style="display:block;margin-top:22px;margin-bottom:8px">Концентрация денежного риска</strong>
    <p class="note" style="margin-top:0">Топ объектов по сумме, которую за них уже доплатил подрядчик (среди тех, кто «кредитует» стройку), с накопленной долей от всей такой суммы по портфелю.</p>
    <div class="tbl-wrap" style="max-height:280px">
      <table class="full">
        <thead><tr><th>Школа</th><th class="r">Подрядчик доплатил, млн ₽</th><th class="r">Накопленная доля</th></tr></thead>
        <tbody id="paretoTbl"></tbody>
      </table>
    </div>
  </div>

  <div class="tabpanel" data-tab="one" hidden>
    <div class="row">
      <select id="selSchool"></select>
      <span class="tag" id="tagR"></span>
      <span class="note" id="metaSchool" style="margin:0"></span>
    </div>
    <div class="chart"><canvas id="cTraj"></canvas></div>
    <p class="note" id="ktLine"></p>
    <p class="note">Синяя — факт, пунктир — план, зелёная — % от суммы контракта, уже выплаченной на эту дату.</p>

    <p class="note" style="margin-top:14px"><strong>Лимит по годам (госпрограмма), млн ₽</strong> — план по годам может не совпадать с суммой контракта: контракт заключается на часть лимита, остальное — лимит без обязательств.</p>
    <table class="mini">
      <thead><tr><th>Год</th><th class="r">План</th><th class="r">Обязательства</th><th class="r">Факт</th><th class="r" title="Лимит без обязательств — план, под который ещё не заключён контракт">ЛБО</th></tr></thead>
      <tbody id="programYearsTbl"></tbody>
    </table>
    <p class="note" id="budgetMismatchNote" style="margin-top:8px;font-weight:600"></p>

    <p class="note" style="margin-top:14px"><strong>Платежи по годам: аванс vs исполнение, млн ₽</strong> — из фактической истории платежей объекта, а не из плана.</p>
    <table class="mini">
      <thead><tr><th>Год</th><th class="r">Аванс</th><th class="r">По актам</th></tr></thead>
      <tbody id="payYearsTbl"></tbody>
    </table>
  </div>
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
  if (activeFilter==='expfail') return o.exp_last_result==='Отрицательное' || o.exp_pending;
  if (activeFilter==='urgent') return o.urgent_risk;
  if (activeFilter==='advance') return o.advance_stuck;
  return true;
}

function renderObjTbl() {
  const rows = DATA.objects.filter(passesFilter).sort((a,b)=>Math.abs(b.gap_rub||0)-Math.abs(a.gap_rub||0));
  document.getElementById('objTbl').innerHTML = rows.map(o => {
    const days = o.days_to_open;
    const overdue = days!=null ? (days<0 ? `<span class="pill err">просрочка ${Math.abs(days)} дн.</span>` : `<span class="pill ok">осталось ${days} дн.</span>`) : '—';
    let overrun;
    if (o.exp_last_result==='Отрицательное') overrun = '<span class="pill err">отклонена</span>';
    else if (o.exp_pending) overrun = '<span class="pill warn">на пересмотре</span>';
    else if (o.exp_last_result==='Положительное' && !o.sd_confirmed) overrun = '<span class="pill warn">СД не подтверждена</span>';
    else overrun = o.exp_overrun!=null ? (o.exp_overrun>5 ? `<span class="pill warn">+${o.exp_overrun}%</span>` : o.exp_overrun+'%') : (o.entered_exp?'без удорожания':'не зашла');
    return `<tr class="clickable" data-uin="${o.uin}"><td title="${o.full}">${o.name}${o.advance_stuck?' <span class="flag">аванс</span>':''}</td><td>${o.municipality||'—'}</td><td>${o.contractor||'—'}</td><td>${o.rp||'—'}</td>` +
      `<td class="r">${o.contract_value??'—'}</td><td class="r">${o.sg}%</td><td class="r">${o.pct??'—'}%</td>` +
      `<td class="r">${moneyCell(o.gap_rub)}</td><td><span class="pill ${STATUS_PILL[o.money_status]}">${STATUS_LABEL[o.money_status]}</span></td><td>${overdue}</td><td class="r">${overrun}</td></tr>`;
  }).join('');
  document.querySelectorAll('#objTbl tr.clickable').forEach(tr => tr.onclick = () => { sel.value = tr.dataset.uin; drawTraj(tr.dataset.uin); switchTab('one'); });
}

document.querySelectorAll('#filterBar .fbtn').forEach(btn => btn.onclick = () => {
  document.querySelectorAll('#filterBar .fbtn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  activeFilter = btn.dataset.f;
  renderObjTbl(); renderMatrix();
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
    renderRollup(); renderObjTbl(); renderMatrix(); switchTab('objects');
  });
}

renderObjTbl();
renderRollup();
document.getElementById('paretoTbl').innerHTML = DATA.pareto.map(o =>
  `<tr><td title="${o.full}">${o.name}</td><td class="r">${moneyCell(-o.gap_rub)}</td><td class="r">${o.cum_pct}%</td></tr>`
).join('');

// Матрица риска: та же классификация money_status (гэп % оплаты минус % готовности,
// порог ±10 п.п.), что и в таблице объектов и в STATUS_PILL — просто как точки, а не строки.
// Подчиняется тому же фильтру (activeFilter/activeContractor), что и таблица — один срез
// для всех визуализаций на странице.
const MATRIX_STATUS = {
  over:     { label: 'Избыток оплаты (оплата выше готовности)', color: '#D94040', shape: 'triangle' },
  credit:   { label: 'Кредитует подрядчик (готовность выше оплаты)', color: '#E8A020', shape: 'rect' },
  balanced: { label: 'Баланс (±10 п.п.)', color: '#27AE60', shape: 'circle' },
  unknown:  { label: 'Нет данных по оплате', color: '#93A8BC', shape: 'circle' },
};
let chartMatrix, matrixDatasets;
function renderMatrix() {
  const matrixObjs = DATA.objects.filter(o => o.pct != null && passesFilter(o));
  matrixDatasets = Object.keys(MATRIX_STATUS).map(status => {
    const cfg = MATRIX_STATUS[status];
    const pts = matrixObjs.filter(o => o.money_status === status);
    return {
      label: cfg.label,
      data: pts.map(o => ({ x: o.sg, y: o.pct, uin: o.uin, name: o.name, full: o.full })),
      backgroundColor: cfg.color,
      borderColor: '#fff',
      borderWidth: 2,
      pointStyle: cfg.shape,
      pointRadius: 6,
      pointHoverRadius: 8,
    };
  }).filter(ds => ds.data.length);
  if (chartMatrix) chartMatrix.destroy();
  chartMatrix = new Chart(document.getElementById('cMatrix'), {
    type: 'scatter',
    data: { datasets: matrixDatasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      onClick: (evt, els) => {
        if (!els.length) return;
        const ds = matrixDatasets[els[0].datasetIndex], pt = ds.data[els[0].index];
        sel.value = pt.uin; drawTraj(pt.uin); switchTab('one');
      },
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 12, usePointStyle: true } },
        datalabels: { display: false },
        tooltip: {
          callbacks: {
            title: items => items[0].raw.full,
            label: item => `Готовность ${item.raw.x}%, оплата ${item.raw.y}%`,
          },
        },
        annotation: {
          annotations: {
            diagLine: { type: 'line', xMin: 0, yMin: 0, xMax: 100, yMax: 100, borderColor: '#5A7189', borderWidth: 1, borderDash: [4, 4] },
            diagOver: { type: 'line', xMin: 0, yMin: 10, xMax: 90, yMax: 100, borderColor: '#93A8BC', borderWidth: 1, borderDash: [2, 3] },
            diagCredit: { type: 'line', xMin: 10, yMin: 0, xMax: 100, yMax: 90, borderColor: '#93A8BC', borderWidth: 1, borderDash: [2, 3] },
          },
        },
      },
      scales: {
        x: { min: 0, max: 100, title: { display: true, text: 'Стройготовность, %' } },
        y: { min: 0, max: 100, title: { display: true, text: 'Оплата по контракту, %' } },
      },
    },
  });
}
renderMatrix();

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
    annotations.exp = { type:'line', xMin:k.exp_marker, xMax:k.exp_marker, borderColor:'#2E7CC4', borderWidth:2, borderDash:[3,3],
      label:{ display:true, content:'Экспертиза', position:'start', backgroundColor:'#2E7CC4', font:{size:10} } };
  }
  const cfg = { type:'line', data:{ labels:rows.map(r=>r.d), datasets:[
    { label:'Факт', data:rows.map(r=>r.sg), borderColor:blue, tension:.25, pointRadius:2 },
    { label:'План', data:rows.map(r=>r.plan), borderColor:blue, borderDash:[5,4], tension:.25, pointRadius:0 },
    { label:'Основные платежи', data:rows.map(r=>r.regPct), borderColor:green, tension:.25, pointRadius:2 },
    { label:'Аванс', data:rows.map(r=>r.advPct), borderColor:green, borderDash:[2,2], tension:.25, pointRadius:0 }
  ]}, options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom'}, datalabels:{display:false}, annotation:{annotations}}, scales:{ y:{min:0,max:100} } } };
  if(chartTraj) chartTraj.destroy(); chartTraj = new Chart(document.getElementById('cTraj'), cfg);

  const obj = DATA.objects.find(o=>o.uin===uin);
  const years = (obj && obj.program_years) || [];
  document.getElementById('programYearsTbl').innerHTML = years.length ? years.map(y => {
    const unbacked = y.unbacked || 0;
    return `<tr><td>${y.year}</td><td class="r">${Math.round(y.plan/1e6)}</td><td class="r">${Math.round(y.obligated/1e6)}</td><td class="r">${Math.round(y.fact/1e6)}</td><td class="r">${unbacked>0?'<span class=\"money neg\">'+Math.round(unbacked/1e6)+'</span>':'—'}</td></tr>`;
  }).join('') : '<tr><td colspan="5">Нет данных по годам</td></tr>';

  const payYears = (obj && obj.pay_years) || [];
  document.getElementById('payYearsTbl').innerHTML = payYears.length ? payYears.map(y =>
    `<tr><td>${y.year}</td><td class="r">${y.advance}</td><td class="r">${y.act}</td></tr>`
  ).join('') : '<tr><td colspan="3">Нет данных по платежам</td></tr>';

  const bm = obj ? obj.budget_mismatch : null;
  const bmEl = document.getElementById('budgetMismatchNote');
  if (bm==null) { bmEl.textContent = ''; }
  else if (bm < -20) { bmEl.innerHTML = `__ICON_WARN__Похоже, не хватает денег: физической работы осталось больше, чем запланировано в бюджете на этот и следующие годы — дефицит ≈ ${Math.abs(Math.round(bm))} млн ₽. Надо думать, откуда доставить.`; bmEl.style.color = '#A32E2E'; }
  else if (bm > 50) { bmEl.innerHTML = `План по годам заметно больше, чем нужно на остаток работ (запас ≈ ${Math.round(bm)} млн ₽) — можно снять и перекинуть на другой объект.`; bmEl.style.color = '#8A5E10'; }
  else { bmEl.innerHTML = `План по годам примерно соответствует остатку работ.`; bmEl.style.color = '#1E8449'; }
}
sel.onchange = e => drawTraj(e.target.value);

drawTraj(sel.value);

// Вкладки. Графики Chart.js, созданные в скрытой (hidden) панели, считают её нулевого
// размера и остаются пустыми — при первом показе панели пересчитываем размер явно.
function switchTab(tab) {
  document.querySelectorAll('.tabbtn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  document.querySelectorAll('.tabpanel').forEach(p => p.hidden = p.dataset.tab !== tab);
  if (tab === 'objects' && chartMatrix) chartMatrix.resize();
  if (tab === 'one' && chartTraj) chartTraj.resize();
}
document.querySelectorAll('.tabbtn').forEach(btn => btn.onclick = () => switchTab(btn.dataset.tab));
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

document.getElementById('redZoneTbl').innerHTML = DATA.red_zone.filter(s=>s.in_red).map(s =>
  `<tr><td title="${s.full}">${s.name}</td><td class="r">${s.sg}%</td><td class="r">${s.gap}</td><td class="r">${s.stage_median}</td><td class="r">${s.threshold}</td><td class="r">+${s.deviation}</td><td>${s.advance_stuck?'<span class="flag">типовой аванс</span>':'—'}</td></tr>`
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

const limitGap = DATA.objects.filter(o=>(o.limit_gap||0)>20).sort((a,b)=>b.limit_gap-a.limit_gap);
document.getElementById('limitGapN').textContent = limitGap.length;
document.getElementById('limitGapTbl').innerHTML = limitGap.map(o =>
  `<tr><td title="${o.full}">${o.name}</td><td class="r">${o.contract_value??'—'}</td><td class="r">+${o.limit_gap}</td></tr>`
).join('');

function initStage() {
  const s = DATA.stage_analysis;
  new Chart(document.getElementById('cStage'), {
    type:'line',
    data:{ labels:s.map(x=>x.stage+'%'), datasets:[
      { label:'По школам (медиана)', data:s.map(x=>x.pay_gap), borderColor:green, backgroundColor:green, tension:.2, pointRadius:4 },
      { label:'По рублям (вес контракта)', data:s.map(x=>x.pay_gap_w), borderColor:green, borderDash:[6,4], tension:.2, pointRadius:3, pointStyle:'rect' },
      { label:'Факт отстаёт от плана, п.п.', data:s.map(x=>x.plan_gap), borderColor:blue, backgroundColor:blue, borderDash:[5,4], tension:.2, pointRadius:4 }
    ]},
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom'}, datalabels:{display:false},
      tooltip:{callbacks:{afterLabel:c=>{
        const row = s[c.dataIndex];
        return 'школ: '+row.n+(row.threshold!=null ? ', порог красной зоны: '+row.threshold+' п.п.' : '');
      }}}},
      scales:{ x:{title:{display:true,text:'Этап готовности (СГ)'}}, y:{title:{display:true,text:'Разрыв, п.п.'}} } }
  });
  const first = s.find(x=>x.n>=10), last = [...s].reverse().find(x=>x.n>=10);
  document.getElementById('stageNote').textContent = first && last
    ? `На старте (${first.stage}% готовности) деньги опережают стройку на ${Math.abs(first.pay_gap)} п.п. — это аванс. К ${last.stage}% готовности стройка опережает деньги уже на ${last.pay_gap} п.п. (по рублям: ${last.pay_gap_w??'—'}). Порог красной зоны на ${last.stage}% — ${last.threshold} п.п.`
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

  document.getElementById('expFailedN').textContent = DATA.stats.n_exp_failed;
  const expFailed = DATA.objects.filter(o=>o.exp_last_result==='Отрицательное' || o.exp_pending);
  document.getElementById('expFailedTbl').innerHTML = expFailed.map(o =>
    `<tr><td title="${o.full}">${o.name}</td><td class="r">${o.sg}%</td><td>${o.exp_last_result==='Отрицательное' ? '<span class=\"pill err\">Отрицательное</span>' : (o.exp_last_result||'—')}</td><td>${o.exp_last_date||'—'}</td><td>${o.exp_pending ? '<span class=\"pill warn\">да</span>' : '—'}</td></tr>`
  ).join('');
  const sdGap = DATA.objects.filter(o=>o.exp_last_result==='Положительное' && !o.sd_confirmed);
  document.getElementById('sdGapNote').textContent = sdGap.length
    ? `Сейчас в этом состоянии ${sdGap.length} ${sdGap.length===1?'школа':'школ'}: ${sdGap.map(o=>o.name).join(', ')}.`
    : 'Сейчас все школы с положительным заключением имеют согласованную смету.';
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
    sub = "48 школ, обновлено 14.09.2026"

    method_html = (
        HEAD_STYLE.replace("__TITLE__", "СГ и выплаты — методика")
        .replace("__SUB__", sub + " — методика управления финансированием портфеля")
        .replace("__NAV__", '<a class="navlink" href="dashboard.html">→ Дашборд по объектам</a>')
        .replace("__WRAP__", "880")
        .replace("__CKS_LOGO__", CKS_LOGO_B64)
        + METHOD_BODY
        + METHOD_SCRIPT.replace("__DATA__", data_json)
    ).replace("__ICON_WARN__", ICON_WARN)
    dashboard_html = (
        HEAD_STYLE.replace("__TITLE__", "СГ и выплаты — дашборд")
        .replace("__SUB__", sub + " — объекты, подрядчики, деньги против готовности")
        .replace("__NAV__", '<a class="navlink" href="index.html">→ Методика и выводы</a>')
        .replace("__WRAP__", "1440")
        .replace("__CKS_LOGO__", CKS_LOGO_B64)
        + DASHBOARD_BODY
        + DASHBOARD_SCRIPT.replace("__DATA__", data_json)
    ).replace("__ICON_WARN__", ICON_WARN)

    out = ROOT / "index.html"
    out.write_text(method_html, encoding="utf-8")
    (ROOT / "sg-pay-analysis.html").write_text(method_html, encoding="utf-8")
    dash_out = ROOT / "dashboard.html"
    dash_out.write_text(dashboard_html, encoding="utf-8")
    print(f"OK: {out} ({out.stat().st_size} bytes) + {dash_out} ({dash_out.stat().st_size} bytes), {payload['stats']['n']} schools")


if __name__ == "__main__":
    main()

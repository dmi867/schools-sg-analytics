#!/usr/bin/env python3
"""Generate index.html from Schools xlsx data."""
import json
import math
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent


def short(n, keep_prefix=False):
    n = n or ""
    if not keep_prefix:
        n = n.replace("МБОУ ", "").replace("МАОУ ", "").replace("МОУ ", "")
    return (n[:42] + "…") if len(n) > 42 else n


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


def spearman(xs, ys):
    def ranks(a):
        order = sorted(range(len(a)), key=lambda i: a[i])
        r = [0.0] * len(a)
        i = 0
        while i < len(a):
            j = i
            while j + 1 < len(a) and a[order[j + 1]] == a[order[i]]:
                j += 1
            avg = (i + j) / 2 + 1
            for k in range(i, j + 1):
                r[order[k]] = avg
            i = j + 1
        return r

    return corr(ranks(xs), ranks(ys))


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


def load_simple_list():
    wb = openpyxl.load_workbook(ROOT / "2408_Акцент_Simple List.xlsx", data_only=True)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    col = {h: i for i, h in enumerate(headers)}
    sl = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        uin = row[col["Код УИН"]]
        if not uin:
            continue
        yr = row[col["Год финансирования"]]
        if uin in sl and yr and str(yr).isdigit() and int(yr) <= int(sl[uin].get("yr") or 0):
            continue
        pay_pct = None
        v = row[col["Процент выплат"]]
        if v is not None:
            try:
                pay_pct = round(float(str(v).replace(",", ".")), 1)
            except ValueError:
                pass
        sl[uin] = {
            "yr": yr,
            "name": row[col["Название объекта"]],
            "pay_pct": pay_pct,
            "exp_in": parse_date(row[col["Дата подачи заявления (захода) на экспертизу"]]),
            "exp_start": parse_date(row[col["Дата начала экспертизы"]]),
            "exp_done": parse_date(row[col["Дата получения заключения (завершения ) экспертизы"]]),
            "ctr_plan": parse_date(row[col["Заключение контракта начало план КСГ"]]),
            "ctr_fact": parse_date(row[col["Заключение контракта начало факт КСГ"]]),
        }
    return sl


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

    cross, per, traj, kt_rows, kt_dates = [], [], {}, [], {}

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
        name = info.get("name", uin)
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
                        pays.append({"d": d, "amt": amt})
        pays.sort(key=lambda x: x["d"])
        total = sum(p["amt"] for p in pays)
        last = series[-1]
        pay_pct = info.get("pay_pct")

        pay_scale = pay_pct / total if (pay_pct is not None and total > 0) else None

        cum = 0
        pi = 0
        pts = []
        for pt in series:
            while pi < len(pays) and pays[pi]["d"] <= pt["d"]:
                cum += pays[pi]["amt"]
                pi += 1
            pts.append(
                {
                    "d": pt["d"].strftime("%d.%m.%y"),
                    "sg": round(pt["fact"], 1),
                    "plan": round(pt["plan"], 1) if pt["plan"] is not None else None,
                    "pay": round(cum / 1e6, 1),
                    "payPct": round(cum * pay_scale, 1) if pay_scale is not None else None,
                }
            )
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

        kt_dates[uin] = {
            "exp_plan": fmt_date(exp.get("plan_end")),
            "exp_fact": fmt_date(exp.get("fact_end")),
            "exp_sl": fmt_date(info.get("exp_done")),
            "smr_start": fmt_date(smr.get("fact_start")),
            "ctr_fact": fmt_date(ctr.get("fact_start") or info.get("ctr_fact")),
            "rs_plan": fmt_date(rs.get("plan_end")),
            "rs_fact": fmt_date(rs.get("fact_end")),
        }

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
            }
        )

        cross.append(
            {
                "uin": uin,
                "name": short(name),
                "full": name,
                "sg": round(last["fact"], 1),
                "pct": pay_pct,
                "pay": round(total / 1e6, 1),
                "gap": sg_pay_gap,
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

    valid = [s for s in cross if s["pct"] is not None and s["pay"] > 0]
    facts = [s["sg"] for s in valid]
    pcts = [s["pct"] for s in valid]
    a, b, r2 = linreg(pcts, facts)
    varying = [p for p in per if p["vary"] and p["r"] is not None]
    varying.sort(key=lambda x: -(x["r"] or -1))
    rs_corr = [p["r"] for p in varying]

    bins = []
    for lo, hi in [(0, 40), (40, 60), (60, 80), (80, 101)]:
        grp = [s["sg"] for s in cross if s["pct"] is not None and lo <= s["pct"] < hi]
        if grp:
            bins.append(
                {
                    "label": f"{lo}–{hi}%",
                    "n": len(grp),
                    "mean_sg": round(sum(grp) / len(grp), 1),
                }
            )

    gaps = [s["gap"] for s in cross if s["gap"] is not None]
    gaps_sorted = sorted(
        [s for s in cross if s["gap"] is not None],
        key=lambda x: -x["gap"],
    )
    kt_sorted = sorted(kt_rows, key=lambda x: (-x["flag_n"], -(x["sg_pay_gap"] or 0)))

    has_rs_kt = any(kt_dates[u].get("rs_plan") or kt_dates[u].get("rs_fact") for u in kt_dates)

    return {
        "stats": {
            "n": len(valid),
            "pearson_sg_pay_pct": round(corr(facts, pcts), 3),
            "spearman_sg_pay_pct": round(spearman(facts, pcts), 3),
            "median_r": round(median(rs_corr), 3) if rs_corr else 0,
            "n_varying": len(varying),
            "n_sg_ahead": sum(1 for g in gaps if g > 55),
            "median_gap": round(median(gaps), 1) if gaps else 0,
            "n_sg_behind_plan": sum(
                1 for p in per if p.get("sg_plan_gap") is not None and p["sg_plan_gap"] < -5
            ),
            "n_smr_before_exp": sum(1 for p in per if "стройка до экспертизы" in p.get("flags", [])),
            "n_kt_issues": sum(1 for p in per if any(f in p.get("flags", []) for f in ("стройка до экспертизы", "экспертиза не бьётся с КСГ", "РС опоздал", "нет заключения экспертизы"))),
            "has_rs_kt": has_rs_kt,
        },
        "attention": [
            {
                "name": s["name"],
                "full": s["full"],
                "sg": s["sg"],
                "pct": s["pct"],
                "gap": s["gap"],
                "flags": next((k["flags"] for k in kt_rows if k["name"] == s["name"]), []),
            }
            for s in gaps_sorted[:8]
        ],
        "kt_attention": kt_sorted[:10],
        "reg": {"a": round(a, 1), "b": round(b, 4), "r2": round(r2, 3)},
        "bins": bins,
        "cross": cross,
        "per": sorted(per, key=lambda x: -(x["r"] if x["r"] is not None else -1)),
        "traj": traj,
        "kt_dates": kt_dates,
        "top12": varying[:12],
        "defaultUin": varying[0]["uin"] if varying else per[0]["uin"],
    }


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>СГ и выплаты — 47 школ</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<style>
  :root { --bg:#f4f3ef; --surface:#fff; --text:#1c1c1c; --muted:#555; --faint:#888;
    --line:#ddd9d0; --accent:#1a4d7a; --green:#2d6a4f;
    --warn-bg:#faf3e6; --warn-border:#c4922a; --info-bg:#edf3f9; --ok-bg:#eaf4ee; }
  * { box-sizing:border-box }
  body { margin:0; font:15px/1.5 "Segoe UI",system-ui,sans-serif; background:var(--bg); color:var(--text) }
  .wrap { max-width:880px; margin:0 auto; padding:24px 18px 56px }
  h1 { font-size:1.5rem; font-weight:650; margin:0 0 4px }
  .sub { color:var(--muted); margin:0 0 20px; font-size:.92rem }
  ul.brief { margin:0; padding-left:1.2rem; color:var(--muted) }
  ul.brief li { margin:6px 0 }
  .kpis { display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin:16px 0 }
  .kpi { background:var(--surface); border:1px solid var(--line); border-radius:6px; padding:12px 14px }
  .kpi .n { font-size:1.35rem; font-weight:700; color:var(--accent); line-height:1.2 }
  .kpi .l { font-size:.75rem; color:var(--muted); margin-top:4px }
  .chart-sm { position:relative; height:220px; margin-top:8px }
  .chart { position:relative; height:280px; margin-top:10px }
  .chart.tall { height:340px }
  .note { font-size:.82rem; color:var(--faint); margin:6px 0 0 }
  .box { background:var(--surface); border:1px solid var(--line); border-radius:6px; padding:14px 16px; margin:12px 0 }
  .mini table { width:100%; font-size:.84rem; border-collapse:collapse }
  .mini th,.mini td { padding:6px 8px; border-bottom:1px solid var(--line); text-align:left }
  .mini th { color:var(--muted); font-weight:600 }
  .mini td.r,.mini th.r { text-align:right; font-variant-numeric:tabular-nums }
  details { background:var(--surface); border:1px solid var(--line); border-radius:6px; margin:10px 0; overflow:hidden }
  details > summary { cursor:pointer; padding:14px 16px; font-weight:600; list-style:none; user-select:none }
  details > summary::-webkit-details-marker { display:none }
  details > summary::after { content:'+'; float:right; color:var(--faint); font-weight:400 }
  details[open] > summary::after { content:'−' }
  details > summary:hover { background:#faf9f6 }
  .detail-body { padding:0 16px 16px; border-top:1px solid var(--line) }
  .cols2 { display:grid; grid-template-columns:1fr 1fr; gap:14px }
  select { font:inherit; padding:6px 10px; border:1px solid var(--line); border-radius:4px; background:#fff; max-width:100% }
  .row { display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin:10px 0 }
  .tag { font-size:.75rem; padding:3px 9px; background:var(--info-bg); border:1px solid var(--line); border-radius:99px; color:var(--accent) }
  table.full { width:100%; border-collapse:collapse; font-size:.84rem }
  table.full th,table.full td { padding:7px 9px; border-bottom:1px solid var(--line) }
  table.full th { background:#eeebe4; position:sticky; top:0; text-align:left }
  table.full td.r,table.full th.r { text-align:right; font-variant-numeric:tabular-nums }
  .tbl-wrap { max-height:420px; overflow:auto; border:1px solid var(--line); border-radius:4px; margin-top:10px }
  .flag { font-size:.72rem; padding:2px 7px; margin:1px 2px 1px 0; display:inline-block; background:#faf3e6; border:1px solid #e0c88a; border-radius:4px; color:#7a5a12 }
  .flag.warn { background:#fdecea; border-color:#e8a8a0; color:#8a2a22 }
  .kpis.four { grid-template-columns:repeat(4,1fr) }
  @media(max-width:900px) { .kpis.four { grid-template-columns:repeat(2,1fr) } }
  @media(max-width:700px) { .kpis,.kpis.four,.cols2 { grid-template-columns:1fr } }
</style>
</head>
<body>
<div class="wrap">
  <h1>СГ и выплаты</h1>
  <p class="sub">47 школ, август 2026</p>

  <ul class="brief">
    <li>У <span id="b1"></span> школ готовность ниже плана больше чем на 5 п.п.</li>
    <li>У <span id="b2"></span> школ готовность заметно выше того, что уже заплатили.</li>
    <li>У <span id="b3"></span> школ стройка началась раньше, чем в КСГ стоит экспертиза.</li>
    <li>Про РС: в КСГ по школам такой точки нет, смотрим контракт и экспертизу.</li>
  </ul>

  <div class="kpis four">
    <div class="kpi"><div class="n" id="k1"></div><div class="l">готовность выше выплат</div></div>
    <div class="kpi"><div class="n" id="k2"></div><div class="l">обычный разрыв, п.п.</div></div>
    <div class="kpi"><div class="n" id="k4"></div><div class="l">факт ниже плана</div></div>
    <div class="kpi"><div class="n" id="k5"></div><div class="l">стройка до экспертизы</div></div>
  </div>

  <div class="box mini">
    <strong style="display:block;margin-bottom:8px">Кого смотреть первым</strong>
    <table>
      <thead><tr><th>Школа</th><th class="r">СГ</th><th class="r">План</th><th class="r">Выпл.</th><th>Что не так</th></tr></thead>
      <tbody id="ktAttn"></tbody>
    </table>
  </div>

  <div class="box">
    <strong style="display:block;margin-bottom:8px">Готовность и выплаты по школам</strong>
    <div class="chart-sm"><canvas id="cMini"></canvas></div>
    <p class="note">Связь слабая: <span id="k3"></span> (чем ближе к 0, тем слабее).</p>
  </div>

  <details id="secCharts">
    <summary>Графики по школе</summary>
    <div class="detail-body">
      <div class="cols2" style="margin-top:14px">
        <div>
          <strong>Готовность и выплаты</strong>
          <div class="chart"><canvas id="cScatter"></canvas></div>
        </div>
        <div>
          <strong>Средняя СГ в группах по выплатам</strong>
          <div class="chart"><canvas id="cBins"></canvas></div>
        </div>
      </div>
      <strong>Одна школа</strong>
      <div class="row">
        <select id="selSchool"></select>
        <span class="tag" id="tagR"></span>
        <span class="note" id="metaSchool" style="margin:0"></span>
      </div>
      <div class="chart"><canvas id="cTraj"></canvas></div>
      <p class="note" id="ktLine"></p>
      <p class="note">Синяя — факт, пунктир — план, зелёная — % от суммы контракта, уже выплаченной на эту дату.</p>
    </div>
  </details>

  <details id="secKt">
    <summary>Все школы: даты и разрывы</summary>
    <div class="detail-body">
      <div class="tbl-wrap">
        <table class="full">
          <thead><tr>
            <th>Школа</th><th class="r">Факт</th><th class="r">План</th><th class="r">Δ</th>
            <th class="r">Выпл.</th><th class="r">Δ</th>
            <th>Экспертиза</th><th>Старт СМР</th><th>Контракт</th><th>Замечания</th>
          </tr></thead>
          <tbody id="ktTbl"></tbody>
        </table>
      </div>
      <p class="note" id="rsNote"></p>
    </div>
  </details>

  <details id="secTable">
    <summary>Краткая таблица</summary>
    <div class="detail-body">
      <div class="tbl-wrap">
        <table class="full">
          <thead><tr>
            <th>Школа</th><th class="r">СГ</th><th class="r">План</th><th class="r">Выпл.</th>
            <th class="r">Разрыв</th><th class="r">Сходятся</th>
          </tr></thead>
          <tbody id="tbl"></tbody>
        </table>
      </div>
    </div>
  </details>

  <details id="secContext">
    <summary>Если коротко</summary>
    <div class="detail-body">
      <ul class="brief" style="margin-top:14px">
        <li>Готовность и деньги считаются отдельно — высокая СГ при 30% выплат это нормально.</li>
        <li>Деньги идут после приёмки, не по цифре с мониторинга.</li>
        <li>«Стройка до экспертизы» — в КСГ экспертиза позже, а СМР уже идут.</li>
        <li>РС в этой выгрузке по школам не выделен.</li>
      </ul>
      <p class="note">Данные: файлы СГ, платежи, Simple List, КСГ+Экспертиза.</p>
    </div>
  </details>
</div>
<script>
const DATA = __DATA__;
const blue='#1a4d7a', blueL='rgba(26,77,122,.35)', green='#2d6a4f';
const charts = { mini:false, scatter:false, bins:false, traj:false };

function flagsHtml(arr) {
  if(!arr||!arr.length) return '—';
  return arr.map(f=>'<span class="flag'+(f.includes('стройка')||f.includes('эксперт')?' warn':'')+'">'+f+'</span>').join('');
}

document.getElementById('k1').textContent = DATA.stats.n_sg_ahead + ' из ' + DATA.stats.n;
document.getElementById('k2').textContent = '+' + DATA.stats.median_gap;
document.getElementById('k3').textContent = DATA.stats.pearson_sg_pay_pct.toFixed(2);
document.getElementById('k4').textContent = DATA.stats.n_sg_behind_plan;
document.getElementById('k5').textContent = DATA.stats.n_smr_before_exp;

document.getElementById('ktAttn').innerHTML = DATA.kt_attention.map(s =>
  `<tr><td title="${s.full}">${s.name}</td><td class="r">${s.sg}%</td><td class="r">${s.plan??'—'}%</td><td class="r">${s.pct??'—'}%</td><td>${flagsHtml(s.flags)}</td></tr>`
).join('');

document.getElementById('b1').textContent = DATA.stats.n_sg_behind_plan;
document.getElementById('b2').textContent = DATA.stats.n_sg_ahead;
document.getElementById('b3').textContent = DATA.stats.n_smr_before_exp;

document.getElementById('ktTbl').innerHTML = [...DATA.per].sort((a,b)=>(b.flags?.length||0)-(a.flags?.length||0)).map(p=>{
  const k = DATA.kt_dates[p.uin]||{};
  const exp = [k.exp_sl,k.exp_plan,k.exp_fact].filter(Boolean).join(' / ') || '—';
  return `<tr><td>${p.name}</td><td class="r">${p.sg}</td><td class="r">${p.plan??'—'}</td><td class="r">${p.sg_plan_gap??'—'}</td><td class="r">${p.pct??'—'}</td><td class="r">${p.sg!=null&&p.pct!=null?(p.sg-p.pct).toFixed(0):'—'}</td><td>${exp}</td><td>${k.smr_start||'—'}</td><td>${k.ctr_fact||'—'}</td><td>${flagsHtml(p.flags)}</td></tr>`;
}).join('');

document.getElementById('rsNote').textContent = DATA.stats.has_rs_kt
  ? 'В КСГ есть «Получение РС».'
  : '«Получение РС» в КСГ по школам нет — смотрим контракт и экспертизу.';

document.getElementById('tbl').innerHTML = DATA.per.map(p => {
  const gap = p.pct!=null ? (p.sg-p.pct).toFixed(0) : '—';
  return `<tr><td>${p.name}</td><td class="r">${p.sg}</td><td class="r">${p.plan??'—'}</td><td class="r">${p.pct??'—'}</td><td class="r">${gap}</td><td class="r">${p.r!=null?p.r.toFixed(2):'—'}</td></tr>`;
}).join('');

const sel = document.getElementById('selSchool');
DATA.per.filter(p=>p.vary).forEach(p=>{
  const o=document.createElement('option'); o.value=p.uin;
  o.textContent=p.name + (p.flags?.length ? ' *' : ''); sel.appendChild(o);
});
sel.value = DATA.defaultUin;
let chartTraj;

function drawTraj(uin) {
  const rows = DATA.traj[uin]||[], m = DATA.per.find(p=>p.uin===uin);
  const k = DATA.kt_dates[uin]||{};
  document.getElementById('tagR').textContent = m&&m.r!=null ? 'корр. '+m.r.toFixed(2) : '';
  document.getElementById('metaSchool').textContent = m ? `${m.name}: ${m.sg}% факт, ${m.plan??'—'}% план, ${m.pct??'—'}% выпл.` : '';
  document.getElementById('ktLine').textContent = `Экспертиза ${k.exp_sl||k.exp_plan||'—'}, СМР с ${k.smr_start||'—'}, контракт ${k.ctr_fact||'—'}` + (m&&m.flags?.length ? '. '+m.flags.join(', ') : '');
  const cfg = { type:'line', data:{ labels:rows.map(r=>r.d), datasets:[
    { label:'Факт', data:rows.map(r=>r.sg), borderColor:blue, tension:.25, pointRadius:2 },
    { label:'План', data:rows.map(r=>r.plan), borderColor:blue, borderDash:[5,4], tension:.25, pointRadius:0 },
    { label:'Выплаты', data:rows.map(r=>r.payPct), borderColor:green, tension:.25, pointRadius:2 }
  ]}, options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom'}}, scales:{ y:{min:0,max:100} } } };
  if(chartTraj) chartTraj.destroy(); chartTraj = new Chart(document.getElementById('cTraj'), cfg);
}
sel.onchange = e => drawTraj(e.target.value);

function initMini() {
  if(charts.mini) return; charts.mini = true;
  const pts = DATA.cross.filter(s=>s.pct!=null);
  new Chart(document.getElementById('cMini'), {
    type:'scatter',
    data:{ datasets:[{ data:pts.map(s=>({x:s.pct,y:s.sg})), backgroundColor:'rgba(26,77,122,.7)', pointRadius:4 }] },
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}, tooltip:{callbacks:{label:c=>`выплаты ${c.raw.x}%, СГ ${c.raw.y}%`}}},
      scales:{ x:{title:{display:true,text:'Выплаты, %'},min:20,max:100,ticks:{maxTicksLimit:5}}, y:{title:{display:true,text:'СГ, %'},min:70,max:100,ticks:{maxTicksLimit:5}} } }
  });
}

function initDetailCharts() {
  if(charts.scatter) return;
  charts.scatter = charts.bins = true;
  const pts = DATA.cross.filter(s=>s.pct!=null);
  const line=[]; for(let x=25;x<=95;x+=5) line.push({x,y:DATA.reg.a+DATA.reg.b*x});
  new Chart(document.getElementById('cScatter'), {
    type:'scatter',
    data:{ datasets:[
      { label:'Школы', data:pts.map(s=>({x:s.pct,y:s.sg})), backgroundColor:'rgba(26,77,122,.75)', pointRadius:4 },
      { label:'Тренд', data:line, type:'line', borderColor:blue, borderDash:[5,4], pointRadius:0 }
    ]},
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom'}},
      scales:{ x:{title:{display:true,text:'Выплаты, %'},min:20,max:100}, y:{title:{display:true,text:'СГ, %'},min:70,max:100} } }
  });
  new Chart(document.getElementById('cBins'), {
    type:'bar',
    data:{ labels:DATA.bins.map(b=>b.label), datasets:[{ data:DATA.bins.map(b=>b.mean_sg), backgroundColor:blue }] },
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}}, scales:{ y:{min:88,max:100} } }
  });
  drawTraj(sel.value);
  charts.traj = true;
}

initMini();
document.getElementById('secCharts').addEventListener('toggle', e => { if(e.target.open) initDetailCharts(); });
</script>
</body>
</html>
"""


def main():
    payload = load_data()
    html = HTML_TEMPLATE.replace("__DATA__", json.dumps(payload, ensure_ascii=False))
    out = ROOT / "index.html"
    out.write_text(html, encoding="utf-8")
    (ROOT / "sg-pay-analysis.html").write_text(html, encoding="utf-8")
    print(f"OK: {out} ({out.stat().st_size} bytes, {payload['stats']['n']} schools)")


if __name__ == "__main__":
    main()

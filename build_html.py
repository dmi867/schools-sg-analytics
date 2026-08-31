#!/usr/bin/env python3
"""Generate index.html from Schools xlsx data."""
import json
import math
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent


def short(n):
    n = (n or "").replace("МБОУ ", "").replace("МАОУ ", "").replace("МОУ ", "")
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


def load_data():
    wb_sl = openpyxl.load_workbook(ROOT / "2408_Акцент_Simple List.xlsx", data_only=True)
    ws = wb_sl.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    col = {h: i for i, h in enumerate(headers)}
    sl = {}
    pay_pct_by_uin = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        uin = row[col["Код УИН"]]
        if uin and uin not in sl:
            sl[uin] = row[col["Название объекта"]]
            v = row[col["Процент выплат"]]
            if v is not None:
                try:
                    pay_pct_by_uin[uin] = round(float(str(v).replace(",", ".")), 1)
                except ValueError:
                    pass

    cross, per, traj = [], [], {}

    for f in sorted(ROOT.glob("*.xlsx")):
        if "_платежи" in f.name or "Акцент" in f.name or "Simple" in f.name:
            continue
        uin = f.stem
        name = sl.get(uin, uin)
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
                    "pay": round(cum / 1e6, 1),
                }
            )
        if len(pts) > 36:
            step = math.ceil(len(pts) / 36)
            pts = pts[::step][:-1] + [pts[-1]]
        traj[uin] = pts

        pay_pct = pay_pct_by_uin.get(uin)
        pays_t = [p["pay"] for p in pts]
        vary = max(pays_t) - min(pays_t) >= 0.01
        r = None
        if vary and len(pts) >= 5:
            r = corr([p["sg"] for p in pts], pays_t)

        cross.append(
            {
                "uin": uin,
                "name": short(name),
                "full": name,
                "sg": round(last["fact"], 1),
                "pct": pay_pct,
                "pay": round(total / 1e6, 1),
                "gap": round(last["fact"] - pay_pct, 1) if pay_pct is not None else None,
            }
        )
        per.append(
            {
                "uin": uin,
                "name": short(name),
                "sg": round(last["fact"], 1),
                "pct": pay_pct,
                "pay": round(total / 1e6, 1),
                "r": round(r, 3) if r is not None else None,
                "vary": vary,
                "n": len(series),
            }
        )

    valid = [s for s in cross if s["pct"] is not None and s["pay"] > 0]
    facts = [s["sg"] for s in valid]
    pcts = [s["pct"] for s in valid]
    a, b, r2 = linreg(pcts, facts)
    varying = [p for p in per if p["vary"] and p["r"] is not None]
    varying.sort(key=lambda x: -(x["r"] or -1))
    rs = [p["r"] for p in varying]

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

    return {
        "stats": {
            "n": len(valid),
            "pearson_sg_pay_pct": round(corr(facts, pcts), 3),
            "spearman_sg_pay_pct": round(spearman(facts, pcts), 3),
            "median_r": round(sorted(rs)[len(rs) // 2], 3) if rs else 0,
            "n_varying": len(varying),
            "n_sg_ahead": sum(1 for g in gaps if g > 15),
            "median_gap": round(sorted(gaps)[len(gaps) // 2], 1) if gaps else 0,
        },
        "attention": [
            {
                "name": s["name"],
                "full": s["full"],
                "sg": s["sg"],
                "pct": s["pct"],
                "gap": s["gap"],
            }
            for s in gaps_sorted[:8]
        ],
        "reg": {"a": round(a, 1), "b": round(b, 4), "r2": round(r2, 3)},
        "bins": bins,
        "cross": cross,
        "per": sorted(per, key=lambda x: -(x["r"] if x["r"] is not None else -1)),
        "traj": traj,
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
  @media(max-width:700px) { .kpis,.cols2 { grid-template-columns:1fr } }
</style>
</head>
<body>
<div class="wrap">
  <h1>Стройготовность и выплаты</h1>
  <p class="sub">47 школ · август 2026</p>

  <ul class="brief">
    <li><strong>Между школами</strong> — связи почти нет: больше выплат ≠ выше готовность.</li>
    <li><strong>Внутри школы</strong> — графики часто идут рядом (стройка и оплата по ходу).</li>
    <li><strong>Платят не по СГ</strong>, а после приёмки работ и документов.</li>
  </ul>

  <div class="kpis">
    <div class="kpi"><div class="n" id="k1"></div><div class="l">школ, где готовность сильно выше выплат</div></div>
    <div class="kpi"><div class="n" id="k2"></div><div class="l">типичный разрыв (СГ − % выплат), п.п.</div></div>
    <div class="kpi"><div class="n" id="k3"></div><div class="l">связь между школами (0–1)</div></div>
  </div>

  <div class="box">
    <strong style="display:block;margin-bottom:8px">Общая картина</strong>
    <div class="chart-sm"><canvas id="cMini"></canvas></div>
    <p class="note">Каждая точка — школа. Точки разбросаны → прямой зависимости нет.</p>
  </div>

  <div class="box mini">
    <strong style="display:block;margin-bottom:8px">Кого смотреть в первую очередь</strong>
    <p class="note" style="margin-top:0">Наибольший разрыв: готовность есть, выплат мало.</p>
    <table>
      <thead><tr><th>Школа</th><th class="r">СГ</th><th class="r">Выплаты</th><th class="r">Разрыв</th></tr></thead>
      <tbody id="attn"></tbody>
    </table>
  </div>

  <details id="secCharts">
    <summary>Графики и динамика по школам</summary>
    <div class="detail-body">
      <div class="cols2" style="margin-top:14px">
        <div>
          <strong>Готовность vs выплаты</strong>
          <div class="chart"><canvas id="cScatter"></canvas></div>
        </div>
        <div>
          <strong>Средняя СГ по группам выплат</strong>
          <div class="chart"><canvas id="cBins"></canvas></div>
        </div>
      </div>
      <strong>Школы с наиболее синхронными графиками</strong>
      <div class="chart tall"><canvas id="cTop"></canvas></div>
      <strong>Одна школа</strong>
      <div class="row">
        <select id="selSchool"></select>
        <span class="tag" id="tagR"></span>
        <span class="note" id="metaSchool" style="margin:0"></span>
      </div>
      <div class="chart"><canvas id="cTraj"></canvas></div>
      <p class="note">Доп. цифры: связь по рейтингу <span id="d2"></span>, внутри школы <span id="d3"></span>, выплаты объясняют <span id="d4"></span> разницы в СГ.</p>
    </div>
  </details>

  <details id="secTable">
    <summary>Таблица всех школ</summary>
    <div class="detail-body">
      <div class="tbl-wrap">
        <table class="full">
          <thead><tr>
            <th>Школа</th><th class="r">СГ %</th><th class="r">% выплат</th>
            <th class="r">Выплаты, млн</th><th class="r">Разрыв</th><th class="r">Вместе</th>
          </tr></thead>
          <tbody id="tbl"></tbody>
        </table>
      </div>
    </div>
  </details>

  <details id="secContext">
    <summary>Почему так и что делать</summary>
    <div class="detail-body">
      <ul class="brief" style="margin-top:14px">
        <li>Готовность на мониторинге и деньги — разные контуры.</li>
        <li>Цепочка: работы → документы приняли → акт → оплата.</li>
        <li>Цель — не «догнать выплатами», а быстрее принимать документы.</li>
        <li>Высокая СГ при низких выплатах — сигнал по приёмке, не по кассе.</li>
      </ul>
      <p class="note">Источник: файлы СГ и платежей, названия — Simple List.</p>
    </div>
  </details>
</div>
<script>
const DATA = __DATA__;
const blue='#1a4d7a', green='#2d6a4f';
const charts = { mini:false, scatter:false, bins:false, top:false, traj:false };

document.getElementById('k1').textContent = DATA.stats.n_sg_ahead + ' из ' + DATA.stats.n;
document.getElementById('k2').textContent = '+' + DATA.stats.median_gap;
document.getElementById('k3').textContent = DATA.stats.pearson_sg_pay_pct.toFixed(2);
document.getElementById('d2').textContent = DATA.stats.spearman_sg_pay_pct.toFixed(2);
document.getElementById('d3').textContent = DATA.stats.median_r.toFixed(2);
document.getElementById('d4').textContent = (DATA.reg.r2*100).toFixed(0)+'%';

document.getElementById('attn').innerHTML = DATA.attention.map(s =>
  `<tr><td title="${s.full}">${s.name}</td><td class="r">${s.sg}%</td><td class="r">${s.pct}%</td><td class="r">+${s.gap}</td></tr>`
).join('');

document.getElementById('tbl').innerHTML = DATA.per.map(p => {
  const gap = p.pct!=null ? (p.sg-p.pct).toFixed(0) : '—';
  return `<tr><td>${p.name}</td><td class="r">${p.sg}</td><td class="r">${p.pct??'—'}</td><td class="r">${p.pay}</td><td class="r">${gap}</td><td class="r">${p.r!=null?p.r.toFixed(2):'—'}</td></tr>`;
}).join('');

const sel = document.getElementById('selSchool');
DATA.per.filter(p=>p.vary&&p.r!=null).forEach(p=>{
  const o=document.createElement('option'); o.value=p.uin;
  o.textContent=`${p.name} (${p.r.toFixed(2)})`; sel.appendChild(o);
});
sel.value = DATA.defaultUin;
let chartTraj;

function drawTraj(uin) {
  const rows = DATA.traj[uin]||[], m = DATA.per.find(p=>p.uin===uin), maxP = Math.max(...rows.map(r=>r.pay),1);
  document.getElementById('tagR').textContent = m&&m.r!=null ? 'вместе '+m.r.toFixed(2) : '';
  document.getElementById('metaSchool').textContent = m ? `${m.name} · ${m.sg}% / ${m.pct}% · ${m.pay} млн` : '';
  const cfg = { type:'line', data:{ labels:rows.map(r=>r.d), datasets:[
    { label:'СГ', data:rows.map(r=>r.sg), borderColor:blue, tension:.25, pointRadius:2 },
    { label:'Выплаты', data:rows.map(r=>Math.round(r.pay/maxP*100)), borderColor:green, tension:.25, pointRadius:2 }
  ]}, options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom'}}, scales:{ y:{min:0,max:105} } } };
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
      scales:{ x:{title:{display:true,text:'% выплат'},min:20,max:100,ticks:{maxTicksLimit:5}}, y:{title:{display:true,text:'СГ %'},min:70,max:105,ticks:{maxTicksLimit:5}} } }
  });
}

function initDetailCharts() {
  if(charts.scatter) return;
  charts.scatter = charts.bins = charts.top = true;
  const pts = DATA.cross.filter(s=>s.pct!=null);
  const line=[]; for(let x=25;x<=95;x+=5) line.push({x,y:DATA.reg.a+DATA.reg.b*x});
  new Chart(document.getElementById('cScatter'), {
    type:'scatter',
    data:{ datasets:[
      { label:'Школы', data:pts.map(s=>({x:s.pct,y:s.sg,name:s.full})), backgroundColor:'rgba(26,77,122,.75)', pointRadius:4 },
      { label:'Тренд', data:line, type:'line', borderColor:blue, borderDash:[5,4], pointRadius:0 }
    ]},
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'bottom'}},
      scales:{ x:{title:{display:true,text:'% выплат'},min:20,max:100}, y:{title:{display:true,text:'СГ %'},min:70,max:105} } }
  });
  new Chart(document.getElementById('cBins'), {
    type:'bar',
    data:{ labels:DATA.bins.map(b=>b.label), datasets:[{ data:DATA.bins.map(b=>b.mean_sg), backgroundColor:blue }] },
    options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}}, scales:{ y:{min:88,max:100} } }
  });
  new Chart(document.getElementById('cTop'), {
    type:'bar',
    data:{ labels:DATA.top12.map(p=>p.name), datasets:[{ data:DATA.top12.map(p=>p.r), backgroundColor:green }] },
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false, plugins:{legend:{display:false}}, scales:{ x:{min:0,max:1} } }
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

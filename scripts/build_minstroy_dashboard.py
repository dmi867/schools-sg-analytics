#!/usr/bin/env python3
"""Дашборд Минстроя: UI как у МинЖКХ (Exon glass-nav) на data/minstroy_dynamics.json.

Выход: dashboard.html

  python3 scripts/build_minstroy_dynamics.py   # обновить JSON
  python3 scripts/build_minstroy_dashboard.py  # собрать HTML
"""
from __future__ import annotations

import base64
import json
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
DATA = ROOT / "data" / "minstroy_dynamics.json"
OUT_HTML = ROOT / "dashboard.html"
CKS_LOGO_B64 = base64.b64encode((ROOT / "assets" / "cks-logo.png").read_bytes()).decode("ascii")
MO_LOGO_B64 = base64.b64encode((ROOT / "assets" / "logo-mo.png").read_bytes()).decode("ascii")


# Дизайн 1:1 с Exon_2 Header + Index (glass-nav, Inter, soft cards, select+search).
EXON_HEAD = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Минстрой — СГ и выплаты</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-annotation@3.0.1/dist/chartjs-plugin-annotation.min.js"></script>
<style>
  :root {
    --background: 210 40% 98%;
    --foreground: 222 47% 11%;
    --muted: 210 30% 96%;
    --muted-foreground: 215 16% 47%;
    --border: 214 32% 88%;
    --primary: 207 90% 35%;
    --card: 0 0% 100%;
    --teal: #1f8682;
    --ok: #16a34a; --err: #dc2626; --warn: #d97706;
    --ok-bg: #ecfdf5; --err-bg: #fef2f2; --warn-bg: #fffbeb;
    --radius: 0.5rem;
  }
  * { box-sizing: border-box }
  body {
    margin: 0; min-height: 100vh;
    font-family: Inter, system-ui, sans-serif;
    font-size: 14px; line-height: 1.5;
    color: hsl(var(--foreground));
    background: hsl(var(--background));
  }
  /* ---- Header как в Exon ---- */
  .glass-nav {
    background: rgba(255,255,255,.85);
    backdrop-filter: blur(16px) saturate(180%);
    -webkit-backdrop-filter: blur(16px) saturate(180%);
    border-bottom: 1px solid hsl(var(--border) / .4);
    position: sticky; top: 0; z-index: 50;
  }
  .container {
    width: 100%; max-width: 1600px; margin: 0 auto; padding: 0 1.5rem;
  }
  .hdr {
    display: flex; align-items: center; justify-content: space-between;
    min-height: 60px; gap: .75rem;
  }
  .brand { display: flex; align-items: center; gap: .75rem; flex-shrink: 0 }
  .brand img.cks { height: 36px; width: auto }
  .brand img.mo { height: 44px; width: auto }
  .brand-meta {
    display: none; padding-left: .5rem;
    border-left: 1px solid hsl(var(--border) / .5); white-space: nowrap;
  }
  @media (min-width: 768px) { .brand-meta { display: block } }
  .brand-meta .t {
    margin: 0; font-size: 1rem; font-weight: 700; line-height: 1.25;
    background: linear-gradient(135deg, hsl(200,60%,23%), hsl(168,50%,38%));
    -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent;
  }
  .brand-meta .s { margin: 0; font-size: .75rem; color: hsl(var(--muted-foreground)); line-height: 1.25 }
  .nav-center { flex: 1; min-width: 0; display: flex; justify-content: center; padding: .375rem 0 }
  .nav-pill {
    display: inline-flex; align-items: center; gap: 1px; padding: 4px; border-radius: 12px;
    background: linear-gradient(135deg, hsl(200 60% 23% / .08), hsl(168 50% 38% / .06));
  }
  .nav-tab {
    position: relative; display: inline-flex; align-items: center; gap: .375rem;
    padding: .5rem 1rem; border: none; border-radius: .5rem; background: transparent;
    font: 500 14px Inter, sans-serif; color: hsl(var(--muted-foreground));
    cursor: pointer; white-space: nowrap; transition: all .3s;
  }
  .nav-tab.active {
    background: #fff; color: hsl(var(--primary));
    box-shadow: 0 4px 6px -1px rgb(0 0 0 / .1), 0 2px 4px -2px rgb(0 0 0 / .1);
  }
  .nav-tab.active::after {
    content: ''; position: absolute; left: 50%; bottom: -4px;
    transform: translateX(-50%); width: 20px; height: 2px;
    border-radius: 99px; background: var(--teal);
  }
  .nav-tab svg { width: 18px; height: 18px; flex-shrink: 0 }
  .nav-tab.active svg { color: var(--teal) }
  .btn-icon {
    display: inline-flex; align-items: center; justify-content: center;
    height: 32px; width: 32px; border-radius: .5rem; flex-shrink: 0;
    border: 1px solid hsl(var(--border)); background: #fff; color: hsl(var(--foreground));
    text-decoration: none; transition: background .15s;
  }
  .btn-icon:hover { background: hsl(var(--muted)) }
  .btn-icon svg { width: 16px; height: 16px }
  /* ---- Main ---- */
  main.container { padding-top: 2rem; padding-bottom: 2rem }
  .stack { display: flex; flex-direction: column; gap: 1.5rem }
  .tabpanel[hidden] { display: none !important }
  .note { margin: 0; font-size: .8125rem; color: hsl(var(--muted-foreground)) }
  /* KPI как карточки Exon StageCard */
  .kpis { display: grid; grid-template-columns: repeat(2, 1fr); gap: .75rem }
  @media (min-width: 640px) { .kpis { grid-template-columns: repeat(4, 1fr) } }
  .kpi {
    background: #fff; border: 2px solid #d1d5db; border-radius: .5rem;
    padding: .75rem; text-align: center;
    transition: transform .2s, box-shadow .2s, border-color .2s, background .2s;
  }
  .kpi:hover { transform: scale(1.02); box-shadow: 0 4px 6px -1px rgb(0 0 0 / .1) }
  .kpi .l { margin: 0 0 .25rem; font-size: .875rem; font-weight: 500; color: hsl(var(--muted-foreground)) }
  .kpi .n { margin: 0; font-size: 1.5rem; font-weight: 700; font-variant-numeric: tabular-nums; color: hsl(var(--foreground)) }
  .kpi .n.ok { color: var(--ok) }
  .kpi .n.warn { color: var(--teal) }
  .kpi .n.err { color: var(--err) }
  .kpi.money { text-align: left }
  .kpi.money .pair { display: flex; justify-content: space-between; gap: .75rem; margin: .35rem 0 0; font-size: .8125rem; color: hsl(var(--muted-foreground)) }
  .kpi.money .pair b { color: hsl(var(--foreground)); font-variant-numeric: tabular-nums }
  /* Фильтры: multi-select + поиск */
  .filters { display: grid; grid-template-columns: 1fr; gap: .75rem }
  @media (min-width: 900px) { .filters { grid-template-columns: 1fr 1.4fr 1fr 1fr 1.2fr } }
  .field {
    display: flex; flex-direction: column; gap: .25rem;
    background: #fff; border: 1px solid hsl(var(--border)); border-radius: .5rem;
    padding: .4rem .65rem .5rem; min-height: 40px;
  }
  .field > span {
    font-size: .65rem; font-weight: 600; letter-spacing: .04em; text-transform: uppercase;
    color: hsl(var(--muted-foreground));
  }
  .field.select-row {
    flex-direction: row; align-items: center; gap: .5rem; padding: 0 .75rem; height: 40px; min-height: 40px;
  }
  .field select, .field input {
    flex: 1; min-width: 0; border: none; outline: none; background: transparent;
    font: inherit; color: hsl(var(--foreground));
  }
  .field select[multiple] {
    height: 88px; width: 100%; padding: 0;
  }
  .field select[multiple] option { padding: .15rem .25rem }
  .field input { height: 100% }
  .field svg { width: 16px; height: 16px; color: hsl(var(--muted-foreground)); flex-shrink: 0 }
  /* Card panels */
  .card {
    background: #fff; border: 1px solid hsl(var(--border)); border-radius: .5rem;
    box-shadow: 0 1px 2px rgb(0 0 0 / .04); overflow: hidden;
  }
  .card-hd {
    display: flex; align-items: flex-start; justify-content: space-between; gap: .75rem;
    padding: 1rem 1.25rem; border-bottom: 1px solid hsl(var(--border));
  }
  .card-hd h2 { margin: 0; font-size: 1rem; font-weight: 600 }
  .card-hd p { margin: .15rem 0 0; font-size: .75rem; color: hsl(var(--muted-foreground)) }
  .card-bd { padding: 1rem 1.25rem }
  .chart { position: relative; height: 280px }
  .chart.tall { height: 480px }
  .chart.traj { height: 340px }
  .marks { display: flex; flex-wrap: wrap; gap: .35rem 1rem; margin: 0 0 .65rem; font-size: .75rem; color: hsl(var(--muted-foreground)) }
  .marks .mi { display: inline-flex; align-items: center; gap: .35rem }
  .marks b { font-weight: 600; color: hsl(var(--foreground)) }
  .dot { width: 8px; height: 8px; border-radius: 99px; border: 2px solid; display: inline-block; box-sizing: border-box; flex-shrink: 0 }
  .dot.fact { background: currentColor }
  .dot.plan { background: #fff }
  .row { display: flex; gap: .75rem; align-items: center; flex-wrap: wrap; margin-bottom: .75rem }
  .row select {
    font: inherit; height: 36px; padding: 0 .75rem; min-width: 260px; max-width: 100%;
    border: 1px solid hsl(var(--border)); border-radius: .5rem; background: #fff;
  }
  .tbl-wrap { max-height: 640px; overflow: auto }
  table.full { width: 100%; min-width: 1180px; border-collapse: collapse; font-size: .8125rem }
  table.full th, table.full td { padding: .7rem .85rem; border-bottom: 1px solid hsl(var(--border)); text-align: left; vertical-align: top }
  table.full th {
    position: sticky; top: 0; z-index: 2; background: hsl(var(--background));
    color: hsl(var(--muted-foreground)); font-size: .6875rem; font-weight: 600;
    letter-spacing: .04em; text-transform: uppercase; white-space: nowrap;
    box-shadow: 0 1px 0 hsl(var(--border));
  }
  table.full td.r, table.full th.r { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap }
  table.dense .obj { min-width: 22rem; max-width: 36rem }
  table.dense .obj-name { display: block; font-weight: 600; line-height: 1.4; overflow-wrap: break-word }
  table.dense .obj-sub { display: block; margin-top: .35rem; font-size: .8rem; color: hsl(var(--muted-foreground)); line-height: 1.35 }
  table.dense .obj-uin { display: block; margin-top: .2rem; font-size: .7rem; color: hsl(var(--muted-foreground)); font-variant-numeric: tabular-nums }
  table.dense .work { white-space: normal; min-width: 8.5rem }
  table.dense .cell-stack { display: flex; flex-direction: column; align-items: flex-end; gap: 2px }
  table.dense .cell-stack.start { align-items: flex-start }
  table.dense .exp-date { font-size: .7rem; color: hsl(var(--muted-foreground)); font-variant-numeric: tabular-nums }
  table.dense .amt { font-weight: 600 }
  table.dense .pct { font-size: .7rem; color: hsl(var(--muted-foreground)) }
  .amt.neg, .pct.neg { color: var(--err) }
  .amt.pos, .pct.pos { color: var(--ok) }
  tr.clickable { cursor: pointer }
  tr.clickable:hover { background: hsl(var(--muted)) }
  tr.row-active { background: rgba(31,134,130,.08) }
  .pill {
    display: inline-flex; align-items: center; font-size: .7rem; font-weight: 600;
    padding: 2px 8px; border-radius: 999px;
  }
  .pill.ok { background: var(--ok-bg); color: #15803d }
  .pill.warn { background: var(--warn-bg); color: #b45309 }
  .pill.err { background: var(--err-bg); color: #b91c1c }
  .quad-tools { display: flex; flex-wrap: wrap; gap: .75rem; align-items: center; margin-bottom: .75rem; font-size: .8125rem; color: hsl(var(--muted-foreground)) }
  .quad-tools input[type=number] { width: 64px; font: inherit; height: 32px; padding: 0 .5rem; border: 1px solid hsl(var(--border)); border-radius: .5rem; color: hsl(var(--foreground)) }
  .quad-cards { display: grid; grid-template-columns: repeat(2, 1fr); gap: .5rem; margin-bottom: .75rem }
  @media (min-width: 800px) { .quad-cards { grid-template-columns: repeat(4, 1fr) } }
  .quad-card { border: 1px solid hsl(var(--border)); border-left: 4px solid var(--q-color, #94a3b8); border-radius: .5rem; padding: .6rem .75rem; cursor: pointer; background: #fff }
  .quad-card.active { box-shadow: inset 0 0 0 1px var(--q-color); background: hsl(var(--muted)) }
  .quad-card .qn { font-size: 1.35rem; font-weight: 700; line-height: 1.1; font-variant-numeric: tabular-nums }
  .quad-card .ql { margin-top: .2rem; font-size: .75rem; font-weight: 600 }
</style>
</head>
<body>
<header class="glass-nav">
  <div class="container hdr">
    <div class="brand">
      <img class="cks" src="data:image/png;base64,__CKS_LOGO__" alt="ЦКС">
      <img class="mo" src="data:image/png;base64,__MO_LOGO__" alt="Московская область">
      <div class="brand-meta">
        <p class="t">Минстрой — СГ и выплаты</p>
        <p class="s">Портфель Минстроя · СГ и выплаты по Simple List</p>
      </div>
    </div>
    <div class="nav-center">
      <div class="nav-pill">
        <button class="nav-tab active" type="button" data-tab="summary">Сводка</button>
        <button class="nav-tab" type="button" data-tab="contractors">Подрядчики</button>
        <button class="nav-tab" type="button" data-tab="objects">Объекты</button>
      </div>
    </div>
  </div>
</header>
<main class="container">
<div class="stack">
"""

PAGE_BODY = r"""
  <div class="tabpanel stack" data-tab="summary">
  <div class="kpis">
    <div class="kpi"><p class="l">Объектов</p><p class="n" id="mzN">—</p></div>
    <div class="kpi"><p class="l">Медиана СГ, %</p><p class="n" id="mzMedSg">—</p></div>
    <div class="kpi"><p class="l">Медиана оплаты, %</p><p class="n" id="mzMedPct">—</p></div>
    <div class="kpi"><p class="l">Кредитует подрядчик</p><p class="n warn" id="mzCredit">—</p></div>
  </div>

  <div class="card">
    <div class="card-hd">
      <div>
        <h2>Бюджет</h2>
        <p id="mzBudgetNote">Лимиты — сколько предусмотрели. Доведено — сколько уже пришло на объекты.</p>
      </div>
    </div>
    <div class="card-bd"><div class="kpis" id="mzBudget"></div></div>
  </div>

  <div class="card">
    <div class="card-hd">
      <div>
        <h2>Объекты</h2>
        <p>Клик по строке — траектория на вкладке «Объекты»</p>
      </div>
    </div>
    <div class="tbl-wrap">
      <table class="full dense">
        <thead><tr>
          <th>Объект</th>
          <th>Вид работ</th>
          <th class="r">Год откр.</th>
          <th>Заказчик</th>
          <th class="r">Лимит, млн</th><th class="r">Контракт, млн</th>
          <th class="r">СГ, %</th><th class="r">Оплата, %</th>
          <th class="r" title="Согласованная смета после экспертизы">После экспертизы, млн</th>
          <th class="r">Δ сметы</th>
          <th title="Наведите на дату: ТЧ — техчасть, СД — смета, ТЧ+СД — оба">Экспертиза</th>
          <th>Статус</th>
        </tr></thead>
        <tbody id="mzTbl"></tbody>
      </table>
    </div>
  </div>
  </div>

  <div class="tabpanel stack" data-tab="contractors" hidden>
  <div class="card">
    <div class="card-hd">
      <div>
        <h2>Квадранты подрядчиков</h2>
        <p>Ось X — стройготовность, Y — оплата. Размер точки — лимит. Крест — пороги.</p>
      </div>
    </div>
    <div class="card-bd">
      <div class="quad-tools">
        <label>Порог СГ, % <input type="number" id="qSg" min="0" max="100" step="5" value="50"></label>
        <label>Порог оплаты, % <input type="number" id="qPay" min="0" max="100" step="5" value="50"></label>
      </div>
      <div class="quad-cards" id="quadCards"></div>
      <div class="chart" style="height:360px"><canvas id="cQuad"></canvas></div>
    </div>
  </div>
  </div>

  <div class="tabpanel stack" data-tab="objects" hidden>
  <div class="filters">
    <label class="field">
      <span>Год открытия · ⌘/Ctrl</span>
      <select id="fltYear" multiple size="4"></select>
    </label>
    <label class="field">
      <span>Заказчик · ⌘/Ctrl</span>
      <select id="fltCustomer" multiple size="4"></select>
    </label>
    <label class="field">
      <span>Вид работ · ⌘/Ctrl</span>
      <select id="fltWork" multiple size="4">
        <option value="капремонт">Капремонт</option>
        <option value="строительство">Строительство</option>
        <option value="реконструкция">Реконструкция</option>
        <option value="прочее">Прочее</option>
      </select>
    </label>
    <label class="field">
      <span>Статус · ⌘/Ctrl</span>
      <select id="fltStatus" multiple size="4">
        <option value="credit">Кредитует</option>
        <option value="over">Избыток оплаты</option>
        <option value="balanced">Баланс</option>
        <option value="sg">СГ &gt; 0</option>
        <option value="exp">Есть экспертиза</option>
        <option value="noexp">Без экспертизы</option>
      </select>
    </label>
    <label class="field select-row">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>
      <input id="fltSearch" type="search" placeholder="Поиск..." autocomplete="off">
    </label>
  </div>

  <div class="card">
    <div class="card-hd">
      <div>
        <h2>Матрица риска: готовность vs оплата</h2>
        <p>Снизу оплата, слева СГ. Зелёная полоса — коридор ±10 п.п. Клик по точке — траектория.</p>
      </div>
    </div>
    <div class="card-bd">
      <div class="chart tall"><canvas id="cMzMatrix"></canvas></div>
    </div>
  </div>

  <div class="card">
    <div class="card-hd">
      <div>
        <h2>Траектория объекта</h2>
        <p id="mzSelMeta">Выберите объект</p>
      </div>
    </div>
    <div class="card-bd">
      <div class="row"><select id="selMz"></select></div>
      <div id="mzFunds" class="marks"></div>
      <div id="mzMarks" class="marks"></div>
      <div class="chart traj"><canvas id="cMzTraj"></canvas></div>
    </div>
  </div>
  </div>
"""

PAGE_SCRIPT = r"""</div>
</main>
<script>
const DATA = __DATA__;
const blue='#0B6CB8', green='#1f8682';
const STATUS_SHORT = { over:'Избыток', credit:'Кредитует', balanced:'Баланс', unknown:'—' };
const STATUS_PILL = { over:'err', credit:'warn', balanced:'ok', unknown:'' };
const MATRIX_STATUS = {
  credit:   { label: 'Кредитует подрядчик', color: '#1f8682', fill: 'rgba(31,134,130,.82)' },
  balanced: { label: 'В коридоре (±10 п.п.)', color: '#16a34a', fill: 'rgba(22,163,74,.85)' },
  over:     { label: 'Избыток оплаты', color: '#0B6CB8', fill: 'rgba(11,108,184,.78)' },
  unknown:  { label: 'Нет данных', color: '#94a3b8', fill: 'rgba(148,163,184,.7)' },
};
const matrixZonesPlugin = {
  id: 'matrixZones',
  beforeDatasetsDraw(chart) {
    const { ctx, scales: { x, y } } = chart;
    if (!x || !y) return;
    const x0 = x.getPixelForValue(0), x1 = x.getPixelForValue(100);
    const y0 = y.getPixelForValue(0), y1 = y.getPixelForValue(100);
    const band = 10;
    ctx.save();
    ctx.beginPath();
    ctx.moveTo(x0, y0); ctx.lineTo(x0, y1); ctx.lineTo(x1, y1); ctx.closePath();
    ctx.fillStyle = 'rgba(31,134,130,.10)'; ctx.fill();
    ctx.beginPath();
    ctx.moveTo(x0, y0); ctx.lineTo(x1, y0); ctx.lineTo(x1, y1); ctx.closePath();
    ctx.fillStyle = 'rgba(11,108,184,.08)'; ctx.fill();
    ctx.beginPath();
    for (let pay = 0; pay <= 100; pay++) {
      const sgHi = Math.min(100, pay + band);
      const px = x.getPixelForValue(pay), py = y.getPixelForValue(sgHi);
      if (pay === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    for (let pay = 100; pay >= 0; pay--) {
      const sgLo = Math.max(0, pay - band);
      ctx.lineTo(x.getPixelForValue(pay), y.getPixelForValue(sgLo));
    }
    ctx.closePath();
    ctx.fillStyle = 'rgba(22,163,74,.18)'; ctx.fill();
    ctx.strokeStyle = 'rgba(22,163,74,.5)'; ctx.lineWidth = 1.5;
    ctx.setLineDash([4, 3]);
    ctx.beginPath();
    ctx.moveTo(x.getPixelForValue(0), y.getPixelForValue(0));
    ctx.lineTo(x.getPixelForValue(100), y.getPixelForValue(100));
    ctx.stroke();
    ctx.restore();
  },
};

function bubbleRadius(contractMln) {
  return Math.max(4, Math.min(14, 3 + Math.sqrt(Math.max(contractMln || 0, 0)) * 0.28));
}

const MZ = DATA;
let mzYears = ['2026'], mzCustomers = [], mzWorks = [], mzStatuses = [];
let mzSearch = '', mzUin = null, mzFocusUin = null;
let chartMzMatrix = null, chartMzTraj = null, chartQuad = null;

function selectedValues(sel) {
  return [...sel.selectedOptions].map(o => o.value);
}
function statusMatch(o, st) {
  if (st === 'credit') return o.money_status === 'credit';
  if (st === 'over') return o.money_status === 'over';
  if (st === 'balanced') return o.money_status === 'balanced';
  if (st === 'sg') return o.sg > 0;
  if (st === 'exp') return !!(o.exp_done || o.exp_agreed_mln != null);
  if (st === 'noexp') return !(o.exp_done || o.exp_agreed_mln != null);
  return true;
}
function mzPass(o) {
  if (mzYears.length && !mzYears.includes(String(o.open_year || ''))) return false;
  if (mzCustomers.length && !mzCustomers.includes(o.customer || '')) return false;
  if (mzWorks.length && !mzWorks.includes(o.bucket)) return false;
  if (mzStatuses.length && !mzStatuses.some(st => statusMatch(o, st))) return false;
  if (mzSearch) {
    const q = mzSearch.toLowerCase();
    const blob = [o.name, o.full, o.contractor, o.customer, o.ownership, o.work, o.uin, o.open_year]
      .filter(Boolean).join(' ').toLowerCase();
    if (!blob.includes(q)) return false;
  }
  return true;
}
function mzFiltered() { return (MZ.objects || []).filter(mzPass); }

function mzSelectObject(uin, { focus = false } = {}) {
  if (!uin) return;
  mzUin = uin;
  if (focus) mzFocusUin = uin;
  const sel = document.getElementById('selMz');
  if (sel) sel.value = mzUin;
  drawMzTraj(mzUin);
  renderMzTbl();
}

function renderMzMatrix() {
  const rows = mzFiltered().filter(o => o.sg != null && o.pct != null);
  const byStatus = { credit:[], over:[], balanced:[], unknown:[] };
  rows.forEach(o => (byStatus[o.money_status] || byStatus.unknown).push(o));
  const datasets = ['credit','balanced','over','unknown'].map(st => {
    const meta = MATRIX_STATUS[st] || MATRIX_STATUS.unknown;
    return {
      label: meta.label,
      data: byStatus[st].map(o => {
        const size = o.limit || o.contract_value || 0;
        return {
          x: o.pct, y: o.sg, r: bubbleRadius(size),
          uin: o.uin, name: o.name,
          contract: o.contract_value, limit: o.limit,
        };
      }),
      backgroundColor: meta.fill,
      borderColor: meta.color,
      borderWidth: 1,
    };
  });
  const cfg = {
    type: 'bubble',
    data: { datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 11 } } },
        datalabels: { display: false },
        tooltip: { callbacks: { label(ctx) {
          const d = ctx.raw;
          return `${d.name}: СГ ${d.y}% · оплата ${d.x}% · лимит ${d.limit ?? '—'} · контракт ${d.contract ?? '—'} млн`;
        }}},
      },
      scales: {
        x: { min:0, max:100, title:{ display:true, text:'Оплата (от лимита), %', font:{ size:13, weight:'600' } } },
        y: { min:0, max:100, title:{ display:true, text:'Стройготовность (факт), %', font:{ size:13, weight:'600' } } },
      },
      onClick(evt, els) {
        if (!els.length) {
          if (mzFocusUin) { mzFocusUin = null; renderMzTbl(); }
          return;
        }
        const d = chartMzMatrix.data.datasets[els[0].datasetIndex].data[els[0].index];
        if (!d || !d.uin) return;
        mzSelectObject(d.uin, { focus: true });
      },
    },
    plugins: [matrixZonesPlugin],
  };
  if (chartMzMatrix) chartMzMatrix.destroy();
  chartMzMatrix = new Chart(document.getElementById('cMzMatrix'), cfg);
}

function parseMzDate(s) {
  if (!s) return null;
  const m = String(s).match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (!m) return null;
  let y = +m[3]; if (y < 100) y += 2000;
  return new Date(y, +m[2]-1, +m[1]).getTime();
}

const MARK_SPEC = {
  exp:  { name: 'Экспертиза', short: 'Эксп.', color: '#7c3aed' },
  smr:  { name: 'СМР', short: 'СМР', color: '#d97706' },
  open: { name: 'Открытие', short: 'Откр.', color: '#0369a1' },
};

function fullDate(s) {
  const m = String(s || '').match(/^(\d{2})\.(\d{2})\.(\d{2})$/);
  if (!m) return s || '';
  return `${m[1]}.${m[2]}.20${m[3]}`;
}

function trajFrame(traj, marks) {
  const by = new Map();
  (traj || []).forEach(p => { if (p && p.d) by.set(p.d, p); });
  const dates = [...by.keys()];
  ['exp', 'smr', 'open'].forEach(k => {
    const m = (marks || {})[k] || {};
    ['plan', 'fact'].forEach(kind => { if (m[kind] && !by.has(m[kind]) && !dates.includes(m[kind])) dates.push(m[kind]); });
  });
  dates.sort((a, b) => (parseMzDate(a) || 0) - (parseMzDate(b) || 0));
  return dates.map(d => by.get(d) || { d, sg: null, pct: null });
}

function markEvents(marks) {
  const events = [];
  ['exp', 'smr', 'open'].forEach(k => {
    const spec = MARK_SPEC[k];
    const m = (marks || {})[k] || {};
    if (m.plan) events.push({ key: k + '-plan', name: spec.name, short: spec.short, color: spec.color, kind: 'plan', d: m.plan });
    if (m.fact) events.push({ key: k + '-fact', name: spec.name, short: spec.short, color: spec.color, kind: 'fact', d: m.fact });
  });
  events.sort((a, b) => (parseMzDate(a.d) || 0) - (parseMzDate(b.d) || 0));
  let prev = null, lane = 0;
  events.forEach(e => {
    const t = parseMzDate(e.d);
    if (prev != null && t != null && Math.abs(t - prev) < 21 * 86400000) lane = (lane + 1) % 3;
    else lane = 0;
    e.lane = lane;
    prev = t;
  });
  return events;
}

const FUND_LABEL = { fed:'Федеральный', mo:'Бюджет МО', mun:'Муниципальный', ibk:'ИБК', kik:'КИК' };

function rub(v) {
  if (v == null || !isFinite(v)) return '—';
  const sign = v < 0 ? '−' : '';
  const a = Math.abs(v);
  const fmt = (n, d) => n.toLocaleString('ru-RU', { minimumFractionDigits: d, maximumFractionDigits: d });
  if (a >= 1000) return sign + fmt(a / 1000, 1) + ' млрд';
  return sign + fmt(a, a >= 100 ? 0 : 1) + ' млн';
}
function esc(s) {
  return String(s ?? '').replace(/[&<>"]/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' }[c]));
}
function fundTotals() {
  const acc = {};
  MZ.objects.forEach(o => (o.funds || []).forEach(f => {
    const s = acc[f.k] || { plan: 0, fact: 0, n: 0 };
    s.plan += f.plan || 0; s.fact += f.fact || 0; s.n += 1;
    acc[f.k] = s;
  }));
  return acc;
}
function renderPortfolio() {
  const acc = fundTotals();
  const fed = acc.fed || { plan: 0, fact: 0, n: 0 };
  const mo = acc.mo || { plan: 0, fact: 0, n: 0 };
  const mun = acc.mun || { plan: 0, fact: 0, n: 0 };
  const ibk = acc.ibk || { plan: 0, fact: 0, n: 0 };
  const kik = acc.kik || { plan: 0, fact: 0, n: 0 };
  const infra = MZ.objects.filter(o => (o.funds || []).some(f => f.k === 'ibk' || f.k === 'kik'));
  const tiles = [
    { l: 'Федеральный', n: fed.n, plan: fed.plan, fact: fed.fact },
    { l: 'Бюджет МО', n: mo.n, plan: mo.plan, fact: mo.fact },
    { l: 'Муниципальный', n: mun.n, plan: mun.plan, fact: mun.fact },
    { l: 'Инфраструктурные кредиты', n: infra.length, plan: ibk.plan + kik.plan, fact: ibk.fact + kik.fact },
  ];
  document.getElementById('mzBudget').innerHTML = tiles.map(t =>
    `<div class="kpi money"><p class="l">${t.l} · ${t.n === 1 ? '1 объект' : t.n + ' объектов'}</p>` +
    `<p class="pair"><span>Лимиты</span><b>${rub(t.plan)}</b></p>` +
    `<p class="pair"><span>Доведено</span><b>${rub(t.fact)}</b></p></div>`
  ).join('');
}

const QUAD_META = {
  stable:  { label: 'Устойчивые', color: '#16a34a' },
  cashgap: { label: 'Кредитует подрядчик', color: '#d97706' },
  advance: { label: 'Аванс впереди работ', color: '#7c3aed' },
  early:   { label: 'Ранняя стадия', color: '#0B6CB8' },
};
let quadState = { sg: 50, pay: 50, active: null };

function quadrantOf(sg, pct) {
  if (sg == null || pct == null) return null;
  const hiG = sg >= quadState.sg, hiP = pct >= quadState.pay;
  if (hiG && hiP) return 'stable';
  if (hiG && !hiP) return 'cashgap';
  if (!hiG && hiP) return 'advance';
  return 'early';
}

function contractorPoints() {
  const byC = new Map();
  MZ.objects.forEach(o => {
    const name = o.contractor || 'Без подрядчика';
    const e = byC.get(name) || { contractor: name, n: 0, limit: 0, wSg: 0, wPct: 0, w: 0 };
    const w = o.limit || o.contract_value || 0;
    e.n += 1; e.limit += w;
    if (o.sg != null && o.pct != null && w > 0) {
      e.wSg += o.sg * w; e.wPct += o.pct * w; e.w += w;
    }
    byC.set(name, e);
  });
  return [...byC.values()].map(e => {
    const avg_sg = e.w ? Math.round(e.wSg / e.w * 10) / 10 : null;
    const avg_pct = e.w ? Math.round(e.wPct / e.w * 10) / 10 : null;
    return Object.assign(e, { avg_sg, avg_pct, quad: quadrantOf(avg_sg, avg_pct) });
  });
}

const quadCrossPlugin = {
  id: 'quadCross',
  beforeDatasetsDraw(chart) {
    const { ctx, scales: { x, y } } = chart;
    if (!x || !y) return;
    const gx = x.getPixelForValue(quadState.sg);
    const py = y.getPixelForValue(quadState.pay);
    ctx.save();
    ctx.strokeStyle = 'rgba(15,23,42,.35)';
    ctx.setLineDash([5, 4]);
    ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(gx, y.top); ctx.lineTo(gx, y.bottom); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(x.left, py); ctx.lineTo(x.right, py); ctx.stroke();
    ctx.restore();
  },
};

function renderQuad() {
  const rows = contractorPoints();
  const counts = { stable: 0, cashgap: 0, advance: 0, early: 0 };
  rows.forEach(r => { if (r.quad) counts[r.quad]++; });
  document.getElementById('quadCards').innerHTML = Object.keys(QUAD_META).map(k => {
    const m = QUAD_META[k];
    const active = quadState.active === k ? ' active' : '';
    return `<div class="quad-card${active}" data-quad="${k}" style="--q-color:${m.color}">` +
      `<div class="qn">${counts[k]}</div><div class="ql">${m.label}</div></div>`;
  }).join('');
  document.querySelectorAll('#quadCards .quad-card').forEach(el => {
    el.onclick = () => {
      quadState.active = quadState.active === el.dataset.quad ? null : el.dataset.quad;
      renderQuad();
    };
  });
  const visible = quadState.active ? rows.filter(r => r.quad === quadState.active) : rows;
  const datasets = Object.keys(QUAD_META).map(k => {
    const m = QUAD_META[k];
    const data = visible.filter(r => r.quad === k).map(r => ({
      x: r.avg_sg, y: r.avg_pct,
      r: Math.max(4, Math.min(16, 3 + Math.sqrt(Math.max(r.limit, 0)) * 0.22)),
      contractor: r.contractor, limit: r.limit, n: r.n,
    }));
    return { label: m.label, data, backgroundColor: m.color + '99', borderColor: m.color, borderWidth: 1 };
  });
  const cfg = {
    type: 'bubble',
    data: { datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 11 } } },
        datalabels: { display: false },
        tooltip: {
          callbacks: {
            label(ctx) {
              const d = ctx.raw;
              return [`${d.contractor}`, `СГ ${d.x}% · оплата ${d.y}%`, `${d.n} объектов · лимит ${rub(d.limit)}`];
            },
          },
        },
      },
      scales: {
        x: { min: 0, max: 100, title: { display: true, text: 'Стройготовность, %' } },
        y: { min: 0, max: 100, title: { display: true, text: 'Оплата, %' } },
      },
      onClick(evt, els) {
        if (!els.length) return;
        const d = chartQuad.data.datasets[els[0].datasetIndex].data[els[0].index];
        if (!d || !d.contractor || d.contractor === 'Без подрядчика') return;
        const input = document.getElementById('fltSearch');
        input.value = d.contractor;
        mzSearch = d.contractor;
        mzFocusUin = null;
        showTab('objects');
        renderMzMatrix();
        renderMzTbl();
      },
    },
    plugins: [quadCrossPlugin],
  };
  if (chartQuad) chartQuad.destroy();
  chartQuad = new Chart(document.getElementById('cQuad'), cfg);
}

function drawMzTraj(uin) {
  const o = MZ.objects.find(x => x.uin === uin);
  const el = document.getElementById('mzSelMeta');
  const marksEl = document.getElementById('mzMarks');
  const fundsEl = document.getElementById('mzFunds');
  if (!o) { el.textContent = ''; if (marksEl) marksEl.innerHTML = ''; if (fundsEl) fundsEl.innerHTML = ''; return; }
  const bits = [];
  if (o.exp_agreed_mln != null) {
    const dlt = o.exp_delta_mln != null
      ? ` (${o.exp_delta_mln>0?'+':''}${o.exp_delta_mln} млн` + (o.exp_overrun!=null ? `, ${o.exp_overrun>0?'+':''}${o.exp_overrun}%` : '') + ')'
      : '';
    bits.push(`смета после эксп. ${o.exp_agreed_mln} млн${dlt}`);
  }
  if (o.exp_done) bits.push(`эксп. ${o.exp_done}`);
  else if (o.exp_in) bits.push(`заход на эксп. ${o.exp_in}`);
  el.textContent = `${o.full || o.name}: СГ ${o.sg ?? '—'}% · оплата ${o.pct ?? '—'}%` +
    (bits.length ? ' · ' + bits.join(' · ') : '');

  if (fundsEl) {
    fundsEl.innerHTML = (o.funds || []).map(f =>
      `<span class="mi"><b>${FUND_LABEL[f.k] || f.k}</b> заложено ${rub(f.plan)} · доведено ${rub(f.fact)}</span>`
    ).join('');
  }

  const marks = o.marks || {};
  if (marksEl) {
    marksEl.innerHTML = ['exp', 'smr', 'open'].map(k => {
      const spec = MARK_SPEC[k];
      const m = marks[k] || {};
      if (!m.plan && !m.fact) return '';
      const parts = [];
      if (m.plan) parts.push(`план ${fullDate(m.plan)}`);
      if (m.fact) parts.push(`факт ${fullDate(m.fact)}`);
      return `<span class="mi" style="color:${spec.color}"><i class="dot fact"></i><b>${spec.name}</b> ${parts.join(' · ')}</span>`;
    }).join('');
  }

  const frame = trajFrame(o.traj || [], marks);
  const labels = frame.map(p => p.d);
  const annotations = {};
  markEvents(marks).forEach(e => {
    const title = e.kind === 'plan' ? 'план' : 'факт';
    annotations[e.key + '-line'] = {
      type: 'line', xMin: e.d, xMax: e.d,
      borderColor: e.color, borderWidth: e.kind === 'fact' ? 1.5 : 1,
      borderDash: e.kind === 'plan' ? [4, 3] : [],
      label: {
        display: true,
        content: `${e.short} ${title}`,
        position: 'end',
        yAdjust: 6 + e.lane * 15,
        color: e.color,
        backgroundColor: 'rgba(255,255,255,.92)',
        font: { size: 10, weight: '600' },
        padding: { top: 1, bottom: 1, left: 3, right: 3 },
      },
    };
    annotations[e.key] = {
      type: 'point', xValue: e.d, yValue: 3 + e.lane * 6, radius: 5,
      backgroundColor: e.kind === 'fact' ? e.color : '#fff',
      borderColor: e.color, borderWidth: 2,
    };
  });
  const cfg = {
    type: 'line',
    data: {
      labels,
      datasets: [
        { label:'СГ факт', data: frame.map(p => p.sg), borderColor: blue, tension:.2, pointRadius:2, spanGaps:true },
        { label:'Оплата %', data: frame.map(p => p.pct), borderColor: green, tension:.2, pointRadius:2, spanGaps:true },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      layout: { padding: { top: 4 } },
      plugins: {
        legend:{ position:'bottom' },
        datalabels:{ display:false },
        annotation: { clip: false, annotations },
        tooltip: {
          callbacks: {
            label(ctx) {
              const v = ctx.parsed.y;
              if (v == null) return null;
              return `${ctx.dataset.label}: ${v}`;
            },
          },
        },
      },
      scales: { y:{ min:0, max:100 } },
    },
  };
  if (chartMzTraj) chartMzTraj.destroy();
  chartMzTraj = new Chart(document.getElementById('cMzTraj'), cfg);
}

function renderMzTbl() {
  const rows = MZ.objects.slice().sort((a,b) => (b.sg||0)-(a.sg||0));
  const wrap = document.getElementById('mzTbl').closest('.tbl-wrap');
  document.getElementById('mzTbl').innerHTML = rows.map(o => {
    const active = o.uin === mzUin ? ' row-active' : '';
    const ovCls = o.exp_overrun > 0 ? 'pos' : (o.exp_overrun < 0 ? 'neg' : '');
    const deltaCell = o.exp_delta_mln != null
      ? `<div class="cell-stack end"><span class="amt ${ovCls}">${o.exp_delta_mln>0?'+':''}${o.exp_delta_mln}</span>` +
        (o.exp_overrun!=null ? `<span class="pct ${ovCls}">${o.exp_overrun>0?'+':''}${o.exp_overrun}%</span>` : '') +
        `</div>`
      : '—';
    const KIND = { both: 'ТЧ+СД — техчасть и смета', tch: 'ТЧ — только техчасть', sd: 'СД — только смета' };
    const kindTip = KIND[o.exp_kind] || '';
    const expTip = [kindTip, o.exp_in ? 'заход ' + o.exp_in : '', o.exp_start ? 'старт ' + o.exp_start : ''].filter(Boolean).join(' · ');
    let expCell = '—';
    if (o.exp_done) {
      expCell = `<span title="${esc(expTip)}">${o.exp_done}</span>`;
    } else if (o.exp_in) {
      expCell = `<span class="pill warn" title="${esc(expTip || ('заход ' + o.exp_in))}">в эксп.</span>`;
    } else if (o.entered_exp) {
      expCell = `<span class="pill warn" title="${esc(kindTip || 'заявка на экспертизу')}">заявка</span>`;
    }
    const riskTip = o.risk ? ` title="${esc(o.risk)}"` : '';
    return `<tr class="clickable${active}" data-uin="${o.uin}">` +
      `<td class="obj"${riskTip}>` +
        `<span class="obj-name">${esc(o.full || o.name)}</span>` +
        (o.contractor ? `<span class="obj-sub">${esc(o.contractor)}</span>` : '') +
        (o.ownership ? `<span class="obj-sub">${esc(o.ownership)}</span>` : '') +
        `<span class="obj-uin">${esc(o.uin)}</span>` +
      `</td>` +
      `<td class="work">${esc(o.work || '—')}</td>` +
      `<td class="r">${o.open_year??'—'}</td>` +
      `<td class="work">${esc(o.customer || '—')}</td>` +
      `<td class="r">${o.limit??'—'}</td>` +
      `<td class="r">${o.contract_value??'—'}</td>` +
      `<td class="r">${o.sg??'—'}</td><td class="r">${o.pct??'—'}</td>` +
      `<td class="r">${o.exp_agreed_mln??'—'}</td>` +
      `<td class="r">${deltaCell}</td>` +
      `<td>${expCell}</td>` +
      `<td><span class="pill ${STATUS_PILL[o.money_status]||''}">${STATUS_SHORT[o.money_status]||'—'}</span></td></tr>`;
  }).join('') || '<tr><td colspan="12" class="note">Нет объектов</td></tr>';
  if (wrap) wrap.scrollTop = 0;
  document.querySelectorAll('#mzTbl tr.clickable').forEach(tr => {
    tr.onclick = () => {
      mzUin = tr.dataset.uin;
      mzFocusUin = null;
      const sel = document.getElementById('selMz');
      if (sel) sel.value = mzUin;
      renderMzTbl();
      window.showTab('objects');
    };
  });
}

(function init() {
  function showTab(tab) {
    document.querySelectorAll('.nav-tab').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
    document.querySelectorAll('.tabpanel').forEach(p => { p.hidden = p.dataset.tab !== tab; });
    const panel = document.querySelector('.tabpanel[data-tab="' + tab + '"]');
    if (panel) panel.getBoundingClientRect();
    if (tab === 'contractors') renderQuad();
    if (tab === 'objects') {
      renderMzMatrix();
      if (mzUin) drawMzTraj(mzUin);
    }
  }
  window.showTab = showTab;
  document.querySelectorAll('.nav-tab').forEach(b => { b.onclick = () => showTab(b.dataset.tab); });

  const m = MZ.meta || {};
  document.getElementById('mzN').textContent = (m.n || MZ.objects.length).toLocaleString('ru-RU');
  document.getElementById('mzMedSg').textContent = m.median_sg != null ? m.median_sg : '—';
  document.getElementById('mzMedPct').textContent = m.median_pct != null ? m.median_pct : '—';
  document.getElementById('mzCredit').textContent = m.n_credit != null ? m.n_credit : '—';
  renderPortfolio();
  const onThresh = (id, key) => {
    const el = document.getElementById(id);
    const apply = () => {
      const v = parseFloat(el.value);
      if (!Number.isNaN(v)) { quadState[key] = Math.max(0, Math.min(100, v)); renderQuad(); }
    };
    el.oninput = apply;
    el.onchange = apply;
  };
  onThresh('qSg', 'sg');
  onThresh('qPay', 'pay');

  const sel = document.getElementById('selMz');
  sel.innerHTML = '';
  MZ.objects.forEach(o => {
    const opt = document.createElement('option');
    opt.value = o.uin; opt.textContent = o.name;
    sel.appendChild(opt);
  });
  mzUin = MZ.objects[0]?.uin;
  if (mzUin) sel.value = mzUin;
  sel.onchange = e => mzSelectObject(e.target.value, { focus: true });

  const yearSel = document.getElementById('fltYear');
  const years = [...new Set(MZ.objects.map(o => o.open_year).filter(Boolean))].sort();
  years.forEach(y => {
    const opt = document.createElement('option');
    opt.value = y; opt.textContent = y;
    if (y === '2026') opt.selected = true;
    yearSel.appendChild(opt);
  });
  const custSel = document.getElementById('fltCustomer');
  const customers = [...new Set(MZ.objects.map(o => o.customer).filter(Boolean))].sort(
    (a, b) => a.localeCompare(b, 'ru')
  );
  customers.forEach(c => {
    const opt = document.createElement('option');
    opt.value = c; opt.textContent = c;
    custSel.appendChild(opt);
  });

  const workSel = document.getElementById('fltWork');
  const statusSel = document.getElementById('fltStatus');
  const refresh = () => {
    mzYears = selectedValues(yearSel);
    mzCustomers = selectedValues(custSel);
    mzWorks = selectedValues(workSel);
    mzStatuses = selectedValues(statusSel);
    mzFocusUin = null;
    renderMzMatrix();
    renderMzTbl();
  };
  yearSel.onchange = refresh;
  custSel.onchange = refresh;
  workSel.onchange = refresh;
  statusSel.onchange = refresh;
  let searchTimer = null;
  document.getElementById('fltSearch').oninput = e => {
    clearTimeout(searchTimer);
    const v = e.target.value.trim();
    searchTimer = setTimeout(() => { mzSearch = v; refresh(); }, 200);
  };
  refresh();
})();
</script>
</body>
</html>
"""


def main():
    if not DATA.exists():
        sys.exit(
            f"No {DATA} — сначала: python3 scripts/build_minstroy_dynamics.py"
        )
    payload = json.loads(DATA.read_text(encoding="utf-8"))
    n = len(payload.get("objects") or [])
    meta = payload.get("meta") or {}
    html = (
        EXON_HEAD.replace("__CKS_LOGO__", CKS_LOGO_B64).replace("__MO_LOGO__", MO_LOGO_B64)
        + PAGE_BODY
        + PAGE_SCRIPT.replace("__DATA__", json.dumps(payload, ensure_ascii=False, separators=(",", ":")))
    )
    OUT_HTML.write_text(html, encoding="utf-8")
    print(
        f"OK {OUT_HTML} — {n} объектов, as_of={meta.get('as_of')}, "
        f"size={OUT_HTML.stat().st_size // 1024} kb"
    )


if __name__ == "__main__":
    main()

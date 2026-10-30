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
FAVICON_SVG = (ROOT / "assets" / "favicon.svg").read_text(encoding="utf-8")
FAVICON_SVG_HREF = "data:image/svg+xml;base64," + base64.b64encode(FAVICON_SVG.encode("utf-8")).decode("ascii")
FAVICON_ICO_HREF = "data:image/x-icon;base64," + base64.b64encode(
    (ROOT / "assets" / "favicon.ico").read_bytes()
).decode("ascii")


# Дизайн 1:1 с Exon_2 Header + Index (glass-nav, Inter, soft cards, select+search).
EXON_HEAD = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Минстрой — СГ и выплаты</title>
<link rel="icon" href="__FAVICON_SVG__" type="image/svg+xml" />
<link rel="alternate icon" href="__FAVICON_ICO__" type="image/x-icon" />
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
  /* Фильтры: кнопки + чекбоксы + чипы */
  .filters { display: flex; flex-direction: column; gap: .65rem }
  .filter-row { display: flex; flex-wrap: wrap; gap: .5rem; align-items: center }
  .filter-dd { position: relative }
  .filter-btn {
    display: inline-flex; align-items: center; gap: .4rem;
    height: 36px; padding: 0 .85rem; border-radius: .5rem;
    border: 1px solid hsl(var(--border)); background: #fff;
    font: 500 13px Inter, sans-serif; color: hsl(var(--foreground));
    cursor: pointer; white-space: nowrap;
    transition: border-color .15s, box-shadow .15s, background .15s;
  }
  .filter-btn:hover { background: hsl(var(--muted)) }
  .filter-btn.open, .filter-btn.has {
    border-color: hsl(var(--primary) / .45);
    background: rgba(31,134,130,.06);
    color: hsl(var(--primary));
  }
  .filter-btn .caret { font-size: .65rem; opacity: .55 }
  .filter-btn .count {
    display: inline-flex; align-items: center; justify-content: center;
    min-width: 1.15rem; height: 1.15rem; padding: 0 .3rem; border-radius: 999px;
    background: var(--teal); color: #fff; font-size: .65rem; font-weight: 700;
  }
  .filter-menu {
    position: absolute; top: calc(100% + 4px); left: 0; z-index: 40;
    min-width: 260px; width: max-content; max-width: min(440px, 92vw); max-height: 320px; overflow: auto;
    background: #fff; border: 1px solid hsl(var(--border)); border-radius: .5rem;
    box-shadow: 0 10px 28px rgb(15 23 42 / .12); padding: .4rem;
  }
  .filter-dd[data-key="customer"] .filter-menu {
    min-width: 360px; max-width: min(520px, 94vw);
  }
  .filter-menu label {
    display: flex; align-items: flex-start; gap: .55rem;
    padding: .5rem .6rem; border-radius: .4rem; cursor: pointer;
    font-size: .8125rem; line-height: 1.35; color: hsl(var(--foreground));
  }
  .filter-menu label:hover { background: hsl(var(--muted)) }
  .filter-menu label span {
    flex: 1; min-width: 12rem; white-space: normal; overflow-wrap: break-word;
    word-break: normal; hyphens: none; text-transform: none;
  }
  .filter-menu input { margin-top: .2rem; accent-color: var(--teal); flex-shrink: 0; width: 15px; height: 15px }
  .filter-search {
    display: flex; align-items: center; gap: .5rem;
    flex: 1; min-width: 180px; height: 36px; padding: 0 .75rem;
    background: #fff; border: 1px solid hsl(var(--border)); border-radius: .5rem;
  }
  .filter-search input {
    flex: 1; min-width: 0; border: none; outline: none; background: transparent;
    font: inherit; color: hsl(var(--foreground)); height: 100%;
  }
  .filter-search svg { width: 16px; height: 16px; color: hsl(var(--muted-foreground)); flex-shrink: 0 }
  .filter-chips { display: flex; flex-wrap: wrap; gap: .4rem; align-items: center; min-height: 0 }
  .chip {
    display: inline-flex; align-items: center; gap: .35rem;
    height: 28px; padding: 0 .55rem 0 .7rem; border-radius: 999px;
    background: rgba(31,134,130,.1); color: hsl(200 60% 23%);
    font-size: .75rem; font-weight: 550; max-width: 280px;
  }
  .chip b { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap }
  .chip button {
    border: none; background: transparent; cursor: pointer; padding: 0;
    width: 16px; height: 16px; border-radius: 99px; color: inherit; opacity: .55;
    display: inline-flex; align-items: center; justify-content: center; font-size: .85rem; line-height: 1;
  }
  .chip button:hover { opacity: 1; background: rgba(0,0,0,.06) }
  .filter-meta { font-size: .75rem; color: hsl(var(--muted-foreground)); margin-left: auto }
  .filter-clear {
    border: none; background: transparent; color: hsl(var(--muted-foreground));
    font: 500 12px Inter, sans-serif; cursor: pointer; padding: .25rem .4rem; border-radius: .35rem;
  }
  .filter-clear:hover { color: hsl(var(--foreground)); background: hsl(var(--muted)) }
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
  .tbl-wrap { max-height: 640px; overflow: auto; border-radius: 0 0 .5rem .5rem }
  table.full {
    width: max-content; min-width: 100%; border-collapse: separate; border-spacing: 0;
    font-size: .78rem; table-layout: auto;
  }
  table.full th, table.full td {
    padding: .55rem .65rem; border-bottom: 1px solid hsl(var(--border));
    text-align: left; vertical-align: middle; background: #fff;
  }
  table.full th {
    position: sticky; top: 0; z-index: 3; background: hsl(var(--background));
    color: hsl(var(--muted-foreground)); font-size: .65rem; font-weight: 600;
    letter-spacing: .03em; text-transform: uppercase; white-space: nowrap;
    box-shadow: 0 1px 0 hsl(var(--border));
  }
  table.full td.r, table.full th.r { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap }
  /* Название объекта — широкая липкая колонка, без сжатия */
  table.dense th.obj, table.dense td.obj {
    position: sticky; left: 0; z-index: 2;
    min-width: 22rem; width: 22rem; max-width: 28rem;
    box-shadow: 1px 0 0 hsl(var(--border));
  }
  table.dense th.obj { z-index: 4; background: hsl(var(--background)) }
  table.dense td.obj { background: #fff }
  tr.clickable:hover td.obj, tr.row-active td.obj { background: hsl(var(--muted)) }
  tr.row-active td.obj { background: rgba(31,134,130,.08) }
  table.dense .obj-name {
    display: block; font-weight: 600; line-height: 1.35;
    white-space: normal; overflow-wrap: anywhere; hyphens: auto;
  }
  table.dense .obj-sub {
    display: block; margin-top: .2rem; font-size: .72rem;
    color: hsl(var(--muted-foreground)); line-height: 1.3;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
  }
  table.dense .obj-uin {
    display: block; margin-top: .15rem; font-size: .65rem;
    color: hsl(var(--muted-foreground)); font-variant-numeric: tabular-nums;
    white-space: nowrap;
  }
  table.dense .work {
    white-space: nowrap; max-width: 9rem;
    overflow: hidden; text-overflow: ellipsis;
  }
  table.dense .cust {
    white-space: nowrap; max-width: 11rem;
    overflow: hidden; text-overflow: ellipsis;
  }
  table.dense .exp-cell { white-space: nowrap; min-width: 7.5rem }
  table.dense .cell-stack { display: flex; flex-direction: column; align-items: flex-end; gap: 1px }
  table.dense .cell-stack.start { align-items: flex-start }
  table.dense .exp-date { font-size: .68rem; color: hsl(var(--muted-foreground)); font-variant-numeric: tabular-nums; white-space: nowrap }
  table.dense .amt { font-weight: 600 }
  table.dense .pct { font-size: .68rem; color: hsl(var(--muted-foreground)) }
  /* рост сметы — плохо (красное), снижение — хорошо (зелёное) */
  .amt.neg, .pct.neg { color: var(--err) }
  .amt.pos, .pct.pos { color: var(--ok) }
  tr.clickable { cursor: pointer }
  tr.clickable:hover { background: hsl(var(--muted)) }
  tr.row-active { background: rgba(31,134,130,.08) }
  .obj-name-row { display: flex; align-items: flex-start; gap: .35rem }
  .obj-name-row .obj-name { flex: 1; min-width: 0 }
  .copy-btn {
    flex-shrink: 0; margin-top: .1rem; width: 22px; height: 22px; padding: 0;
    border: 1px solid transparent; border-radius: .35rem; background: transparent;
    color: hsl(var(--muted-foreground)); cursor: pointer; font-size: .7rem; line-height: 1;
    opacity: 0; transition: opacity .15s, background .15s, border-color .15s;
  }
  tr:hover .copy-btn, tr.row-active .copy-btn { opacity: .85 }
  .copy-btn:hover { opacity: 1; background: hsl(var(--muted)); border-color: hsl(var(--border)); color: hsl(var(--foreground)) }
  .copy-btn.ok { color: var(--ok); opacity: 1 }
  .filter-tog {
    display: inline-flex; align-items: center; gap: .4rem; height: 36px; padding: 0 .85rem;
    border-radius: .5rem; border: 1px solid hsl(var(--border)); background: #fff;
    font: 500 13px Inter, sans-serif; color: hsl(var(--foreground)); cursor: pointer;
    user-select: none; white-space: nowrap;
  }
  .filter-tog input { accent-color: var(--teal); margin: 0 }
  .filter-tog.on { border-color: hsl(var(--primary) / .45); background: rgba(31,134,130,.06); color: hsl(var(--primary)) }
  .kpi .sub { margin: .15rem 0 0; font-size: .75rem; font-weight: 500; color: hsl(var(--muted-foreground)) }
  .filter-dd[data-key="contractor"] .filter-menu { min-width: 320px; max-width: min(480px, 94vw) }
  .pill {
    display: inline-flex; align-items: center; font-size: .7rem; font-weight: 600;
    padding: 2px 8px; border-radius: 999px;
  }
  .pill.ok { background: var(--ok-bg); color: #15803d }
  .pill.warn { background: var(--warn-bg); color: #b45309 }
  .pill.err { background: var(--err-bg); color: #b91c1c }
  .pill.muted { background: hsl(var(--muted)); color: hsl(var(--muted-foreground)) }
  .cell-stack.start .pill { align-self: flex-start }
  .muted { color: hsl(var(--muted-foreground)) }
  .quad-tools { display: flex; flex-wrap: wrap; gap: .75rem; align-items: center; margin-bottom: .75rem; font-size: .8125rem; color: hsl(var(--muted-foreground)) }
  .quad-tools input[type=number] { width: 64px; font: inherit; height: 32px; padding: 0 .5rem; border: 1px solid hsl(var(--border)); border-radius: .5rem; color: hsl(var(--foreground)) }
  .quad-cards { display: grid; grid-template-columns: repeat(2, 1fr); gap: .5rem; margin-bottom: .75rem }
  @media (min-width: 800px) { .quad-cards { grid-template-columns: repeat(4, 1fr) } }
  .quad-card { border: 1px solid hsl(var(--border)); border-left: 4px solid var(--q-color, #94a3b8); border-radius: .5rem; padding: .6rem .75rem; cursor: pointer; background: #fff }
  .quad-card.active { box-shadow: inset 0 0 0 1px var(--q-color); background: hsl(var(--muted)) }
  .quad-card .qn { font-size: 1.35rem; font-weight: 700; line-height: 1.1; font-variant-numeric: tabular-nums }
  .quad-card .ql { margin-top: .2rem; font-size: .75rem; font-weight: 600 }
  .quad-card .qd { margin-top: .3rem; font-size: .6875rem; font-weight: 400; color: hsl(var(--muted-foreground)); line-height: 1.35 }
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
        <p class="s" id="brandSub">Стройготовность и оплата · портфель объектов</p>
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
  <div class="filters">
    <div class="filter-row" id="filterRow">
      <div class="filter-dd" data-key="year">
        <button type="button" class="filter-btn" id="btnYear">Год открытия <span class="caret">▾</span></button>
        <div class="filter-menu" id="menuYear" hidden></div>
      </div>
      <div class="filter-dd" data-key="customer">
        <button type="button" class="filter-btn" id="btnCustomer">Заказчик <span class="caret">▾</span></button>
        <div class="filter-menu" id="menuCustomer" hidden></div>
      </div>
      <div class="filter-dd" data-key="contractor">
        <button type="button" class="filter-btn" id="btnContractor">Подрядчик <span class="caret">▾</span></button>
        <div class="filter-menu" id="menuContractor" hidden></div>
      </div>
      <div class="filter-dd" data-key="work">
        <button type="button" class="filter-btn" id="btnWork">Вид работ <span class="caret">▾</span></button>
        <div class="filter-menu" id="menuWork" hidden></div>
      </div>
      <div class="filter-dd" data-key="status">
        <button type="button" class="filter-btn" id="btnStatus">Статус <span class="caret">▾</span></button>
        <div class="filter-menu" id="menuStatus" hidden></div>
      </div>
      <label class="filter-tog" id="togSg" title="Скрыть объекты с нулевой стройготовностью">
        <input type="checkbox" id="fltSgPos"> СГ &gt; 0
      </label>
      <div class="filter-search">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>
        <input id="fltSearch" type="search" placeholder="Поиск по объекту, подрядчику, УИН…" autocomplete="off">
      </div>
    </div>
    <div class="filter-row">
      <div class="filter-chips" id="filterChips"></div>
      <span class="filter-meta" id="filterMeta"></span>
      <button type="button" class="filter-clear" id="filterClear" hidden>Сбросить</button>
    </div>
  </div>

  <div class="tabpanel stack" data-tab="summary">
  <div class="kpis">
    <div class="kpi"><p class="l">Объектов</p><p class="n" id="mzN">—</p></div>
    <div class="kpi" title="СГ и оплата отличаются не больше чем на 10 п.п.">
      <p class="l">В коридоре (±10 п.п.)</p>
      <p class="n ok" id="mzCorr">—</p>
      <p class="sub" id="mzCorrPct"></p>
    </div>
    <div class="kpi"><p class="l">Медиана СГ, %</p><p class="n" id="mzMedSg">—</p></div>
    <div class="kpi"><p class="l">Кредитует подрядчик</p><p class="n warn" id="mzCredit">—</p></div>
  </div>

  <div class="card">
    <div class="card-hd">
      <div>
        <h2>Бюджет по источникам</h2>
        <p id="mzBudgetNote">Лимиты — предусмотрено. Доведено — уже пришло на объекты (в выбранном срезе).</p>
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
          <th class="obj">Объект</th>
          <th>Вид</th>
          <th class="r">Год</th>
          <th>Заказчик</th>
          <th class="r" title="млн руб.">Лимит, млн</th>
          <th class="r" title="млн руб.">Контракт, млн</th>
          <th class="r">СГ</th><th class="r">Оплата</th>
          <th class="r" title="Согласованная смета после экспертизы, млн руб.">После эксп.</th>
          <th class="r" title="Изменение сметы, млн руб.">Δ сметы</th>
          <th title="ТЧ — техчасть, СД — смета">Экспертиза</th>
          <th>Статус</th>
        </tr></thead>
        <tbody id="mzTbl"></tbody>
      </table>
    </div>
  </div>
  </div>

  <div class="tabpanel stack" data-tab="contractors" hidden>
  <div class="kpis" id="contrKpis">
    <div class="kpi" title="Число подрядчиков в текущем фильтре (без «Без подрядчика»)">
      <p class="l">Подрядчиков</p><p class="n" id="cN">—</p>
    </div>
    <div class="kpi" title="Объекты, где СГ выше оплаты больше чем на 10 п.п. — подрядчик фактически кредитует стройку">
      <p class="l">Кредитуют объекты</p><p class="n warn" id="cCredit">—</p>
    </div>
    <div class="kpi" title="Подрядчики в квадранте «Аванс»: средняя оплата ≥ порога, средняя СГ ниже порога">
      <p class="l">Аванс впереди работ</p><p class="n" id="cAdvance">—</p>
    </div>
    <div class="kpi" title="Сумма лимитов объектов у подрядчиков в срезе">
      <p class="l">Лимит в срезе</p><p class="n" id="cLimit">—</p>
    </div>
  </div>

  <div class="card">
    <div class="card-hd">
      <div>
        <h2>Карта подрядчиков</h2>
        <p>Точка — подрядчик: средние СГ и оплата (взвешены по лимиту), размер — лимит. Крест — пороги: делят карту на 4 группы. Клик по карточке или точке — фильтр / объекты.</p>
      </div>
    </div>
    <div class="card-bd">
      <div class="quad-tools">
        <label title="Вертикальная граница на карте">Порог СГ, % <input type="number" id="qSg" min="0" max="100" step="5" value="50"></label>
        <label title="Горизонтальная граница на карте">Порог оплаты, % <input type="number" id="qPay" min="0" max="100" step="5" value="50"></label>
      </div>
      <div class="quad-cards" id="quadCards"></div>
      <div class="chart" style="height:400px"><canvas id="cQuad"></canvas></div>
    </div>
  </div>

  <div class="card">
    <div class="card-hd">
      <div>
        <h2>Рейтинг подрядчиков</h2>
        <p>Сортировка по лимиту. СГ и оплата — средние по портфелю, взвешенные по лимиту. Клик — объекты подрядчика</p>
      </div>
    </div>
    <div class="tbl-wrap">
      <table class="full dense">
        <thead><tr>
          <th>Подрядчик</th>
          <th class="r">Объектов</th>
          <th class="r" title="Сумма лимитов объектов, млн руб.">Лимит, млн</th>
          <th class="r" title="Средняя стройготовность, взвешенная по лимиту">СГ ср., %</th>
          <th class="r" title="Средний % оплаты, взвешенный по лимиту">Оплата ср., %</th>
          <th class="r" title="Число объектов, где СГ заметно выше оплаты (подрядчик кредитует)">Кредитует</th>
          <th class="r" title="Сумма рублёвого разрыва на объектах с кредитованием, млн руб.">Разрыв, млн</th>
          <th title="Квадрант относительно порогов СГ и оплаты">Группа</th>
        </tr></thead>
        <tbody id="contrTbl"></tbody>
      </table>
    </div>
  </div>
  </div>

  <div class="tabpanel stack" data-tab="objects" hidden>
  <div class="card">
    <div class="card-hd">
      <div>
        <h2>Матрица риска: готовность vs оплата</h2>
        <p id="mzMatrixNote">Снизу оплата, слева СГ. Зелёная полоса — коридор ±10 п.п. Клик по точке — только этот объект; клик по фону — снять фокус.</p>
      </div>
    </div>
    <div class="card-bd">
      <div class="chart tall"><canvas id="cMzMatrix"></canvas></div>
    </div>
  </div>

  <div class="card">
    <div class="card-hd">
      <div>
        <h2>Объекты</h2>
        <p>По текущему фильтру. Клик — траектория ниже</p>
      </div>
    </div>
    <div class="tbl-wrap">
      <table class="full dense">
        <thead><tr>
          <th class="obj">Объект</th>
          <th>Вид</th>
          <th class="r">Год</th>
          <th>Заказчик</th>
          <th class="r" title="млн руб.">Лимит, млн</th>
          <th class="r" title="млн руб.">Контракт, млн</th>
          <th class="r">СГ</th><th class="r">Оплата</th>
          <th class="r" title="После экспертизы, млн руб.">После эксп.</th>
          <th class="r" title="Изменение сметы, млн руб.">Δ сметы</th>
          <th>Экспертиза</th>
          <th>Статус</th>
        </tr></thead>
        <tbody id="mzObjTbl"></tbody>
      </table>
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
let mzYears = ['2026'], mzCustomers = [], mzContractors = [], mzWorks = [], mzStatuses = [];
let mzSgPos = false;
let mzSearch = '', mzUin = null, mzFocusUin = null;
let chartMzMatrix = null, chartMzTraj = null, chartQuad = null;

const WORK_OPTS = [
  { v: 'капремонт', l: 'Капремонт' },
  { v: 'строительство', l: 'Строительство' },
  { v: 'реконструкция', l: 'Реконструкция' },
  { v: 'прочее', l: 'Прочее' },
];
const STATUS_OPTS = [
  { v: 'credit', l: 'Кредитует' },
  { v: 'over', l: 'Избыток оплаты' },
  { v: 'balanced', l: 'Баланс / коридор' },
  { v: 'exp', l: 'Есть экспертиза' },
  { v: 'noexp', l: 'Без экспертизы' },
];
const EXP_KIND = {
  both: { l: 'ТЧ+СД', tip: 'Техчасть и смета', c: 'ok' },
  tch:  { l: 'ТЧ', tip: 'Только техчасть', c: 'ok' },
  sd:   { l: 'СД', tip: 'Только смета', c: 'ok' },
};

function statusMatch(o, st) {
  if (st === 'credit') return o.money_status === 'credit';
  if (st === 'over') return o.money_status === 'over';
  if (st === 'balanced') return o.money_status === 'balanced';
  if (st === 'exp') return !!(o.exp_done || o.exp_agreed_mln != null || o.exp_kind);
  if (st === 'noexp') return !(o.exp_done || o.exp_agreed_mln != null || o.exp_kind);
  return true;
}
function mzPass(o) {
  if (mzYears.length && !mzYears.includes(String(o.open_year || ''))) return false;
  if (mzCustomers.length && !mzCustomers.includes(o.customer || '')) return false;
  if (mzContractors.length && !mzContractors.includes(o.contractor || '')) return false;
  if (mzWorks.length && !mzWorks.includes(o.bucket)) return false;
  if (mzStatuses.length && !mzStatuses.some(st => statusMatch(o, st))) return false;
  if (mzSgPos && !(o.sg > 0)) return false;
  if (mzSearch) {
    const q = mzSearch.toLowerCase();
    const blob = [o.name, o.full, o.contractor, o.customer, o.ownership, o.work, o.uin, o.open_year]
      .filter(Boolean).join(' ').toLowerCase();
    if (!blob.includes(q)) return false;
  }
  return true;
}
function copyText(text, btn) {
  const t = String(text || '').trim();
  if (!t) return;
  const done = () => {
    if (!btn) return;
    btn.classList.add('ok');
    btn.textContent = '✓';
    setTimeout(() => { btn.classList.remove('ok'); btn.textContent = '⧉'; }, 900);
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(t).then(done).catch(() => {});
  } else {
    const ta = document.createElement('textarea');
    ta.value = t; document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); done(); } catch (e) {}
    ta.remove();
  }
}
function expCellHtml(o) {
  const kind = EXP_KIND[o.exp_kind];
  const tip = [
    kind ? kind.tip : '',
    o.exp_done ? 'завершена ' + o.exp_done : '',
    o.exp_start ? 'старт ' + o.exp_start : '',
    o.exp_in ? 'заход ' + o.exp_in : '',
  ].filter(Boolean).join(' · ');
  const bits = [];
  if (kind) bits.push(`<span class="pill ${kind.c}">${kind.l}</span>`);
  if (o.exp_done) {
    bits.push(`<span class="exp-date">${esc(o.exp_done)}</span>`);
  } else if (o.exp_in || o.exp_start) {
    bits.push(`<span class="pill warn">в эксп.</span>`);
    bits.push(`<span class="exp-date">${esc(o.exp_start || o.exp_in)}</span>`);
  } else if (o.entered_exp) {
    bits.push(`<span class="pill warn">заявка</span>`);
  } else if (!kind) {
    return '<span class="muted">—</span>';
  }
  return `<div class="cell-stack start exp-cell" title="${esc(tip)}">${bits.join('')}</div>`;
}
function toggleIn(arr, v) {
  const i = arr.indexOf(v);
  if (i >= 0) arr.splice(i, 1);
  else arr.push(v);
  return arr;
}
function mzFiltered() { return (MZ.objects || []).filter(mzPass); }

function mzSelectObject(uin, { focus = false, goObjects = false } = {}) {
  if (!uin) return;
  mzUin = uin;
  if (focus) mzFocusUin = uin;
  const sel = document.getElementById('selMz');
  if (sel) sel.value = mzUin;
  if (goObjects) window.showTab('objects');
  else {
    drawMzTraj(mzUin);
    renderMzMatrix();
    renderMzTbl();
  }
}
function clearMzFocus() {
  if (!mzFocusUin) return;
  mzFocusUin = null;
  renderMzMatrix();
  renderMzTbl();
  if (window.renderFilterChrome) window.renderFilterChrome();
}

function renderMzMatrix() {
  let rows = mzFiltered().filter(o => o.sg != null && o.pct != null);
  const note = document.getElementById('mzMatrixNote');
  if (mzFocusUin) {
    const focused = rows.filter(o => o.uin === mzFocusUin);
    if (focused.length) {
      rows = focused;
      const o = focused[0];
      if (note) note.textContent = `Фокус: ${o.full || o.name}. Клик по фону матрицы — снять.`;
    } else if (note) {
      note.textContent = 'Выбранный объект вне текущего фильтра.';
    }
  } else if (note) {
    note.textContent = 'Снизу оплата, слева СГ. Зелёная полоса — коридор ±10 п.п. Клик по точке — только этот объект; клик по фону — снять фокус.';
  }
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
          return `${d.name}: СГ ${d.y}% · оплата ${d.x}% · лимит ${mlnFmt(d.limit)} · контракт ${mlnFmt(d.contract)} млн`;
        }}},
      },
      scales: {
        x: { min:0, max:100, title:{ display:true, text:'Оплата (от лимита), %', font:{ size:13, weight:'600' } } },
        y: { min:0, max:100, title:{ display:true, text:'Стройготовность (факт), %', font:{ size:13, weight:'600' } } },
      },
      onClick(evt, els) {
        if (!els.length) {
          clearMzFocus();
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

/** Сумма в млн руб., всегда с 2 знаками: 1,00 / 70 402,00 */
function mlnFmt(v, { signed = false } = {}) {
  if (v == null || !isFinite(v)) return '—';
  const n = Number(v);
  const sign = signed && n > 0 ? '+' : '';
  return sign + n.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function rub(v) {
  const s = mlnFmt(v);
  return s === '—' ? s : s + ' млн';
}
function esc(s) {
  return String(s ?? '').replace(/[&<>"]/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' }[c]));
}
function median(xs) {
  if (!xs.length) return null;
  const a = xs.slice().sort((x, y) => x - y);
  return Math.round(a[Math.floor(a.length / 2)] * 10) / 10;
}
function fundTotals(rows) {
  const acc = {};
  rows.forEach(o => (o.funds || []).forEach(f => {
    const s = acc[f.k] || { plan: 0, fact: 0, n: 0 };
    s.plan += f.plan || 0; s.fact += f.fact || 0; s.n += 1;
    acc[f.k] = s;
  }));
  return acc;
}
function renderSummaryKpis() {
  const rows = mzFiltered();
  const withBoth = rows.filter(o => o.sg != null && o.pct != null);
  const sgs = rows.map(o => o.sg).filter(v => v != null);
  const nCredit = rows.filter(o => o.money_status === 'credit').length;
  const nCorr = withBoth.filter(o => o.money_status === 'balanced').length;
  const medSg = median(sgs);
  document.getElementById('mzN').textContent = rows.length.toLocaleString('ru-RU');
  document.getElementById('mzCorr').textContent = nCorr.toLocaleString('ru-RU');
  const pctEl = document.getElementById('mzCorrPct');
  if (pctEl) {
    pctEl.textContent = withBoth.length
      ? (Math.round(1000 * nCorr / withBoth.length) / 10) + '% от среза с СГ и оплатой'
      : '';
  }
  document.getElementById('mzMedSg').textContent = medSg != null ? medSg : '—';
  document.getElementById('mzCredit').textContent = nCredit.toLocaleString('ru-RU');
}
function renderPortfolio() {
  const rows = mzFiltered();
  const acc = fundTotals(rows);
  const fed = acc.fed || { plan: 0, fact: 0, n: 0 };
  const mo = acc.mo || { plan: 0, fact: 0, n: 0 };
  const mun = acc.mun || { plan: 0, fact: 0, n: 0 };
  const ibk = acc.ibk || { plan: 0, fact: 0, n: 0 };
  const kik = acc.kik || { plan: 0, fact: 0, n: 0 };
  const infra = rows.filter(o => (o.funds || []).some(f => f.k === 'ibk' || f.k === 'kik'));
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
  stable:  {
    label: 'Устойчивые',
    color: '#16a34a',
    hint: 'СГ и оплата ≥ порогов — работы и деньги идут вместе',
  },
  cashgap: {
    label: 'Кредитует подрядчик',
    color: '#d97706',
    hint: 'СГ ≥ порога, оплата ниже — стройка впереди денег',
  },
  advance: {
    label: 'Аванс впереди работ',
    color: '#7c3aed',
    hint: 'Оплата ≥ порога, СГ ниже — деньги впереди готовности',
  },
  early:   {
    label: 'Ранняя стадия',
    color: '#0B6CB8',
    hint: 'СГ и оплата ниже порогов — старт или пробуксовка',
  },
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
  mzFiltered().forEach(o => {
    const name = o.contractor || 'Без подрядчика';
    const e = byC.get(name) || {
      contractor: name, n: 0, limit: 0, wSg: 0, wPct: 0, w: 0,
      n_credit: 0, gap_rub: 0,
    };
    const w = o.limit || o.contract_value || 0;
    e.n += 1; e.limit += w;
    if (o.money_status === 'credit') {
      e.n_credit += 1;
      e.gap_rub += Math.abs(o.gap_rub || 0);
    }
    if (o.sg != null && o.pct != null && w > 0) {
      e.wSg += o.sg * w; e.wPct += o.pct * w; e.w += w;
    }
    byC.set(name, e);
  });
  return [...byC.values()].map(e => {
    const avg_sg = e.w ? Math.round(e.wSg / e.w * 10) / 10 : null;
    const avg_pct = e.w ? Math.round(e.wPct / e.w * 10) / 10 : null;
    return Object.assign(e, {
      avg_sg, avg_pct,
      gap_rub: Math.round(e.gap_rub * 10) / 10,
      quad: quadrantOf(avg_sg, avg_pct),
    });
  }).sort((a, b) => (b.limit || 0) - (a.limit || 0));
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

function openContractor(name) {
  if (!name || name === 'Без подрядчика') return;
  mzContractors = [name];
  mzFocusUin = null;
  mzSearch = '';
  const input = document.getElementById('fltSearch');
  if (input) input.value = '';
  window.showTab('objects');
  if (window.refreshFilters) window.refreshFilters();
}
function renderContrTbl(rows) {
  const el = document.getElementById('contrTbl');
  if (!el) return;
  const list = rows.filter(r => r.contractor !== 'Без подрядчика');
  el.innerHTML = list.map(r => {
    const m = QUAD_META[r.quad] || { label: '—', color: '#94a3b8' };
    return `<tr class="clickable" data-contractor="${esc(r.contractor)}">` +
      `<td class="obj"><span class="obj-name">${esc(r.contractor)}</span></td>` +
      `<td class="r">${r.n}</td>` +
      `<td class="r">${r.limit ? mlnFmt(r.limit) : '—'}</td>` +
      `<td class="r">${r.avg_sg ?? '—'}</td>` +
      `<td class="r">${r.avg_pct ?? '—'}</td>` +
      `<td class="r">${r.n_credit || 0}</td>` +
      `<td class="r">${r.gap_rub ? mlnFmt(r.gap_rub) : '—'}</td>` +
      `<td><span class="pill" style="background:${m.color}22;color:${m.color}">${esc(m.label)}</span></td>` +
      `</tr>`;
  }).join('') || '<tr><td colspan="8" class="note">Нет подрядчиков по фильтру</td></tr>';
  el.querySelectorAll('tr.clickable').forEach(tr => {
    tr.onclick = () => openContractor(tr.dataset.contractor);
  });
}
function renderQuad() {
  const rows = contractorPoints();
  const named = rows.filter(r => r.contractor !== 'Без подрядчика');
  const counts = { stable: 0, cashgap: 0, advance: 0, early: 0 };
  named.forEach(r => { if (r.quad) counts[r.quad]++; });
  const nCreditObj = mzFiltered().filter(o => o.money_status === 'credit').length;
  const limitSum = named.reduce((s, r) => s + (r.limit || 0), 0);
  document.getElementById('cN').textContent = named.length.toLocaleString('ru-RU');
  document.getElementById('cCredit').textContent = nCreditObj.toLocaleString('ru-RU');
  document.getElementById('cAdvance').textContent = (counts.advance || 0).toLocaleString('ru-RU');
  document.getElementById('cLimit').textContent = rub(limitSum);

  document.getElementById('quadCards').innerHTML = Object.keys(QUAD_META).map(k => {
    const m = QUAD_META[k];
    const active = quadState.active === k ? ' active' : '';
    return `<div class="quad-card${active}" data-quad="${k}" style="--q-color:${m.color}" title="${esc(m.hint)}">` +
      `<div class="qn">${counts[k]}</div>` +
      `<div class="ql">${m.label}</div>` +
      `<div class="qd">${esc(m.hint)}</div></div>`;
  }).join('');
  document.querySelectorAll('#quadCards .quad-card').forEach(el => {
    el.onclick = () => {
      quadState.active = quadState.active === el.dataset.quad ? null : el.dataset.quad;
      renderQuad();
    };
  });
  renderContrTbl(quadState.active ? named.filter(r => r.quad === quadState.active) : named);

  const visible = quadState.active ? named.filter(r => r.quad === quadState.active) : named;
  const datasets = Object.keys(QUAD_META).map(k => {
    const m = QUAD_META[k];
    const data = visible.filter(r => r.quad === k && r.avg_sg != null && r.avg_pct != null).map(r => ({
      x: r.avg_sg, y: r.avg_pct,
      r: Math.max(5, Math.min(18, 3 + Math.sqrt(Math.max(r.limit, 0)) * 0.22)),
      contractor: r.contractor, limit: r.limit, n: r.n, n_credit: r.n_credit, gap_rub: r.gap_rub,
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
              return [
                d.contractor,
                `СГ ${d.x}% · оплата ${d.y}%`,
                `${d.n} объектов · лимит ${rub(d.limit)}`,
                d.n_credit ? `кредитует ${d.n_credit} · разрыв ${rub(d.gap_rub)}` : 'без кредитования',
              ];
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
        if (!d || !d.contractor) return;
        openContractor(d.contractor);
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
      ? ` (${mlnFmt(o.exp_delta_mln, { signed: true })} млн` + (o.exp_overrun!=null ? `, ${o.exp_overrun>0?'+':''}${o.exp_overrun}%` : '') + ')'
      : '';
    bits.push(`смета после эксп. ${mlnFmt(o.exp_agreed_mln)} млн${dlt}`);
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

const WORK_SHORT = {
  капремонт: 'Капремонт',
  строительство: 'Строительство',
  реконструкция: 'Реконструкция',
  прочее: 'Прочее',
};
function objectRowHtml(o, { stayOnObjects = false } = {}) {
  const active = o.uin === mzUin || o.uin === mzFocusUin ? ' row-active' : '';
  // рост сметы (Δ>0) — красный, снижение — зелёный
  const ovCls = o.exp_overrun > 0 ? 'neg' : (o.exp_overrun < 0 ? 'pos' : '');
  const deltaCell = o.exp_delta_mln != null
    ? `<div class="cell-stack end"><span class="amt ${ovCls}">${mlnFmt(o.exp_delta_mln, { signed: true })}</span>` +
      (o.exp_overrun!=null ? `<span class="pct ${ovCls}">${o.exp_overrun>0?'+':''}${o.exp_overrun}%</span>` : '') +
      `</div>`
    : '—';
  const afterExp = o.exp_agreed_mln != null
    ? `<div class="cell-stack end"><span class="amt">${mlnFmt(o.exp_agreed_mln)}</span>` +
      (o.exp_plan_mln != null ? `<span class="pct">план ${mlnFmt(o.exp_plan_mln)}</span>` : '') +
      `</div>`
    : '—';
  const workShort = WORK_SHORT[o.bucket] || (o.work || '—');
  const tipBits = [
    o.full || o.name,
    o.work,
    o.customer,
    o.ownership,
    o.contractor ? 'подрядчик: ' + o.contractor : '',
  ].filter(Boolean).join('\n');
  const copyPayload = [(o.full || o.name || ''), o.uin].filter(Boolean).join(' · ');
  return `<tr class="clickable${active}" data-uin="${o.uin}" data-stay="${stayOnObjects ? '1' : '0'}">` +
    `<td class="obj" title="${esc(tipBits)}">` +
      `<div class="obj-name-row">` +
        `<span class="obj-name">${esc(o.full || o.name)}</span>` +
        `<button type="button" class="copy-btn" data-copy="${esc(copyPayload)}" title="Скопировать название и УИН">⧉</button>` +
      `</div>` +
      (o.contractor ? `<span class="obj-sub" title="${esc(o.contractor)}">${esc(o.contractor)}</span>` : '') +
      `<span class="obj-uin">${esc(o.uin)}</span>` +
    `</td>` +
    `<td class="work" title="${esc(o.work || '')}">${esc(workShort)}</td>` +
    `<td class="r">${o.open_year??'—'}</td>` +
    `<td class="cust" title="${esc(o.customer || '')}">${esc(o.customer || '—')}</td>` +
    `<td class="r">${mlnFmt(o.limit)}</td>` +
    `<td class="r">${mlnFmt(o.contract_value)}</td>` +
    `<td class="r">${o.sg??'—'}</td><td class="r">${o.pct??'—'}</td>` +
    `<td class="r">${afterExp}</td>` +
    `<td class="r">${deltaCell}</td>` +
    `<td class="exp-cell">${expCellHtml(o)}</td>` +
    `<td><span class="pill ${STATUS_PILL[o.money_status]||''}">${STATUS_SHORT[o.money_status]||'—'}</span></td></tr>`;
}
function bindObjectRows(rootSel) {
  document.querySelectorAll(rootSel + ' tr.clickable').forEach(tr => {
    tr.onclick = () => {
      mzSelectObject(tr.dataset.uin, {
        focus: true,
        goObjects: tr.dataset.stay !== '1',
      });
      if (tr.dataset.stay === '1') {
        document.getElementById('cMzTraj')?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    };
  });
  document.querySelectorAll(rootSel + ' .copy-btn').forEach(btn => {
    btn.onclick = e => {
      e.stopPropagation();
      copyText(btn.getAttribute('data-copy'), btn);
    };
  });
}
function renderMzTbl() {
  let filtered = mzFiltered().slice().sort((a,b) => (b.sg||0)-(a.sg||0));
  if (mzFocusUin) {
    const one = filtered.filter(o => o.uin === mzFocusUin);
    if (one.length) filtered = one;
  }
  const wrap = document.getElementById('mzTbl')?.closest('.tbl-wrap');
  document.getElementById('mzTbl').innerHTML =
    filtered.map(o => objectRowHtml(o)).join('') ||
    '<tr><td colspan="12" class="note">Нет объектов по фильтру</td></tr>';
  if (wrap) wrap.scrollTop = 0;
  bindObjectRows('#mzTbl');

  const objTbl = document.getElementById('mzObjTbl');
  if (objTbl) {
    const wrap2 = objTbl.closest('.tbl-wrap');
    objTbl.innerHTML =
      filtered.map(o => objectRowHtml(o, { stayOnObjects: true })).join('') ||
      '<tr><td colspan="12" class="note">Нет объектов по фильтру</td></tr>';
    if (wrap2) wrap2.scrollTop = 0;
    bindObjectRows('#mzObjTbl');
  }
  if (window.renderFilterChrome) window.renderFilterChrome();
}

(function init() {
  function showTab(tab) {
    document.querySelectorAll('.nav-tab').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
    document.querySelectorAll('.tabpanel').forEach(p => { p.hidden = p.dataset.tab !== tab; });
    const panel = document.querySelector('.tabpanel[data-tab="' + tab + '"]');
    if (panel) panel.getBoundingClientRect();
    if (tab === 'summary') { renderSummaryKpis(); renderPortfolio(); renderMzTbl(); }
    if (tab === 'contractors') renderQuad();
    if (tab === 'objects') {
      renderMzMatrix();
      renderMzTbl();
      if (mzUin) drawMzTraj(mzUin);
    }
  }
  window.showTab = showTab;
  document.querySelectorAll('.nav-tab').forEach(b => { b.onclick = () => showTab(b.dataset.tab); });

  const m = MZ.meta || {};
  const asOf = m.as_of ? m.as_of.split('-').reverse().join('.') : '';
  const sub = document.getElementById('brandSub');
  if (sub) {
    sub.textContent = (asOf ? 'Данные на ' + asOf + ' · ' : '') + 'СГ и оплата по портфелю Минстроя';
  }
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

  const years = [...new Set(MZ.objects.map(o => o.open_year).filter(Boolean))].sort();

  function prettyCust(s) {
    if (!s) return '';
    const u = String(s).toUpperCase().replace(/Ё/g, 'Е');
    if (u.includes('ДИРЕКЦИЯ ЗАКАЗЧИКА КАПИТАЛЬНОГО') || /(^|\s)ДЗКС(\s|$)/.test(u)) return 'ГКУ МО ДЗКС';
    if (u.includes('ЕДИНОГО ЗАКАЗЧИКА') || /(^|\s)ДЕЗ(\s|$)/.test(u)) return 'ГКУ МО ДЕЗ';
    if (s === 'ГКУ МО ДЗКС' || s === 'ГКУ МО ДЕЗ') return s;
    // уже короткое «Администрация …» / МКУ / МБУ — как есть (кавычки у администраций убираем)
    if (/^Администрация/.test(s)) return String(s).replace(/[«»"]/g, '').replace(/\s+/g, ' ').trim();
    if (/^МКУ |^МБУ |^ГБУЗ |^Упр\. /.test(s)) return s;
    // Адм. … городского округа X
    let m = String(s).match(/(?:Адм\.|Администрация)\s+(?:городского\s+округа\s+)?(.+?)(?:\s+Московской|\s*,|$)/i);
    if (m) {
      let place = m[1].replace(/[A-ZА-ЯЁ]{2,}/g, w => w[0] + w.slice(1).toLowerCase()).trim();
      place = place.replace(/\s+городского\s+округа$/i, '').replace(/[«»"]/g, '').replace(/ого$/i, 'ий');
      return place ? `Администрация ${place}` : 'Администрация';
    }
    return String(s)
      .replace(/[A-ZА-ЯЁ]{2,}/g, w => w[0] + w.slice(1).toLowerCase())
      .replace(/\s+/g, ' ')
      .trim();
  }
  function customersForSelectedYears() {
    const rows = MZ.objects.filter(o => {
      if (!o.customer) return false;
      if (!mzYears.length) return true;
      return mzYears.includes(String(o.open_year || ''));
    });
    return [...new Set(rows.map(o => o.customer))]
      .sort((a, b) => prettyCust(a).localeCompare(prettyCust(b), 'ru'));
  }
  function contractorsForSelectedYears() {
    const rows = MZ.objects.filter(o => {
      if (!o.contractor) return false;
      if (!mzYears.length) return true;
      return mzYears.includes(String(o.open_year || ''));
    });
    return [...new Set(rows.map(o => o.contractor))]
      .sort((a, b) => a.localeCompare(b, 'ru'));
  }

  function fillCheckMenu(menu, opts, selected, onToggle) {
    if (!opts.length) {
      menu.innerHTML = '<div style="padding:.65rem .7rem;font-size:.78rem;color:hsl(var(--muted-foreground))">Нет вариантов</div>';
      return;
    }
    menu.innerHTML = opts.map(o => {
      const checked = selected.includes(o.v) ? ' checked' : '';
      return `<label><input type="checkbox" value="${esc(o.v)}"${checked}><span title="${esc(o.l)}">${esc(o.l)}</span></label>`;
    }).join('');
    menu.querySelectorAll('input').forEach(inp => {
      inp.onchange = () => { onToggle(inp.value); refreshFilters(); };
    });
  }
  function syncMenus() {
    const custOpts = customersForSelectedYears();
    const ctrOpts = contractorsForSelectedYears();
    const availC = new Set(custOpts);
    const availT = new Set(ctrOpts);
    mzCustomers = mzCustomers.filter(c => availC.has(c));
    mzContractors = mzContractors.filter(c => availT.has(c));
    fillCheckMenu(
      document.getElementById('menuYear'),
      years.map(y => ({ v: y, l: y })),
      mzYears,
      v => { mzYears = toggleIn(mzYears.slice(), v); }
    );
    fillCheckMenu(
      document.getElementById('menuCustomer'),
      custOpts.map(c => ({ v: c, l: prettyCust(c) })),
      mzCustomers,
      v => { mzCustomers = toggleIn(mzCustomers.slice(), v); }
    );
    fillCheckMenu(
      document.getElementById('menuContractor'),
      ctrOpts.map(c => ({ v: c, l: c })),
      mzContractors,
      v => { mzContractors = toggleIn(mzContractors.slice(), v); }
    );
    fillCheckMenu(
      document.getElementById('menuWork'),
      WORK_OPTS,
      mzWorks,
      v => { mzWorks = toggleIn(mzWorks.slice(), v); }
    );
    fillCheckMenu(
      document.getElementById('menuStatus'),
      STATUS_OPTS,
      mzStatuses,
      v => { mzStatuses = toggleIn(mzStatuses.slice(), v); }
    );
    const sgInp = document.getElementById('fltSgPos');
    const sgLab = document.getElementById('togSg');
    if (sgInp) sgInp.checked = mzSgPos;
    if (sgLab) sgLab.classList.toggle('on', mzSgPos);
  }
  window.renderFilterChrome = function renderFilterChrome() {
    const map = [
      ['btnYear', mzYears.length],
      ['btnCustomer', mzCustomers.length],
      ['btnContractor', mzContractors.length],
      ['btnWork', mzWorks.length],
      ['btnStatus', mzStatuses.length],
    ];
    map.forEach(([id, n]) => {
      const btn = document.getElementById(id);
      if (!btn) return;
      btn.classList.toggle('has', n > 0);
      let count = btn.querySelector('.count');
      if (n > 0) {
        if (!count) {
          count = document.createElement('span');
          count.className = 'count';
          btn.insertBefore(count, btn.querySelector('.caret'));
        }
        count.textContent = String(n);
      } else if (count) {
        count.remove();
      }
    });
    const chips = [];
    mzYears.forEach(y => chips.push({ key: 'year', v: y, l: 'Год ' + y }));
    mzCustomers.forEach(c => chips.push({ key: 'customer', v: c, l: prettyCust(c) }));
    mzContractors.forEach(c => chips.push({ key: 'contractor', v: c, l: c }));
    mzWorks.forEach(w => {
      const opt = WORK_OPTS.find(x => x.v === w);
      chips.push({ key: 'work', v: w, l: opt ? opt.l : w });
    });
    mzStatuses.forEach(s => {
      const opt = STATUS_OPTS.find(x => x.v === s);
      chips.push({ key: 'status', v: s, l: opt ? opt.l : s });
    });
    if (mzSgPos) chips.push({ key: 'sgpos', v: '1', l: 'СГ > 0' });
    if (mzFocusUin) {
      const fo = MZ.objects.find(x => x.uin === mzFocusUin);
      chips.push({ key: 'focus', v: mzFocusUin, l: 'Объект: ' + (fo ? (fo.name || fo.uin) : mzFocusUin) });
    }
    if (mzSearch) chips.push({ key: 'search', v: mzSearch, l: '«' + mzSearch + '»' });
    const box = document.getElementById('filterChips');
    box.innerHTML = chips.map((c, i) =>
      `<span class="chip"><b title="${esc(c.l)}">${esc(c.l)}</b>` +
      `<button type="button" data-i="${i}" aria-label="Убрать">×</button></span>`
    ).join('');
    box.querySelectorAll('button').forEach(btn => {
      btn.onclick = () => {
        const c = chips[+btn.dataset.i];
        if (!c) return;
        if (c.key === 'year') mzYears = mzYears.filter(x => x !== c.v);
        if (c.key === 'customer') mzCustomers = mzCustomers.filter(x => x !== c.v);
        if (c.key === 'contractor') mzContractors = mzContractors.filter(x => x !== c.v);
        if (c.key === 'work') mzWorks = mzWorks.filter(x => x !== c.v);
        if (c.key === 'status') mzStatuses = mzStatuses.filter(x => x !== c.v);
        if (c.key === 'sgpos') {
          mzSgPos = false;
          const sgInp = document.getElementById('fltSgPos');
          if (sgInp) sgInp.checked = false;
        }
        if (c.key === 'focus') { mzFocusUin = null; }
        if (c.key === 'search') {
          mzSearch = '';
          document.getElementById('fltSearch').value = '';
        }
        refreshFilters();
      };
    });
    const n = mzFiltered().length;
    document.getElementById('filterMeta').textContent =
      n.toLocaleString('ru-RU') + ' из ' + MZ.objects.length.toLocaleString('ru-RU');
    const clear = document.getElementById('filterClear');
    const dirty = mzYears.length || mzCustomers.length || mzContractors.length || mzWorks.length
      || mzStatuses.length || mzSearch || mzSgPos || mzFocusUin;
    clear.hidden = !dirty;
  };
  function closeMenus(except) {
    document.querySelectorAll('.filter-menu').forEach(m => {
      if (m !== except) m.hidden = true;
    });
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('open'));
  }
  function wireDropdown(btnId, menuId) {
    const btn = document.getElementById(btnId);
    const menu = document.getElementById(menuId);
    btn.onclick = e => {
      e.stopPropagation();
      const willOpen = menu.hidden;
      closeMenus();
      if (willOpen) {
        menu.hidden = false;
        btn.classList.add('open');
      }
    };
  }
  wireDropdown('btnYear', 'menuYear');
  wireDropdown('btnCustomer', 'menuCustomer');
  wireDropdown('btnContractor', 'menuContractor');
  wireDropdown('btnWork', 'menuWork');
  wireDropdown('btnStatus', 'menuStatus');
  document.addEventListener('click', () => closeMenus());
  document.querySelectorAll('.filter-menu').forEach(m => {
    m.addEventListener('click', e => e.stopPropagation());
  });
  document.getElementById('fltSgPos').onchange = e => {
    mzSgPos = !!e.target.checked;
    refreshFilters();
  };

  function refreshFilters() {
    // фокус объекта не сбрасываем при снятии чипа focus — только при смене остальных фильтров
    syncMenus();
    renderSummaryKpis();
    renderPortfolio();
    renderMzMatrix();
    renderMzTbl();
    const activeTab = document.querySelector('.nav-tab.active')?.dataset.tab;
    if (activeTab === 'contractors') renderQuad();
  }
  window.refreshFilters = refreshFilters;
  document.getElementById('filterClear').onclick = () => {
    mzYears = [];
    mzCustomers = [];
    mzContractors = [];
    mzWorks = [];
    mzStatuses = [];
    mzSgPos = false;
    mzFocusUin = null;
    mzSearch = '';
    document.getElementById('fltSearch').value = '';
    const sgInp = document.getElementById('fltSgPos');
    if (sgInp) sgInp.checked = false;
    refreshFilters();
  };
  let searchTimer = null;
  document.getElementById('fltSearch').oninput = e => {
    clearTimeout(searchTimer);
    const v = e.target.value.trim();
    searchTimer = setTimeout(() => { mzSearch = v; refreshFilters(); }, 180);
  };
  syncMenus();
  refreshFilters();
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
        EXON_HEAD.replace("__CKS_LOGO__", CKS_LOGO_B64)
        .replace("__MO_LOGO__", MO_LOGO_B64)
        .replace("__FAVICON_SVG__", FAVICON_SVG_HREF)
        .replace("__FAVICON_ICO__", FAVICON_ICO_HREF)
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

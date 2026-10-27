# Стройготовность и выплаты — 48 школ

Два интерактивных отчёта: методика (`index.html`/`sg-pay-analysis.html`) и операционный
дашборд (`dashboard.html`), связаны кросс-навигацией. Оба генерируются `build_html.py`.

**Онлайн:** https://dmi867.github.io/schools-sg-analytics/

Исходные Excel-файлы и данные в репозиторий не загружаются (см. `.gitignore`) — только
готовые страницы и код генератора.

Актуальные выгрузки в `data/raw/`: **1409** Simple List и ПИР (срезы 14.09.2026).

## Структура проекта

```
build_html.py       — генератор обеих страниц (единственный код в репозитории)
scripts/
  convert_per_object.py  — пересобирает data/per_object.json из xlsx-пар по школам
                            (запускать заново, если пришла новая выгрузка по объектам)
data/
  raw/               — «сырые» источники: Simple List, выгрузка ПИР, КСГ+Экспертиза
  addresses.json      — адрес/округ по УИН
  finance2026.json    — освоение бюджета 2026 по УИН
  per_object.json     — стройготовность + платежи по каждой школе (из raw xlsx-пар)
BACKLOG.md           — что из требований руководства/минфин-разбора ещё не сделано
```

## Обновить отчёт

Если обновились файлы в `data/raw/` (Simple List, ПИР, КСГ):

```bash
python3 build_html.py
```

Если пришла новая выгрузка по объектам (пары `УИН.xlsx` + `УИН_платежи.xlsx`) — сначала
положить их в `data/raw/` (или другую папку и передать путь первым аргументом) и пересобрать
`data/per_object.json`:

```bash
python3 scripts/convert_per_object.py data/raw
python3 build_html.py
```

Либо заполнить единый шаблон (одна таблица: СГ план/факт + платежи):

```bash
python3 scripts/make_sg_pay_template.py          # создать/обновить шаблон
# заполнить data/templates/шаблон_СГ_и_платежи.xlsx (лист «Данные»)
python3 scripts/import_sg_pay_template.py data/templates/шаблон_СГ_и_платежи.xlsx
python3 build_html.py
```

Либо динамика Минстроя из папки срезов `Simple_List/` (без сноса):

```bash
python3 scripts/build_minstroy_dynamics.py
python3 build_html.py
```

Затем:

```bash
git add index.html dashboard.html sg-pay-analysis.html
git commit -m "Update analytics"
git push
```

После push страницы обновятся через 1–2 минуты.

## Локально

```bash
python3 -m http.server 8080
```

Открыть: http://localhost:8080/index.html и http://localhost:8080/dashboard.html

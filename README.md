# Минстрой — СГ и выплаты · методика по школам

**Онлайн:** https://dmi867.github.io/schools-sg-analytics/

- `dashboard.html` — операционный дашборд по всему портфелю Минстроя (Exon UI:
  сводка / подрядчики / объекты). Данные из срезов Simple List.
- `index.html` / `sg-pay-analysis.html` — методика и выводы по школам (узкий портфель).

Исходные Excel в репозиторий не загружаются (см. `.gitignore`) — только готовые
страницы, JSON и код генераторов.

Актуальный срез Simple List: **2510** (25.10.2026).

## Структура

```
scripts/build_minstroy_dynamics.py  — Simple_List/*.xlsx → data/minstroy_dynamics.json
scripts/build_minstroy_dashboard.py — JSON → dashboard.html
build_html.py                      — методика index.html / sg-pay-analysis.html
Simple_List/                       — срезы Акцента (локально)
data/minstroy_dynamics.json        — портфель Минстроя для дашборда
data/per_object.json               — школы для методики
```

## Обновить дашборд Минстроя

Положить свежий `DDMM_Акцент_Simple List.xlsx` в `Simple_List/`, затем:

```bash
python3 scripts/build_minstroy_dynamics.py
python3 scripts/build_minstroy_dashboard.py
```

## Обновить методику по школам

```bash
python3 build_html.py
```

Если пришла новая выгрузка по объектам (пары `УИН.xlsx` + `УИН_платежи.xlsx`):

```bash
python3 scripts/convert_per_object.py data/raw
python3 build_html.py
```

Или единый шаблон СГ/платежей:

```bash
python3 scripts/make_sg_pay_template.py
# заполнить data/templates/шаблон_СГ_и_платежи.xlsx
python3 scripts/import_sg_pay_template.py data/templates/шаблон_СГ_и_платежи.xlsx
python3 build_html.py
```

Затем на GitHub Pages:

```bash
git add dashboard.html index.html sg-pay-analysis.html data/minstroy_dynamics.json
git commit -m "Update MinStroy dashboard"
git push
```

## Локально

```bash
python3 -m http.server 8080
```

Открыть: http://localhost:8080/dashboard.html и http://localhost:8080/index.html

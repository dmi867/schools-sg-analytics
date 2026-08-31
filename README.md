# Стройготовность и выплаты — 47 школ

Интерактивный отчёт по связи СГ и выплат.

**Онлайн:** https://dmi867.github.io/schools-sg-analytics/

Исходные Excel-файлы в репозиторий не загружаются — только готовая страница.

## Обновить отчёт

```bash
python3 build_html.py
git add index.html sg-pay-analysis.html
git commit -m "Update analytics"
git push
```

После push страница обновится через 1–2 минуты.

## Локально

```bash
python3 -m http.server 8080
```

Открыть: http://localhost:8080

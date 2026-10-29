# Минстрой — СГ и выплаты

**Онлайн:** http://sgp.cksdash.site/

Репозиторий: https://github.com/cks-hub/minstroy-sg-pay

- `dashboard.html` — операционный дашборд портфеля Минстроя (СГ и оплата).
- `index.html` / `sg-pay-analysis.html` — методика по школам.

## Обновить дашборд

```bash
# свежий Simple List в Simple_List/
python3 scripts/build_minstroy_dynamics.py
python3 scripts/build_minstroy_dashboard.py
```

Методика: `python3 build_html.py`


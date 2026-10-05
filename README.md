# tcmb-exchange-rates

Downloads the official daily exchange rates published by the Central Bank of the Republic of Türkiye (TCMB) for any date range and exports them to Excel or CSV.

- Reads TCMB's public XML bulletins (`tcmb.gov.tr/kurlar/YYYYMM/DDMMYYYY.xml`)
- Skips weekends without a request; public holidays (HTTP 404) are skipped automatically
- Caches every bulletin on disk, so repeated runs do not hit the server again
- Polite by default: short pause between requests, retries on network errors
- Excel output: a summary sheet (one row per day, one column per currency, line chart) and a filterable data sheet with buying/selling and banknote rates
- Handles currencies quoted per 100 units (e.g. JPY)

## Usage

```bash
python -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python tcmb_rates.py --start 2026-09-01 --end 2026-09-30 --codes USD EUR GBP --out rates.xlsx
.venv/bin/python tcmb_rates.py --days 90 --out rates.csv
.venv/bin/python tcmb_rates.py --days 7 --codes            # all currencies
```

Example: September 2026, three currencies → `72 rows, 24 days -> rates.xlsx`.

## Tests

```bash
.venv/bin/python -m pytest -q
```

HTTP is replaced with a fake session in tests, so they run offline.

## Türkçe

TCMB'nin her gün yayımladığı gösterge niteliğindeki döviz kurlarını istenen tarih aralığı için indirir ve Excel ya da CSV olarak verir. Hafta sonları ve resmî tatiller atlanır, indirilen bültenler önbelleğe alınır. Muhasebe, fiyatlandırma ve raporlama işlerinde elle kur girişini ortadan kaldırır.

## License

MIT

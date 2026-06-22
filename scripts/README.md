# Naver Finance 주요뉴스 크롤러

This script extracts the 주요뉴스 list from Naver Finance and saves the results to an Excel file with columns: 제목, URL, 언론사, 발행날짜.

Prerequisites
- A Python virtual environment (you mentioned one exists). Activate it before installing dependencies.

Install dependencies

```powershell
# activate your venv (example on Windows)
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Run

```powershell
python scripts\crawl_naver_finance.py --pages 1 --output news.xlsx
```

Options
- `--pages`: Number of pages to crawl (default 1)
- `--date`: Optional date filter like `2026-06-22` (format YYYY-MM-DD)
- `--output`: Output Excel filename (default `output.xlsx`)
- `--delay`: Seconds to wait between page requests (default 1.0)

Notes
- Be mindful of `robots.txt` and service terms when scraping. Add rate-limiting or caching for larger crawls.
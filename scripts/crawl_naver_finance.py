#!/usr/bin/env python3
"""Naver Finance 주요뉴스 크롤러

Usage:
  python crawl_naver_finance.py --pages 1 --output output.xlsx

Extracts: 제목, URL, 언론사, 발행날짜
"""
from __future__ import annotations

import argparse
import time
from typing import List, Dict
from urllib.parse import urljoin

from openpyxl import Workbook
import requests
from bs4 import BeautifulSoup


HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9",
}


def fetch_page(session: requests.Session, url: str, params: dict | None = None) -> str:
    resp = session.get(url, params=params, headers=HEADERS, timeout=15)
    resp.raise_for_status()
    return resp.text


def parse_news_from_html(html: str, base_url: str = "https://finance.naver.com") -> List[Dict[str, str]]:
    soup = BeautifulSoup(html, "html.parser")
    results: List[Dict[str, str]] = []

    # Select list items under the main news list
    for li in soup.select("ul.newsList li"):
        # Title and URL
        a = li.select_one("dd.articleSubject a") or li.find("a")
        if not a:
            continue
        title = a.get_text(strip=True)
        href = a.get("href", "")
        url = urljoin(base_url, href)

        # Press and date are inside dd.articleSummary > span.press and span.wdate
        summary = li.select_one("dd.articleSummary")
        press = ""
        date = ""
        if summary:
            press_el = summary.select_one("span.press")
            date_el = summary.select_one("span.wdate")
            if press_el:
                press = press_el.get_text(strip=True)
            if date_el:
                date = date_el.get_text(strip=True)

        results.append({"제목": title, "URL": url, "언론사": press, "발행날짜": date})

    return results


def crawl(pages: int = 1, date: str | None = None, delay: float = 1.0) -> List[Dict[str, str]]:
    base_page = "https://finance.naver.com/news/mainnews.naver"
    session = requests.Session()
    all_items: List[Dict[str, str]] = []

    for p in range(1, pages + 1):
        params = {"page": p}
        if date:
            params["date"] = date

        html = fetch_page(session, base_page, params=params)
        items = parse_news_from_html(html, base_url=base_page)
        all_items.extend(items)
        time.sleep(delay)

    return all_items


def save_to_excel(items: List[Dict[str, str]], output: str) -> None:
    if not items:
        print("No items to save.")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "news"
    headers = ["제목", "URL", "언론사", "발행날짜"]
    ws.append(headers)
    for it in items:
        row = [it.get(h, "") for h in headers]
        ws.append(row)
    wb.save(output)
    print(f"Saved {len(items)} rows to {output}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Naver Finance 주요뉴스 크롤러")
    parser.add_argument("--pages", type=int, default=1, help="Number of pages to crawl (default: 1)")
    parser.add_argument("--date", type=str, default=None, help="Date filter in YYYY-MM-DD (optional)")
    parser.add_argument("--output", type=str, default="output.xlsx", help="Output Excel filename")
    parser.add_argument("--delay", type=float, default=1.0, help="Delay between page requests in seconds")

    args = parser.parse_args()

    items = crawl(pages=args.pages, date=args.date, delay=args.delay)
    save_to_excel(items, args.output)


if __name__ == "__main__":
    main()

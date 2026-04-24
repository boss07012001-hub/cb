#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TWSE 可轉換公司債競拍監控系統 v3
- 以「投標開始日」前5天發送通知
- cb168 歷史高低價查詢連結（含搜尋提示）
"""

import os
import re
import smtplib
import requests
import traceback
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

EMAIL_SENDER      = os.environ.get("EMAIL_SENDER", "")
EMAIL_PASSWORD    = os.environ.get("EMAIL_PASSWORD", "")
EMAIL_RECEIVER    = os.environ.get("EMAIL_RECEIVER", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
LINE_NOTIFY_TOKEN = os.environ.get("LINE_NOTIFY_TOKEN", "")

DAYS_AHEAD = 5   # 投標開始日前幾天通知
NOTIFY_FIELD = "投標開始日"   # 以此欄位為基準


def fetch_auction_data() -> list[dict]:
    today = datetime.now()
    year_roc = today.year - 1911

    urls_to_try = [
        "https://www.twse.com.tw/rwd/zh/announcement/auction?response=json",
        f"https://www.twse.com.tw/rwd/zh/announcement/auction?response=json&year={today.year}",
        f"https://www.twse.com.tw/rwd/zh/announcement/auction?response=json&year={year_roc}",
        "https://www.twse.com.tw/zh/announcement/auction?response=json",
    ]

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.twse.com.tw/zh/announcement/auction.html",
        "Accept": "application/json, text/javascript, */*",
        "Accept-Language": "zh-TW,zh;q=0.9",
    }

    for url in urls_to_try:
        try:
            print(f"  嘗試 URL: {url}")
            resp = requests.get(url, headers=headers, timeout=20)
            print(f"  HTTP 狀態碼: {resp.status_code}")
            if resp.status_code != 200:
                continue
            try:
                data = resp.json()
            except Exception:
                continue
            print(f"  回應 stat: {data.get('stat', 'N/A')}")
            if "fields" in data:
                print(f"  欄位名稱: {data['fields']}")
            rows = data.get("data", [])
            print(f"  資料列數: {len(rows)}")
            if rows:
                print(f"  第一列範例: {rows[0]}")
                fields = data.get("fields", [])
                records = []
                for row in rows:
                    if fields:
                        record = dict(zip(fields, row))
                    else:
                        record = {f"col_{i}": v for i, v in enumerate(row)}
                    records.append(record)
                return records
        except Exception as e:
            print(f"  例外: {e}")
            continue

    print("⚠️ 所有 URL 均無法取得資料")
    return []


def tw_date_to_datetime(val: str) -> datetime | None:
    if not val:
        return None
    val = str(val).strip().replace("/", "-").replace(".", "-")
    m = re.match(r"^(\d{3})-(\d{2})-(\d{2})$", val)
    if m:
        try:
            return datetime(int(m.group(1)) + 1911, int(m.group(2)), int(m.group(3)))
        except Exception:
            pass
    m2 = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", val)
    if m2:
        try:
            return datetime(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))
        except Exception:
            pass
    return None


def get_stock_info(record: dict) -> tuple[str, str]:
    name_keys = ["證券名稱", "股票名稱", "公司名稱", "名稱", "col_1"]
    code_keys = ["證券代號", "股票代號", "代號", "col_2"]
    name = ""
    for k in name_keys:
        if record.get(k):
            name = str(record[k]).strip()
            break
    code = ""
    for k in code_keys:
        if record.get(k):
            code = str(record[k]).strip()
            break
    return name, code


def get_upcoming_auctions(records: list[dict]) -> list[dict]:
    """以「投標開始日」為基準，篩選前5天內需通知的項目"""
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    # 通知條件：投標開始日 在今天 到 今天+5天 之間
    threshold = today + timedelta(days=DAYS_AHEAD)

    upcoming = []
    for record in records:
        # 取得投標開始日
        bid_start_val = record.get(NOTIFY_FIELD, "")
        bid_start_dt = tw_date_to_datetime(bid_start_val)
        if bid_start_dt is None:
            continue

        # 取得開標日期（用於顯示）
        auction_dt = tw_date_to_datetime(record.get("開標日期", ""))

        if today <= bid_start_dt <= threshold:
            days_left = (bid_start_dt - today).days
            record["_bid_start_dt"] = bid_start_dt
            record["_auction_dt"]   = auction_dt
            record["_days_left"]    = days_left
            upcoming.append(record)

    # 去重（以證券代號為 key）
    seen = set()
    unique = []
    for r in upcoming:
        _, code = get_stock_info(r)
        key = code o

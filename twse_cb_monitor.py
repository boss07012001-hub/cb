#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TWSE 可轉換公司債競拍監控系統 v2
修正：自動偵測 API 欄位格式，支援多種日期欄位名稱
"""

import os
import re
import json
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

DAYS_AHEAD = 7  # 提前幾天通知（改為7天確保不漏）


# ══════════════════════════════════════════════
# 1. 抓取 TWSE 競拍公告（多重 URL 嘗試）
# ══════════════════════════════════════════════
def fetch_auction_data() -> list[dict]:
    """嘗試多個 URL 格式抓取競拍資料"""
    today = datetime.now()
    year_roc = today.year - 1911  # 民國年

    urls_to_try = [
        # 不帶年份
        "https://www.twse.com.tw/rwd/zh/announcement/auction?response=json",
        # 帶西元年
        f"https://www.twse.com.tw/rwd/zh/announcement/auction?response=json&year={today.year}",
        # 帶民國年
        f"https://www.twse.com.tw/rwd/zh/announcement/auction?response=json&year={year_roc}",
        # 舊版 URL
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

            # 嘗試解析 JSON
            try:
                data = resp.json()
            except Exception:
                print(f"  無法解析 JSON，跳過")
                continue

            print(f"  回應 stat: {data.get('stat', 'N/A')}")
            print(f"  回應 keys: {list(data.keys())}")

            # 印出 fields（欄位名稱）
            if "fields" in data:
                print(f"  欄位名稱: {data['fields']}")

            # 取得資料列
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
                        # 沒有 fields，直接用索引
                        record = {f"col_{i}": v for i, v in enumerate(row)}
                    records.append(record)
                return records

        except Exception as e:
            print(f"  例外: {e}")
            continue

    print("⚠️ 所有 URL 均無法取得資料")
    return []


def tw_date_to_datetime(val: str) -> datetime | None:
    """解析民國年或西元年日期字串"""
    if not val:
        return None
    val = str(val).strip().replace("/", "-").replace(".", "-")

    # 民國年 114-04-29
    m = re.match(r"^(\d{3})-(\d{2})-(\d{2})$", val)
    if m:
        try:
            y = int(m.group(1)) + 1911
            return datetime(y, int(m.group(2)), int(m.group(3)))
        except Exception:
            pass

    # 西元年 2025-04-29
    m2 = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", val)
    if m2:
        try:
            return datetime(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))
        except Exception:
            pass

    # 純數字 20250429
    m3 = re.match(r"^(\d{4})(\d{2})(\d{2})$", val)
    if m3:
        try:
            return datetime(int(m3.group(1)), int(m3.group(2)), int(m3.group(3)))
        except Exception:
            pass

    return None


def find_date_field(record: dict) -> datetime | None:
    """自動尋找記錄中的日期欄位"""
    # 優先搜尋這些關鍵字
    priority_keys = ["開標日期", "競拍日期", "拍賣日期", "投標結束日", "結束日期"]
    for key in priority_keys:
        if key in record:
            dt = tw_date_to_datetime(record[key])
            if dt:
                return dt

    # 搜尋所有欄位，找看起來像日期的值
    for key, val in record.items():
        if not val:
            continue
        dt = tw_date_to_datetime(str(val))
        if dt and dt.year >= 2024:
            return dt

    return None


def get_stock_info(record: dict) -> tuple[str, str]:
    """從記錄中提取股票名稱和代號"""
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
    """篩選未來 DAYS_AHEAD 天內的競拍"""
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    threshold = today + timedelta(days=DAYS_AHEAD)

    upcoming = []
    for record in records:
        dt = find_date_field(record)
        if dt is None:
            continue
        if today <= dt <= threshold:
            record["_auction_dt"] = dt
            record["_days_left"] = (dt - today).days
            upcoming.append(record)

    return sorted(upcoming, key=lambda x: x["_auction_dt"])


# ══════════════════════════════════════════════
# 2. 查詢歷史可轉債
# ══════════════════════════════════════════════
def fetch_cb_history(stock_code: str) -> list[dict]:
    url = f"https://www.twse.com.tw/rwd/zh/bond/BFI82U?response=json&strDate=&endDate=&stockNo={stock_code}"
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://www.twse.com.tw/"}
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        data = resp.json()
        if data.get("stat") == "OK" and data.get("data"):
            fields = data.get("fields", [])
            return [dict(zip(fields, row)) for row in data["data"]]
    except Exception as e:
        print(f"查詢歷史可轉債失敗（{stock_code}）：{e}")
    return []


# ══════════════════════════════════════════════
# 3. 發送 Email
# ══════════════════════════════════════════════
def send_email(subject: str, html_body: str) -> bool:
    if not all([EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECEIVER]):
        print("[跳過] Email 設定不完整")
        return False
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = EMAIL_SENDER
        msg["To"]      = EMAIL_RECEIVER
        msg.attach(MIMEText(html_body, "html", "utf-8"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, msg.as_string())
        print("✅ Email 發送成功")
        return True
    except Exception as e:
        print(f"❌ Email 發送失敗：{e}")
        return False


def build_html(auctions: list[dict]) -> str:
    today_str = datetime.now().strftime("%Y/%m/%d")
    rows_html = ""
    for a in auctions:
        name, code = get_stock_info(a)
        dt_str = a["_auction_dt"].strftime("%Y/%m/%d")
        days   = a["_days_left"]

        # 歷史可轉債
        cb_history = fetch_cb_history(code) if code else []
        if cb_history:
            cb_lines = ""
            for h in cb_history[:3]:
                cb_id = h.get("債券代號", h.get("代號", ""))
                cb_lines += f"<li>{cb_id}</li>"
            history_html = f"<ul style='margin:0;padding-left:18px'>{cb_lines}</ul>"
        else:
            history_html = "無歷史紀錄"

        badge_color = "#d32f2f" if days <= 2 else "#ff6f00" if days <= 4 else "#1565c0"
        rows_html += f"""
        <tr>
          <td style='padding:10px;border-bottom:1px solid #eee;'><b>{name}</b><br><span style='color:#666;font-size:12px'>{code}</span></td>
          <td style='padding:10px;border-bottom:1px solid #eee;text-align:center'>{dt_str}</td>
          <td style='padding:10px;border-bottom:1px solid #eee;text-align:center'>
            <span style='background:{badge_color};color:#fff;border-radius:12px;padding:2px 10px;font-size:13px'>{days} 天後</span>
          </td>
          <td style='padding:10px;border-bottom:1px solid #eee;font-size:13px'>{history_html}</td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html lang="zh-TW"><head><meta charset="UTF-8">
<style>
  body{{font-family:"Microsoft JhengHei",Arial,sans-serif;background:#f0f2f5;padding:20px;}}
  .card{{background:#fff;border-radius:12px;padding:24px;max-width:720px;margin:auto;box-shadow:0 2px 12px rgba(0,0,0,0.1);}}
  h2{{color:#1565c0;margin-top:0;}}
  table{{width:100%;border-collapse:collapse;}}
  th{{background:#1565c0;color:#fff;padding:10px;text-align:left;}}
  .footer{{font-size:11px;color:#aaa;text-align:center;margin-top:16px;}}
</style></head>
<body><div class="card">
  <h2>🔔 可轉債競拍預警通知</h2>
  <p style="color:#666">📅 通知日期：{today_str}　｜　共 <b>{len(auctions)}</b> 筆即將競拍</p>
  <table>
    <tr>
      <th>標的名稱</th><th>開標日期</th><th>距今天數</th><th>歷史可轉債</th>
    </tr>
    {rows_html}
  </table>
  <p class="footer">
    資料來源：<a href="https://www.twse.com.tw/zh/announcement/auction.html">臺灣證券交易所競價拍賣公告</a><br>
    本通知由自動化系統產生，內容僅供參考，不構成投資建議。
  </p>
</div></body></html>"""


# ══════════════════════════════════════════════
# 主流程
# ══════════════════════════════════════════════
def main():
    today_str = datetime.now().strftime("%Y/%m/%d")
    print(f"\n{'='*55}")
    print(f"  TWSE 可轉債競拍監控  v2  |  {today_str}")
    print(f"{'='*55}\n")

    # 抓取資料
    print("📡 抓取 TWSE 競拍公告...")
    records = fetch_auction_data()
    print(f"   取得 {len(records)} 筆原始資料\n")

    if not records:
        print("⚠️ 無法取得資料，傳送警告 Email")
        send_email(
            f"【可轉債監控警告】{today_str} 無法取得 TWSE 資料",
            "<p>今日執行時無法從 TWSE 取得競拍資料，請手動確認。</p>"
            "<p><a href='https://www.twse.com.tw/zh/announcement/auction.html'>點此查看官網</a></p>"
        )
        return

    # 篩選即將競拍
    upcoming = get_upcoming_auctions(records)
    print(f"📋 未來 {DAYS_AHEAD} 天內即將競拍：{len(upcoming)} 筆\n")

    if not upcoming:
        print("✅ 目前無即將競拍項目，無需通知。")
        return

    for a in upcoming:
        name, code = get_stock_info(a)
        print(f"  → {name}（{code}）開標日：{a['_auction_dt'].strftime('%Y/%m/%d')}，距今 {a['_days_left']} 天")

    # 發送 Email
    print("\n📧 發送 Email 通知...")
    subject = f"【可轉債競拍預警】{len(upcoming)} 筆即將競拍 | {today_str}"
    html = build_html(upcoming)
    send_email(subject, html)

    print("\n🎉 完成！")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        raise

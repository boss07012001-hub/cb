#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TWSE 可轉換公司債競拍監控系統 v3
- 以「投標開始日」前5天發送通知
- cb168 歷史高低價查詢連結（含一鍵複製搜尋字）
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

DAYS_AHEAD    = 5
NOTIFY_FIELD  = "投標開始日"


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
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    threshold = today + timedelta(days=DAYS_AHEAD)
    upcoming = []
    for record in records:
        bid_start_val = record.get(NOTIFY_FIELD, "")
        bid_start_dt = tw_date_to_datetime(bid_start_val)
        if bid_start_dt is None:
            continue
        auction_dt = tw_date_to_datetime(record.get("開標日期", ""))
        if today <= bid_start_dt <= threshold:
            days_left = (bid_start_dt - today).days
            record["_bid_start_dt"] = bid_start_dt
            record["_auction_dt"]   = auction_dt
            record["_days_left"]    = days_left
            upcoming.append(record)
    seen = set()
    unique = []
    for r in upcoming:
        _, code = get_stock_info(r)
        key = code or str(r)
        if key not in seen:
            seen.add(key)
            unique.append(r)
    return sorted(unique, key=lambda x: x["_bid_start_dt"])


def extract_base_name(name: str) -> str:
    name = name.strip()
    name = re.sub(r'[一二三四五六七八九十]+$', '', name)
    name = re.sub(r'\d+$', '', name)
    return name.strip()


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
        bid_start_str = a["_bid_start_dt"].strftime("%Y/%m/%d")
        auction_str   = a["_auction_dt"].strftime("%Y/%m/%d") if a.get("_auction_dt") else "—"
        days          = a["_days_left"]
        base_name     = extract_base_name(name)

        cb168_btn = f"""
        <div style='margin-top:8px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;'>
          <a href="https://cb168.netlify.app/" target="_blank"
            style="background:#e8f4fd;color:#1565c0;border:1px solid #1565c0;
                   border-radius:6px;padding:4px 12px;font-size:12px;
                   text-decoration:none;white-space:nowrap;">
            🔍 點此查歷史高低價
          </a>
          <span style='font-size:12px;color:#555;'>→ 複製搜尋字：</span>
          <code style='background:#f0f0f0;border:1px solid #ccc;border-radius:4px;
                       padding:2px 10px;font-size:13px;font-family:inherit;
                       color:#c0392b;font-weight:bold;cursor:text;
                       user-select:all;'>{base_name}</code>
        </div>"""

        badge_color = "#d32f2f" if days <= 1 else "#ff6f00" if days <= 3 else "#1565c0"
        rows_html += f"""
        <tr>
          <td style='padding:12px 10px;border-bottom:1px solid #eee;'>
            <b style='font-size:15px'>{name}</b><br>
            <span style='color:#888;font-size:12px'>股票代號：{code}</span>
            {cb168_btn}
          </td>
          <td style='padding:12px 10px;border-bottom:1px solid #eee;text-align:center;white-space:nowrap'>
            {bid_start_str}<br>
            <span style='color:#888;font-size:11px'>開標：{auction_str}</span>
          </td>
          <td style='padding:12px 10px;border-bottom:1px solid #eee;text-align:center'>
            <span style='background:{badge_color};color:#fff;border-radius:12px;
                         padding:3px 12px;font-size:13px;font-weight:bold'>{days} 天後</span>
          </td>
        </tr>"""

    cb168_tip = """
    <div style='background:#fff8e1;border-left:4px solid #ff6f00;padding:12px 16px;
                border-radius:4px;margin-top:20px;font-size:13px;color:#555;'>
      <b>📊 如何查詢歷史可轉債高低價？</b><br>
      1. 點「查歷史高低價」按鈕進入 cb168 網站<br>
      2. 點一下紅色文字（公司名稱）即可全選<br>
      3. Ctrl+C 複製 → 貼到 cb168 搜尋框 → 查詢
    </div>"""

    return f"""<!DOCTYPE html>
<html lang="zh-TW"><head><meta charset="UTF-8">
<style>
  body{{font-family:"Microsoft JhengHei",Arial,sans-serif;background:#f0f2f5;padding:20px;}}
  .card{{background:#fff;border-radius:12px;padding:24px;max-width:720px;margin:auto;box-shadow:0 2px 12px rgba(0,0,0,0.1);}}
  h2{{color:#1565c0;margin-top:0;}}
  table{{width:100%;border-collapse:collapse;}}
  th{{background:#1565c0;color:#fff;padding:10px 12px;text-align:left;font-size:14px;}}
  tr:hover td{{background:#f8f9ff;}}
  .footer{{font-size:11px;color:#aaa;text-align:center;margin-top:20px;}}
</style></head>
<body><div class="card">
  <h2>🔔 可轉債競拍預警通知</h2>
  <p style="color:#666;margin-bottom:16px">📅 通知日期：{today_str}　｜　共 <b>{len(auctions)}</b> 筆即將開放投標</p>
  <table>
    <tr>
      <th>標的名稱</th>
      <th style='text-align:center'>投標開始日 / 開標日</th>
      <th style='text-align:center'>距投標開始</th>
    </tr>
    {rows_html}
  </table>
  {cb168_tip}
  <p class="footer">
    資料來源：<a href="https://www.twse.com.tw/zh/announcement/auction.html">臺灣證券交易所競價拍賣公告</a>　｜　
    歷史高低價：<a href="https://cb168.netlify.app/">cb168.netlify.app</a><br>
    本通知由自動化系統產生，內容僅供參考，不構成投資建議。
  </p>
</div></body></html>"""


def main():
    today_str = datetime.now().strftime("%Y/%m/%d")
    print(f"\n{'='*55}")
    print(f"  TWSE 可轉債競拍監控  v3  |  {today_str}")
    print(f"{'='*55}\n")

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

    upcoming = get_upcoming_auctions(records)
    print(f"📋 投標開始日在未來 {DAYS_AHEAD} 天內：{len(upcoming)} 筆\n")

    if not upcoming:
        print("✅ 目前無即將開放投標項目，無需通知。")
        return

    for a in upcoming:
        name, code = get_stock_info(a)
        print(f"  → {name}（{code}）投標開始：{a['_bid_start_dt'].strftime('%Y/%m/%d')}，距今 {a['_days_left']} 天")

    print("\n📧 發送 Email 通知...")
    subject = f"【可轉債競拍預警】{len(upcoming)} 筆即將開放投標 | {today_str}"
    html = build_html(upcoming)
    send_email(subject, html)

    print("\n🎉 完成！")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        raise

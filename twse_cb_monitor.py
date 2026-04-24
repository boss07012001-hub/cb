#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TWSE 可轉換公司債競拍監控系統
每日自動抓取競拍公告，提前5天通知，並附帶 AI 分析報告
通知管道：LINE Notify + Email
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

# ─────────────────────────────────────────────
# 環境變數設定（於 GitHub Secrets 中填寫）
# ─────────────────────────────────────────────
LINE_NOTIFY_TOKEN = os.environ.get("LINE_NOTIFY_TOKEN", "")
EMAIL_SENDER      = os.environ.get("EMAIL_SENDER", "")       # Gmail 帳號
EMAIL_PASSWORD    = os.environ.get("EMAIL_PASSWORD", "")     # Gmail 應用程式密碼
EMAIL_RECEIVER    = os.environ.get("EMAIL_RECEIVER", "")     # 收件人 Email
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")  # Claude API Key

DAYS_AHEAD = 5  # 提前幾天通知


# ══════════════════════════════════════════════
# 1. 抓取 TWSE 競拍公告
# ══════════════════════════════════════════════
def fetch_auction_list(year: int) -> list[dict]:
    """
    抓取指定年度的競拍公告清單
    回傳格式：[{name, stockCode, cbCode, auctionDate, ...}, ...]
    """
    url = f"https://www.twse.com.tw/rwd/zh/announcement/auction?response=json&year={year}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://www.twse.com.tw/zh/announcement/auction.html",
    }
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        records = []
        if data.get("stat") == "OK" and data.get("data"):
            fields = data.get("fields", [])
            for row in data["data"]:
                record = dict(zip(fields, row))
                records.append(record)
        return records
    except Exception as e:
        print(f"[警告] 抓取競拍清單失敗：{e}")
        return []


def parse_auction_date(record: dict) -> datetime | None:
    """嘗試解析競拍日期欄位（民國年或西元年）"""
    # 常見欄位名稱
    for key in ["競拍日期", "拍賣日期", "日期", "auctionDate"]:
        val = record.get(key, "")
        if not val:
            continue
        val = val.strip().replace("/", "-")
        # 民國年 e.g. "114/05/15" or "114-05-15"
        m = re.match(r"^(\d{3})-(\d{2})-(\d{2})$", val)
        if m:
            y = int(m.group(1)) + 1911
            mo, d = int(m.group(2)), int(m.group(3))
            return datetime(y, mo, d)
        # 西元年 e.g. "2025-05-15"
        m2 = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", val)
        if m2:
            return datetime(int(m2.group(1)), int(m2.group(2)), int(m2.group(3)))
    return None


def get_upcoming_auctions() -> list[dict]:
    """取得未來 DAYS_AHEAD 天內的競拍項目"""
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    threshold = today + timedelta(days=DAYS_AHEAD)

    results = []
    for year in [today.year, today.year + 1]:  # 跨年時也一起查
        for record in fetch_auction_list(year):
            auction_dt = parse_auction_date(record)
            if auction_dt is None:
                continue
            if today <= auction_dt <= threshold:
                days_left = (auction_dt - today).days
                record["_auction_dt"]  = auction_dt
                record["_days_left"]   = days_left
                results.append(record)

    # 去重（以可轉債代號為 key）
    seen = set()
    unique = []
    for r in results:
        key = r.get("債券代號", "") or r.get("可轉債代號", "") or str(r)
        if key not in seen:
            seen.add(key)
            unique.append(r)

    return sorted(unique, key=lambda x: x["_auction_dt"])


# ══════════════════════════════════════════════
# 2. 查詢歷史可轉債紀錄（TWSE 公開資料）
# ══════════════════════════════════════════════
def fetch_cb_history(stock_code: str) -> list[dict]:
    """
    查詢該標的股過去曾發行的可轉債（透過 TWSE 轉換公司債資料）
    回傳：[{cbCode, name, issueDate, conversionPrice, high, low}, ...]
    """
    url = (
        f"https://www.twse.com.tw/rwd/zh/bond/BFI82U"
        f"?response=json&strDate=&endDate=&stockNo={stock_code}"
    )
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://www.twse.com.tw/"}
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        records = []
        if data.get("stat") == "OK" and data.get("data"):
            fields = data.get("fields", [])
            for row in data["data"]:
                records.append(dict(zip(fields, row)))
        return records
    except Exception as e:
        print(f"[警告] 查詢可轉債歷史失敗（{stock_code}）：{e}")
        return []


def get_cb_price_range(cb_code: str) -> dict:
    """
    查詢指定可轉債的歷史最高/最低成交價
    使用 TWSE 轉換公司債行情查詢
    """
    url = f"https://www.twse.com.tw/rwd/zh/bond/BOND_NY?bondId={cb_code}&response=json"
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://www.twse.com.tw/"}
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        highs, lows = [], []
        if data.get("stat") == "OK" and data.get("data"):
            for row in data["data"]:
                # 欄位：日期、最高價、最低價、收盤價 ... 依實際欄位順序
                if len(row) >= 4:
                    try:
                        highs.append(float(str(row[1]).replace(",", "")))
                        lows.append(float(str(row[2]).replace(",", "")))
                    except (ValueError, IndexError):
                        pass
        if highs and lows:
            return {"high": max(highs), "low": min(lows)}
    except Exception as e:
        print(f"[警告] 查詢可轉債行情失敗（{cb_code}）：{e}")
    return {}


# ══════════════════════════════════════════════
# 3. Claude AI 分析
# ══════════════════════════════════════════════
def ai_analyze(auction: dict, cb_history: list[dict]) -> str:
    """
    呼叫 Claude API 分析：
    1. 公司為何此次發行可轉債
    2. 未來股價走勢預估
    """
    if not ANTHROPIC_API_KEY:
        return "（未設定 ANTHROPIC_API_KEY，跳過 AI 分析）"

    stock_name = (
        auction.get("股票名稱")
        or auction.get("公司名稱")
        or auction.get("name", "未知")
    )
    stock_code = (
        auction.get("股票代號")
        or auction.get("stockCode", "")
    )
    cb_code = (
        auction.get("債券代號")
        or auction.get("可轉債代號", "")
    )
    auction_date = auction["_auction_dt"].strftime("%Y-%m-%d")
    days_left = auction["_days_left"]

    history_text = "無歷史可轉債紀錄。"
    if cb_history:
        lines = []
        for h in cb_history[:5]:
            lines.append(str(h))
        history_text = "\n".join(lines)

    prompt = f"""
你是一位台灣資本市場分析師，請根據以下資訊提供專業分析（繁體中文，300字內）：

【標的資訊】
- 公司名稱：{stock_name}（股票代號：{stock_code}）
- 此次可轉債代號：{cb_code}
- 競拍日期：{auction_date}（距今 {days_left} 天）

【歷史可轉債紀錄】
{history_text}

請分析：
1. 📋 公司此次發行可轉債的可能原因（資金用途、財務狀況推測）
2. 📈 發行可轉債對未來股價走勢的影響評估（正面/中性/負面）
3. ⚠️ 投資人需注意的風險提示

請以條列式呈現，語氣專業但易懂。
"""

    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    body = {
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 1000,
        "messages": [{"role": "user", "content": prompt}],
    }
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers=headers,
            json=body,
            timeout=30,
        )
        resp.raise_for_status()
        result = resp.json()
        return result["content"][0]["text"].strip()
    except Exception as e:
        return f"（AI 分析失敗：{e}）"


# ══════════════════════════════════════════════
# 4. 組裝通知內容
# ══════════════════════════════════════════════
def build_report(auction: dict, cb_history: list[dict], ai_text: str) -> tuple[str, str]:
    """
    回傳 (line_text, html_body)
    """
    stock_name = (
        auction.get("股票名稱")
        or auction.get("公司名稱", "未知")
    )
    stock_code = auction.get("股票代號") or auction.get("stockCode", "")
    cb_code    = auction.get("債券代號") or auction.get("可轉債代號", "")
    auction_date = auction["_auction_dt"].strftime("%Y/%m/%d")
    days_left  = auction["_days_left"]

    # ── 歷史可轉債摘要 ──
    if cb_history:
        history_lines = []
        for h in cb_history[:5]:
            cb_id   = h.get("債券代號", h.get("代號", ""))
            price_r = get_cb_price_range(cb_id) if cb_id else {}
            high_s  = f"{price_r['high']:.2f}" if price_r.get("high") else "N/A"
            low_s   = f"{price_r['low']:.2f}"  if price_r.get("low")  else "N/A"
            history_lines.append(
                f"  • {cb_id} | 歷史最高：{high_s} / 最低：{low_s}"
            )
        history_block = "\n".join(history_lines)
        history_html  = "".join(
            f"<li>{l.strip()}</li>" for l in history_lines
        )
    else:
        history_block = "  此標的尚無歷史可轉債紀錄"
        history_html  = "<li>此標的尚無歷史可轉債紀錄</li>"

    # ── LINE 純文字 ──
    line_msg = f"""
🔔【可轉債競拍預警】距今 {days_left} 天

📌 標的：{stock_name}（{stock_code}）
🏷️ 可轉債代號：{cb_code}
📅 競拍日期：{auction_date}

📂 歷史可轉債紀錄：
{history_block}

🤖 AI 分析：
{ai_text}

🔗 官方公告：https://www.twse.com.tw/zh/announcement/auction.html
""".strip()

    # ── Email HTML ──
    html_body = f"""
<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<style>
  body {{ font-family: "Microsoft JhengHei", Arial, sans-serif; background:#f5f5f5; padding:20px; }}
  .card {{ background:#fff; border-radius:10px; padding:24px; max-width:680px;
           margin:auto; box-shadow:0 2px 8px rgba(0,0,0,0.12); }}
  h2 {{ color:#d32f2f; margin-top:0; }}
  .badge {{ display:inline-block; background:#ff6f00; color:#fff;
            border-radius:20px; padding:3px 12px; font-size:13px; }}
  table {{ width:100%; border-collapse:collapse; margin:12px 0; }}
  th {{ background:#1565c0; color:#fff; padding:8px 12px; text-align:left; }}
  td {{ padding:8px 12px; border-bottom:1px solid #eee; }}
  .ai-box {{ background:#f0f4ff; border-left:4px solid #1565c0;
             padding:14px 16px; border-radius:4px; margin-top:16px;
             white-space:pre-wrap; line-height:1.7; }}
  .footer {{ font-size:12px; color:#999; margin-top:20px; text-align:center; }}
</style>
</head>
<body>
<div class="card">
  <h2>🔔 可轉債競拍預警通知</h2>
  <span class="badge">距競拍日還有 {days_left} 天</span>

  <table style="margin-top:16px;">
    <tr><th colspan="2">基本資訊</th></tr>
    <tr><td>標的名稱</td><td><b>{stock_name}</b>（{stock_code}）</td></tr>
    <tr><td>可轉債代號</td><td>{cb_code}</td></tr>
    <tr><td>競拍日期</td><td>{auction_date}</td></tr>
  </table>

  <table>
    <tr><th>歷史可轉債紀錄</th><th>歷史最高價</th><th>歷史最低價</th></tr>
    {_build_history_rows(cb_history)}
  </table>

  <div class="ai-box">
    <b>🤖 AI 智能分析</b><br><br>
    {ai_text.replace(chr(10), '<br>')}
  </div>

  <p class="footer">
    資料來源：<a href="https://www.twse.com.tw/zh/announcement/auction.html">臺灣證券交易所競價拍賣公告</a><br>
    本通知由自動化系統產生，分析內容僅供參考，不構成投資建議。
  </p>
</div>
</body>
</html>
"""
    return line_msg, html_body


def _build_history_rows(cb_history: list[dict]) -> str:
    if not cb_history:
        return "<tr><td colspan='3'>此標的尚無歷史可轉債紀錄</td></tr>"
    rows = []
    for h in cb_history[:5]:
        cb_id   = h.get("債券代號", h.get("代號", "N/A"))
        price_r = get_cb_price_range(cb_id) if cb_id and cb_id != "N/A" else {}
        high_s  = f"{price_r['high']:.2f}" if price_r.get("high") else "N/A"
        low_s   = f"{price_r['low']:.2f}"  if price_r.get("low")  else "N/A"
        rows.append(f"<tr><td>{cb_id}</td><td>{high_s}</td><td>{low_s}</td></tr>")
    return "".join(rows)


# ══════════════════════════════════════════════
# 5. 發送 LINE Notify
# ══════════════════════════════════════════════
def send_line_notify(message: str) -> bool:
    if not LINE_NOTIFY_TOKEN:
        print("[跳過] 未設定 LINE_NOTIFY_TOKEN")
        return False
    try:
        resp = requests.post(
            "https://notify-api.line.me/api/notify",
            headers={"Authorization": f"Bearer {LINE_NOTIFY_TOKEN}"},
            data={"message": "\n" + message},
            timeout=15,
        )
        if resp.status_code == 200:
            print("✅ LINE 通知發送成功")
            return True
        else:
            print(f"❌ LINE 通知失敗：{resp.status_code} {resp.text}")
            return False
    except Exception as e:
        print(f"❌ LINE 通知例外：{e}")
        return False


# ══════════════════════════════════════════════
# 6. 發送 Email
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


# ══════════════════════════════════════════════
# 7. 主流程
# ══════════════════════════════════════════════
def main():
    today_str = datetime.now().strftime("%Y/%m/%d")
    print(f"\n{'='*50}")
    print(f"  TWSE 可轉債競拍監控系統  |  {today_str}")
    print(f"{'='*50}\n")

    # 取得即將競拍項目
    upcoming = get_upcoming_auctions()

    if not upcoming:
        print("✅ 未來 5 天內無可轉債競拍，無需通知。")
        return

    print(f"🔍 發現 {len(upcoming)} 筆即將競拍項目，開始分析...\n")

    for idx, auction in enumerate(upcoming, 1):
        stock_name = (
            auction.get("股票名稱")
            or auction.get("公司名稱", "未知")
        )
        stock_code = auction.get("股票代號") or auction.get("stockCode", "")
        days_left  = auction["_days_left"]

        print(f"[{idx}/{len(upcoming)}] 處理：{stock_name}（{stock_code}），距今 {days_left} 天")

        # 查歷史可轉債
        cb_history = fetch_cb_history(stock_code) if stock_code else []
        print(f"    歷史可轉債數量：{len(cb_history)} 筆")

        # AI 分析
        print("    呼叫 AI 分析中...")
        ai_text = ai_analyze(auction, cb_history)

        # 組裝報告
        line_msg, html_body = build_report(auction, cb_history, ai_text)

        # 發送通知
        subject = f"【可轉債競拍預警】{stock_name}（{stock_code}）距今 {days_left} 天 | {today_str}"
        send_line_notify(line_msg)
        send_email(subject, html_body)
        print()

    print("🎉 所有通知發送完畢！")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        raise

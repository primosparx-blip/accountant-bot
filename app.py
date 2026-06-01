"""
PrimoAccountingBot — Telegram Accounting Agent
"""
import os, json, base64, datetime, requests, traceback, threading, time
from flask import Flask, request, jsonify
import anthropic
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

app = Flask(__name__)

TELEGRAM_TOKEN    = os.environ["ACCOUNTING_BOT_TOKEN"]
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]
SHEET_ID          = os.environ.get("SHEET_ID", "1mity1H5znYDITK9QLYORYD-UGt689LmY-fS29m13VLE")
GOOGLE_TOKEN_JSON = os.environ.get("GOOGLE_TOKEN_JSON", "")
REPORT_EMAIL      = os.environ.get("REPORT_EMAIL", "georgejgsolomon@gmail.com")
SENDGRID_API_KEY  = os.environ.get("SENDGRID_API_KEY", "")
SENDGRID_FROM     = os.environ.get("SENDGRID_FROM", "georgejgsolomon@gmail.com")
TELEGRAM_API      = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"

claude = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

EXTRACT_PROMPT = """
You are an accounting AI. Analyze this invoice/bill image and extract fields.
Return ONLY valid JSON — no markdown, no explanation.
{
  "vendor":         "Company name on the invoice",
  "invoice_date":   "YYYY-MM-DD",
  "invoice_number": "Invoice reference number or N/A",
  "description":    "Brief description (max 40 chars)",
  "category":       "One of: Office Supplies | Software/Cloud | Shipping | Facilities | Marketing | Travel | Utilities | Professional Services | Food & Entertainment | Other",
  "currency":       "USD",
  "amount":         123.45,
  "tax":            12.34,
  "confidence":     "high | medium | low"
}
If a numeric field cannot be read use 0. If text cannot be read use Unknown.
"""

SHEET_HEADERS = ["ID","Date Received","Invoice Date","Invoice #","Vendor",
                 "Description","Category","Currency","Amount","Tax","Total","Status","Source"]

def safe_float(val):
    try:
        return float(val or 0)
    except (ValueError, TypeError):
        return 0.0

def get_gspread_client():
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/gmail.modify",
              "https://www.googleapis.com/auth/calendar"]
    if GOOGLE_TOKEN_JSON:
        token_data = json.loads(GOOGLE_TOKEN_JSON)
    else:
        with open("token.json") as f:
            token_data = json.load(f)
    creds = Credentials.from_authorized_user_info(token_data, SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return gspread.authorize(creds)

def get_or_create_sheet(sh, title, headers):
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=len(headers))
        ws.append_row(headers)
    if ws.row_values(1) != headers:
        ws.insert_row(headers, 1)
    return ws

def get_all_invoices():
    gc      = get_gspread_client()
    sh      = gc.open_by_key(SHEET_ID)
    ws      = get_or_create_sheet(sh, "Invoice Log", SHEET_HEADERS)
    records = ws.get_all_records()
    return [r for r in records if r.get("ID","") not in ("", "ID")]

def append_to_sheet(data, source="telegram"):
    gc     = get_gspread_client()
    sh     = gc.open_by_key(SHEET_ID)
    ws     = get_or_create_sheet(sh, "Invoice Log", SHEET_HEADERS)
    rows   = ws.get_all_values()
    inv_id = f"INV-{len(rows):03d}"
    today  = datetime.date.today().strftime("%Y-%m-%d")
    amount = safe_float(data.get("amount", 0))
    tax    = safe_float(data.get("tax", 0))
    ws.append_row([
        inv_id, today,
        data.get("invoice_date", today),
        data.get("invoice_number", "N/A"),
        data.get("vendor", "Unknown"),
        data.get("description", ""),
        data.get("category", "Other"),
        data.get("currency", "USD"),
        amount, tax,
        round(amount + tax, 2),
        "Pending", source
    ])
    refresh_summaries(sh)
    return inv_id

def refresh_summaries(sh):
    records = sh.worksheet("Invoice Log").get_all_records()
    rows    = [r for r in records if r.get("ID","") not in ("","ID")]
    if not rows:
        return

    cat_data = {}
    for r in rows:
        cat = r.get("Category","Other") or "Other"
        amt = safe_float(r.get("Amount", 0))
        tax = safe_float(r.get("Tax", 0))
        if cat not in cat_data:
            cat_data[cat] = {"count":0,"amount":0,"tax":0}
        cat_data[cat]["count"]  += 1
        cat_data[cat]["amount"] += amt
        cat_data[cat]["tax"]    += tax

    total_spend = sum(v["amount"] for v in cat_data.values()) or 1
    cat_headers = ["Category","Invoice Count","Subtotal","Tax","Total","% of Spend"]
    ws_cat = get_or_create_sheet(sh, "Category Summary", cat_headers)
    ws_cat.clear()
    ws_cat.append_row(cat_headers)
    for cat, v in sorted(cat_data.items(), key=lambda x: -x[1]["amount"]):
        pct = round((v["amount"] / total_spend) * 100, 1)
        ws_cat.append_row([cat, v["count"], round(v["amount"],2),
                           round(v["tax"],2), round(v["amount"]+v["tax"],2), f"{pct}%"])

    month_data = {}
    for r in rows:
        month = str(r.get("Invoice Date","") or r.get("Date Received",""))[:7]
        amt = safe_float(r.get("Amount", 0))
        tax = safe_float(r.get("Tax", 0))
        if month not in month_data:
            month_data[month] = {"count":0,"amount":0,"tax":0}
        month_data[month]["count"]  += 1
        month_data[month]["amount"] += amt
        month_data[month]["tax"]    += tax

    month_headers = ["Month","Invoice Count","Subtotal","Tax","Total"]
    ws_m = get_or_create_sheet(sh, "Monthly Summary", month_headers)
    ws_m.clear()
    ws_m.append_row(month_headers)
    for month, v in sorted(month_data.items()):
        ws_m.append_row([month, v["count"], round(v["amount"],2),
                         round(v["tax"],2), round(v["amount"]+v["tax"],2)])

def get_weekly_rows(rows):
    today    = datetime.date.today()
    week_ago = today - datetime.timedelta(days=7)
    result   = []
    for r in rows:
        try:
            d = datetime.date.fromisoformat(str(r.get("Date Received",""))[:10])
            if d >= week_ago:
                result.append(r)
        except:
            pass
    return result

def get_monthly_rows(rows):
    month = datetime.date.today().strftime("%Y-%m")
    return [r for r in rows if str(r.get("Date Received","")).startswith(month)]

def build_summary(rows, label="All Time"):
    if not rows:
        return f"No invoices for {label}."
    total_amt = sum(safe_float(r.get("Amount",0)) for r in rows)
    total_tax = sum(safe_float(r.get("Tax",0)) for r in rows)
    cat_totals = {}
    for r in rows:
        cat = r.get("Category","Other") or "Other"
        cat_totals[cat] = cat_totals.get(cat,0) + safe_float(r.get("Amount",0))
    top = sorted(cat_totals.items(), key=lambda x: -x[1])[:5]
    lines = [
        f"📊 *{label} Summary*",
        f"📋 Invoices: {len(rows)}",
        f"💵 Subtotal: ${total_amt:,.2f}",
        f"🧾 Tax: ${total_tax:,.2f}",
        f"💰 Total: ${total_amt+total_tax:,.2f}",
        "",
        "📂 *Top Categories:*"
    ]
    for cat, amt in top:
        lines.append(f"  • {cat}: ${amt:,.2f}")
    return "\n".join(lines)

def send_weekly_email(rows):
    if not SENDGRID_API_KEY:
        return "not configured"
    weekly      = get_weekly_rows(rows)
    monthly     = get_monthly_rows(rows)
    week_total  = sum(safe_float(r.get("Amount",0))+safe_float(r.get("Tax",0)) for r in weekly)
    month_total = sum(safe_float(r.get("Amount",0))+safe_float(r.get("Tax",0)) for r in monthly)
    cat_totals  = {}
    for r in rows:
        cat = r.get("Category","Other") or "Other"
        cat_totals[cat] = cat_totals.get(cat,0) + safe_float(r.get("Amount",0))
    rows_html = ""
    for i, r in enumerate(weekly):
        bg  = "#F5F5FC" if i%2==0 else "#EAEAF5"
        amt = safe_float(r.get("Amount",0)) + safe_float(r.get("Tax",0))
        rows_html += (f'<tr style="background:{bg}"><td style="padding:8px">{r.get("Vendor","")}</td>'
                      f'<td style="padding:8px">{r.get("Category","")}</td>'
                      f'<td style="padding:8px">{r.get("Date Received","")}</td>'
                      f'<td style="padding:8px">${amt:,.2f}</td></tr>')
    cat_rows = ""
    for cat, amt in sorted(cat_totals.items(), key=lambda x: -x[1]):
        cat_rows += f'<tr><td style="padding:6px">{cat}</td><td style="padding:6px">${amt:,.2f}</td></tr>'
    html = f"""
    <html><body style="font-family:Arial,sans-serif;max-width:620px;margin:auto;padding:20px">
    <div style="background:#0A0A1A;padding:24px;border-radius:8px;margin-bottom:24px">
      <h2 style="color:#00FFB3;margin:0">📊 Weekly Accounting Report</h2>
      <p style="color:#888;margin:8px 0 0">{datetime.date.today().strftime('%B %d, %Y')}</p>
    </div>
    <h3>This Week ({len(weekly)} invoices)</h3>
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #ddd">
    <tr style="background:#1A1A3E;color:white">
      <th style="padding:10px;text-align:left">Vendor</th>
      <th style="padding:10px;text-align:left">Category</th>
      <th style="padding:10px;text-align:left">Date</th>
      <th style="padding:10px;text-align:left">Total</th>
    </tr>{rows_html}</table>
    <p><strong>Week Total: ${week_total:,.2f}</strong></p>
    <hr><h3>Month to Date</h3>
    <p>Invoices: <strong>{len(monthly)}</strong> | Total: <strong>${month_total:,.2f}</strong></p>
    <hr><h3>All Time by Category</h3>
    <table width="60%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #ddd">
    <tr style="background:#1A1A3E;color:white">
      <th style="padding:8px;text-align:left">Category</th>
      <th style="padding:8px;text-align:left">Total</th>
    </tr>{cat_rows}</table>
    <hr><p style="color:#aaa;font-size:12px">Sent every Monday 8am · PrimoAccountingBot</p>
    </body></html>"""
    message = Mail(from_email=SENDGRID_FROM, to_emails=REPORT_EMAIL,
                   subject=f"📊 Weekly Accounting Report — {datetime.date.today().strftime('%b %d, %Y')}",
                   html_content=html)
    SendGridAPIClient(SENDGRID_API_KEY).send(message)
    return "sent"

def weekly_scheduler():
    while True:
        now = datetime.datetime.utcnow()
        if now.weekday() == 0 and now.hour == 12 and now.minute < 5:
            try:
                rows = get_all_invoices()
                send_weekly_email(rows)
            except Exception as e:
                print(f"Scheduler error: {e}")
            time.sleep(360)
        time.sleep(60)

def send_message(chat_id, text, parse_mode="Markdown"):
    requests.post(f"{TELEGRAM_API}/sendMessage",
                  json={"chat_id":chat_id,"text":text,"parse_mode":parse_mode})

def get_file_url(file_id):
    r    = requests.get(f"{TELEGRAM_API}/getFile", params={"file_id":file_id})
    path = r.json()["result"]["file_path"]
    return f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{path}"

def extract_invoice(image_b64, media_type="image/jpeg"):
    message = claude.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=512,
        messages=[{"role":"user","content":[
            {"type":"image","source":{"type":"base64","media_type":media_type,"data":image_b64}},
            {"type":"text","text":EXTRACT_PROMPT}
        ]}]
    )
    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    return json.loads(raw)

@app.route("/telegram", methods=["POST"])
def telegram_webhook():
    update   = request.json
    if not update:
        return "ok"
    msg      = update.get("message", {})
    chat_id  = msg.get("chat", {}).get("id")
    text     = msg.get("text", "").strip().lower()
    photo    = msg.get("photo")
    document = msg.get("document")

    if not chat_id:
        return "ok"

    if photo or (document and document.get("mime_type","").startswith("image")):
        send_message(chat_id, "📸 Got it! Reading your invoice...")
        try:
            file_id   = photo[-1]["file_id"] if photo else document["file_id"]
            file_url  = get_file_url(file_id)
            img_data  = requests.get(file_url).content
            image_b64 = base64.standard_b64encode(img_data).decode("utf-8")
            data      = extract_invoice(image_b64)
            inv_id    = append_to_sheet(data, "telegram")
            conf_emoji = {"high":"✅","medium":"⚠️","low":"🔴"}.get(data.get("confidence","?"),"❓")
            amount = safe_float(data.get("amount",0))
            tax    = safe_float(data.get("tax",0))
            send_message(chat_id,
                f"✅ *Invoice Logged!*\n\n"
                f"🆔 ID: `{inv_id}`\n"
                f"🏢 Vendor: {data.get('vendor','Unknown')}\n"
                f"📅 Date: {data.get('invoice_date','N/A')}\n"
                f"🔢 Invoice \\#: {data.get('invoice_number','N/A')}\n"
                f"📦 Category: {data.get('category','Other')}\n"
                f"💵 Amount: {data.get('currency','TTD')} {amount:,.2f}\n"
                f"🧾 Tax: {data.get('currency','TTD')} {tax:,.2f}\n"
                f"💰 Total: {data.get('currency','TTD')} {amount+tax:,.2f}\n\n"
                f"{conf_emoji} Confidence: {data.get('confidence','?')}\n"
                f"📊 Google Sheet updated!"
            )
        except Exception as e:
            traceback.print_exc()
            send_message(chat_id, f"❌ Error processing invoice: {str(e)[:100]}")
        return "ok"

    try:
        rows = get_all_invoices()
    except:
        rows = []

    if text in ("/start","start","help","/help"):
        send_message(chat_id,
            "👋 *PrimoAccountingBot*\n\n"
            "📸 Send a photo of any invoice or bill\n\n"
            "*Commands:*\n"
            "• /summary — all time totals\n"
            "• /week — this week\n"
            "• /month — this month\n"
            "• /categories — spend by category\n"
            "• /report — send email report now"
        )
    elif text in ("/summary","summary"):
        send_message(chat_id, build_summary(rows, "All Time"))
    elif text in ("/week","week"):
        send_message(chat_id, build_summary(get_weekly_rows(rows), "This Week"))
    elif text in ("/month","month"):
        send_message(chat_id, build_summary(get_monthly_rows(rows), "This Month"))
    elif text in ("/categories","categories"):
        cat_totals = {}
        for r in rows:
            cat = r.get("Category","Other") or "Other"
            cat_totals[cat] = cat_totals.get(cat,0) + safe_float(r.get("Amount",0))
        lines = ["📂 *Spend by Category*\n"]
        for cat, amt in sorted(cat_totals.items(), key=lambda x: -x[1]):
            lines.append(f"• {cat}: ${amt:,.2f}")
        send_message(chat_id, "\n".join(lines))
    elif text in ("/report","report"):
        send_message(chat_id, "📧 Sending email report...")
        result = send_weekly_email(rows)
        send_message(chat_id, f"✅ Report sent to {REPORT_EMAIL}!" if result=="sent" else "❌ Email not configured")
    else:
        send_message(chat_id, "📸 Send a photo of a bill, or use:\n/summary /week /month /categories /report")

    return "ok"

@app.route("/api/summary", methods=["GET"])
def api_summary():
    try:
        rows        = get_all_invoices()
        weekly      = get_weekly_rows(rows)
        monthly     = get_monthly_rows(rows)
        total_all   = sum(safe_float(r.get("Amount",0))+safe_float(r.get("Tax",0)) for r in rows)
        total_week  = sum(safe_float(r.get("Amount",0))+safe_float(r.get("Tax",0)) for r in weekly)
        total_month = sum(safe_float(r.get("Amount",0))+safe_float(r.get("Tax",0)) for r in monthly)
        pending     = [r for r in rows if r.get("Status","") == "Pending"]
        cat_totals  = {}
        for r in rows:
            cat = r.get("Category","Other") or "Other"
            cat_totals[cat] = cat_totals.get(cat,0) + safe_float(r.get("Amount",0))
        top_cats = sorted(cat_totals.items(), key=lambda x: -x[1])[:3]
        return jsonify({
            "total_invoices": len(rows),
            "pending_invoices": len(pending),
            "total_spend_all_time": round(total_all,2),
            "total_spend_this_week": round(total_week,2),
            "total_spend_this_month": round(total_month,2),
            "top_categories": [{"category":c,"amount":round(a,2)} for c,a in top_cats],
            "status": "ok"
        })
    except Exception as e:
        return jsonify({"status":"error","message":str(e)})

@app.route("/", methods=["GET"])
def health():
    return {"status":"ok","bot":"PrimoAccountingBot","time":str(datetime.datetime.now())}

threading.Thread(target=weekly_scheduler, daemon=True).start()

if __name__ == "__main__":
    print("PrimoAccountingBot starting on http://localhost:5001")
    app.run(host="0.0.0.0", port=5001, debug=False)

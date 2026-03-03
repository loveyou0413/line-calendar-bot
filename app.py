import os
import json
import hashlib
import hmac
import base64
import logging
import threading
import time as time_module
from datetime import datetime, timedelta, timezone
from http.server import HTTPServer, BaseHTTPRequestHandler
from io import BytesIO
from urllib.parse import unquote

import anthropic
from google.oauth2 import service_account
from googleapiclient.discovery import build
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

# ===== 設定 =====
LINE_CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
GOOGLE_CALENDAR_ID = os.environ.get("GOOGLE_CALENDAR_ID", "")
GOOGLE_CREDENTIALS_PATH = "/app/credentials.json"
PORT = int(os.environ.get("PORT", "8000"))
REPORT_HOUR = int(os.environ.get("REPORT_HOUR", "18"))
ADMIN_USER_ID = os.environ.get("ADMIN_USER_ID", "")
NAS_EXTERNAL_URL = os.environ.get("NAS_EXTERNAL_URL", "")  # 例如 https://yourname.synology.me:58443
REPORT_DIR = "/app/reports"

# 建立報表資料夾
os.makedirs(REPORT_DIR, exist_ok=True)

# ===== 時區 =====
TW_TZ = timezone(timedelta(hours=8))

# ===== 日誌設定 =====
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)

# ===== 暫存：等待確認的行程 =====
# key: source_id (group or user), value: {event_data, existing_event_id, message_timestamp, expire_time}
pending_confirmations = {}

# ===== 今日新增/修改的行程記錄 =====
daily_event_log = []
daily_event_log_date = None

# ===== Google Calendar =====
def get_calendar_service():
    credentials = service_account.Credentials.from_service_account_file(
        GOOGLE_CREDENTIALS_PATH,
        scopes=["https://www.googleapis.com/auth/calendar"]
    )
    return build("calendar", "v3", credentials=credentials)

# ===== Claude API 解析 =====
def parse_event_with_claude(message_text):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    now = datetime.now(TW_TZ)
    today = now.strftime("%Y-%m-%d")
    current_year = now.strftime("%Y")

    prompt = f"""你是一個行程解析助手。請從以下 Line 群組訊息中擷取行程資訊。
今天的日期是 {today}。

請先判斷這是「一般行程」還是「同仁請假」。
判斷方式：如果訊息中包含「請假」、「休假」、「特休」、「病假」、「事假」、「補休」等關鍵字，就是「同仁請假」類型。

=== 如果是「一般行程」，回傳以下 JSON ===
{{
    "type": "event",
    "title": "行程/會議名稱",
    "date": "YYYY-MM-DD",
    "start_time": "HH:MM",
    "end_time": "HH:MM",
    "location": "完整地點",
    "location_city": "縣市簡稱",
    "host": "主持人姓名",
    "attendees": ["出席人員1", "出席人員2"],
    "staff": "會議幕僚姓名",
    "meeting_url": "線上會議連結",
    "meeting_id": "會議號",
    "meeting_password": "會議密碼",
    "notes": "其他備註資訊"
}}

=== 如果是「同仁請假」，回傳以下 JSON ===
{{
    "type": "leave",
    "dates": ["YYYY-MM-DD"],
    "leaves": [
        {{"name": "同仁姓名", "start_time": "HH:MM", "end_time": "HH:MM"}}
    ]
}}

規則：
1. 如果沒有提到年份，預設使用 {current_year} 年
2. 如果沒有提到結束時間，一般行程預設為開始時間後 1 小時
3. 如果某個欄位訊息中沒有提到，填入 null
4. 日期請務必轉換成 YYYY-MM-DD 格式
5. 時間請務必轉換成 24 小時制 HH:MM 格式
6. 「下週一」「這週五」等相對日期請根據今天日期計算出確切日期
7. 只回傳 JSON，不要有任何其他文字
8. 線上會議資訊：如果訊息中有會議連結、會議號、密碼等資訊，請分別擷取
9. 會議號和密碼可能用「/」分隔（例如「2510 771 7063/2026」表示會議號為「2510 771 7063」，密碼為「2026」）
10. 主持人可能以「主持人」、「主席」、「召集人」等詞彙標示
11. 請假中的「0213」代表 2 月 13 日，「0315」代表 3 月 15 日
12. 請假時間如「(09-1030)」代表 09:00 到 10:30
13. 如果請假跨多天，dates 要列出每一天
14. 一則請假訊息中可能包含多位同仁的請假資訊
15. 出席人員格式規則：
    - 人名後面的(O)代表會出席，(X)代表不出席，請保留這個標記，括號用半形
    - 「幕僚」後面的人名代表會議幕僚，請格式化為「幕僚(人名)」，括號用半形
    - 多位人員用頓號「、」分隔
    - 例如原文「朱理事長(O) 幕僚 貝珊」應整理成 attendees: ["朱理事長(O)", "幕僚(貝珊)"]，staff: "貝珊"
16. location_city 只填縣市簡稱：台北市→台北、新北市→新北、桃園市→桃園、台中市→台中、台南市→台南、高雄市→高雄，以此類推。如果地點是線上或無法判斷縣市則填 null

訊息內容：
{message_text}"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1024,
        messages=[{"role": "user", "content": prompt}]
    )

    response_text = response.content[0].text.strip()
    if response_text.startswith("```"):
        response_text = response_text.split("\n", 1)[1]
        response_text = response_text.rsplit("```", 1)[0]
        response_text = response_text.strip()

    return json.loads(response_text)

# ===== 搜尋重複行程 =====
def find_duplicate_event(service, event_data):
    date_str = event_data["date"]
    time_min = f"{date_str}T00:00:00+08:00"
    time_max = f"{date_str}T23:59:59+08:00"

    events_result = service.events().list(
        calendarId=GOOGLE_CALENDAR_ID,
        timeMin=time_min,
        timeMax=time_max,
        singleEvents=True,
        orderBy="startTime"
    ).execute()

    new_title = (event_data.get("title") or "").strip()
    new_location = (event_data.get("location") or "").strip()

    for evt in events_result.get("items", []):
        if evt.get("summary") == "同仁休假登記":
            continue
        evt_title = (evt.get("summary") or "").strip()
        evt_location = (evt.get("location") or "").strip()

        title_match = new_title and evt_title and new_title == evt_title
        location_match = new_location and evt_location and new_location == evt_location

        if title_match or location_match:
            return evt
    return None

# ===== 建立/更新日曆事件 =====
def build_event_body(event_data, message_timestamp=None):
    description_parts = []

    if event_data.get("host"):
        description_parts.append(f"🎤 主持人：{event_data['host']}")
    if event_data.get("attendees"):
        attendees_str = "、".join(event_data["attendees"])
        description_parts.append(f"📋 出席人員：{attendees_str}")

    has_meeting_info = (
        event_data.get("meeting_url")
        or event_data.get("meeting_id")
        or event_data.get("meeting_password")
    )
    if has_meeting_info:
        description_parts.append("")
        description_parts.append("💻 線上會議資訊：")
        if event_data.get("meeting_url"):
            description_parts.append("會議連結：")
            description_parts.append(event_data["meeting_url"])
        if event_data.get("meeting_id"):
            description_parts.append(f"會議號：{event_data['meeting_id']}")
        if event_data.get("meeting_password"):
            description_parts.append(f"會議密碼：{event_data['meeting_password']}")

    if event_data.get("notes"):
        description_parts.append("")
        description_parts.append(f"📝 備註：{event_data['notes']}")

    if message_timestamp:
        description_parts.append("")
        description_parts.append(f"⏱ 登記時間：{message_timestamp}")

    description = "\n".join(description_parts)

    date_str = event_data["date"]
    start_time = event_data.get("start_time", "09:00")
    end_time = event_data.get("end_time", "10:00")

    event = {
        "summary": event_data.get("title", "未命名行程"),
        "description": description,
        "start": {
            "dateTime": f"{date_str}T{start_time}:00",
            "timeZone": "Asia/Taipei",
        },
        "end": {
            "dateTime": f"{date_str}T{end_time}:00",
            "timeZone": "Asia/Taipei",
        },
    }

    if event_data.get("location"):
        event["location"] = event_data["location"]

    return event

def create_calendar_event(event_data, message_timestamp=None):
    service = get_calendar_service()
    event_body = build_event_body(event_data, message_timestamp)
    return service.events().insert(calendarId=GOOGLE_CALENDAR_ID, body=event_body).execute()

def update_calendar_event(event_id, event_data, message_timestamp=None):
    service = get_calendar_service()
    event_body = build_event_body(event_data, message_timestamp)
    return service.events().update(calendarId=GOOGLE_CALENDAR_ID, eventId=event_id, body=event_body).execute()

# ===== 記錄行程到每日 log =====
def log_daily_event(event_data):
    global daily_event_log, daily_event_log_date

    now = datetime.now(TW_TZ)
    today_str = now.strftime("%Y-%m-%d")

    if daily_event_log_date != today_str:
        daily_event_log = []
        daily_event_log_date = today_str

    date_str = event_data.get("date", "")
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d")
        meeting_date = f"{dt.month}/{dt.day}"
    except ValueError:
        meeting_date = date_str

    # 出席人員：排除幕僚項目
    attendees = event_data.get("attendees") or []
    non_staff_attendees = [a for a in attendees if not a.startswith("幕僚")]
    attendees_str = "、".join(non_staff_attendees) if non_staff_attendees else ""

    daily_event_log.append({
        "meeting_date": meeting_date,
        "staff": event_data.get("staff") or "",
        "location_city": event_data.get("location_city") or "",
        "title": event_data.get("title") or "未命名行程",
        "attendees": attendees_str,
    })

# ===== 請假整合 =====
def find_leave_event(service, date_str):
    time_min = f"{date_str}T00:00:00+08:00"
    time_max = f"{date_str}T23:59:59+08:00"

    events_result = service.events().list(
        calendarId=GOOGLE_CALENDAR_ID,
        timeMin=time_min,
        timeMax=time_max,
        q="同仁休假登記",
        singleEvents=True,
        orderBy="startTime"
    ).execute()

    for evt in events_result.get("items", []):
        if evt.get("summary") == "同仁休假登記":
            return evt
    return None

def parse_existing_leaves(description):
    leaves = []
    if not description:
        return leaves
    for line in description.strip().split("\n"):
        line = line.strip()
        if line and line[0].isdigit() and "." in line:
            content = line.split(".", 1)[1]
            leaves.append(content)
    return leaves

def format_leave_description(leave_entries):
    return "\n".join(f"{i}.{entry}" for i, entry in enumerate(leave_entries, 1))

def handle_leave(event_data):
    service = get_calendar_service()
    results = []
    dates = event_data.get("dates", [])
    leaves = event_data.get("leaves", [])

    for date_str in dates:
        existing_event = find_leave_event(service, date_str)

        new_entries = []
        for leave in leaves:
            name = leave.get("name", "未知")
            start = leave.get("start_time", "").replace(":", "")
            end = leave.get("end_time", "").replace(":", "")
            if start and end:
                new_entries.append(f"{name}請休{start}-{end}")
            else:
                new_entries.append(f"{name}請休（時間未指定）")

        if existing_event:
            existing_desc = existing_event.get("description", "")
            existing_leaves = parse_existing_leaves(existing_desc)

            for new_entry in new_entries:
                new_name = new_entry.split("請休")[0]
                existing_leaves = [e for e in existing_leaves if not e.startswith(f"{new_name}請休")]
                existing_leaves.append(new_entry)

            existing_event["description"] = format_leave_description(existing_leaves)
            updated = service.events().update(
                calendarId=GOOGLE_CALENDAR_ID,
                eventId=existing_event["id"],
                body=existing_event
            ).execute()

            results.append({"date": date_str, "action": "updated", "entries": existing_leaves})
        else:
            description = format_leave_description(new_entries)
            end_date = (datetime.strptime(date_str, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")

            event = {
                "summary": "同仁休假登記",
                "description": description,
                "start": {"date": date_str},
                "end": {"date": end_date},
            }
            service.events().insert(calendarId=GOOGLE_CALENDAR_ID, body=event).execute()
            results.append({"date": date_str, "action": "created", "entries": new_entries})

    return results

# ===== Line 訊息 =====
def reply_message(reply_token, text):
    url = "https://api.line.me/v2/bot/message/reply"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"
    }
    body = {
        "replyToken": reply_token,
        "messages": [{"type": "text", "text": text}]
    }
    resp = requests.post(url, headers=headers, json=body)
    if resp.status_code != 200:
        logger.error(f"Line reply failed: {resp.status_code} {resp.text}")

def push_message(user_id, text):
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"
    }
    body = {
        "to": user_id,
        "messages": [{"type": "text", "text": text}]
    }
    resp = requests.post(url, headers=headers, json=body)
    if resp.status_code != 200:
        logger.error(f"Line push failed: {resp.status_code} {resp.text}")

def verify_signature(body, signature):
    hash_value = hmac.new(
        LINE_CHANNEL_SECRET.encode("utf-8"),
        body,
        hashlib.sha256
    ).digest()
    expected_signature = base64.b64encode(hash_value).decode("utf-8")
    return hmac.compare_digest(expected_signature, signature)

# ===== 格式化確認訊息 =====
def format_event_confirmation(event_data, action="登記"):
    lines = [f"✅ 行程已{action}！", ""]
    lines.append(f"📌 {event_data.get('title', '未命名行程')}")
    lines.append(f"📅 {event_data.get('date', '未指定')}")

    start = event_data.get("start_time", "")
    end = event_data.get("end_time", "")
    if start and end:
        lines.append(f"🕐 {start} ~ {end}")
    elif start:
        lines.append(f"🕐 {start}")

    if event_data.get("location"):
        lines.append(f"📍 {event_data['location']}")
    if event_data.get("host"):
        lines.append(f"🎤 主持人：{event_data['host']}")
    if event_data.get("attendees"):
        lines.append(f"👥 {'、'.join(event_data['attendees'])}")

    has_meeting_info = (
        event_data.get("meeting_url")
        or event_data.get("meeting_id")
        or event_data.get("meeting_password")
    )
    if has_meeting_info:
        lines.append("")
        lines.append("💻 線上會議資訊：")
        if event_data.get("meeting_url"):
            lines.append("會議連結：")
            lines.append(event_data["meeting_url"])
        if event_data.get("meeting_id"):
            lines.append(f"會議號：{event_data['meeting_id']}")
        if event_data.get("meeting_password"):
            lines.append(f"會議密碼：{event_data['meeting_password']}")

    if event_data.get("notes"):
        lines.append(f"📝 {event_data['notes']}")

    return "\n".join(lines)

def format_leave_confirmation(results):
    lines = ["✅ 休假已登記！", ""]
    for result in results:
        action = "已更新" if result["action"] == "updated" else "已新增"
        lines.append(f"📅 {result['date']}（{action}）")
        for entry in result["entries"]:
            lines.append(f"  • {entry}")
        lines.append("")
    return "\n".join(lines).strip()

# ===== 每日報表 =====
def generate_daily_report_excel():
    global daily_event_log

    wb = Workbook()
    ws = wb.active
    now = datetime.now(TW_TZ)
    ws.title = "Report"

    headers = ["會議時間", "幕僚", "會議地點", "會議名稱", "出席人員"]

    # 寫入標題列（只設定粗體，不加邊框避免相容性問題）
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)

    if daily_event_log:
        for row_idx, entry in enumerate(daily_event_log, 2):
            ws.cell(row=row_idx, column=1, value=entry["meeting_date"])
            ws.cell(row=row_idx, column=2, value=entry["staff"])
            ws.cell(row=row_idx, column=3, value=entry["location_city"])
            ws.cell(row=row_idx, column=4, value=entry["title"])
            ws.cell(row=row_idx, column=5, value=entry["attendees"])

    # 設定欄寬
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 30

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def send_daily_report():
    """每日報表：產生 Excel 並傳送下載連結到管理員私訊"""
    if not ADMIN_USER_ID:
        logger.warning("ADMIN_USER_ID not set, skipping daily report")
        return

    global daily_event_log

    now = datetime.now(TW_TZ)
    date_str = now.strftime("%m/%d")

    if not daily_event_log:
        push_message(ADMIN_USER_ID, f"📊 今日行程更新報表（{date_str}）\n\n今日無新增或修改的行程。")
        daily_event_log = []
        logger.info("Daily report sent (no events)")
        return

    # 產生 Excel 並存檔
    try:
        excel_buf = generate_daily_report_excel()
        filename = f"report_{now.strftime('%Y%m%d')}.xlsx"
        filepath = os.path.join(REPORT_DIR, filename)

        with open(filepath, "wb") as f:
            f.write(excel_buf.read())

        logger.info(f"Excel report saved: {filepath}")

        # 清理 7 天前的舊報表
        for old_file in os.listdir(REPORT_DIR):
            old_path = os.path.join(REPORT_DIR, old_file)
            if os.path.isfile(old_path):
                file_age = now.timestamp() - os.path.getmtime(old_path)
                if file_age > 7 * 86400:
                    os.remove(old_path)
                    logger.info(f"Deleted old report: {old_file}")

        # 組合訊息
        text_lines = [f"📊 今日行程更新報表（{date_str}）", ""]
        text_lines.append(f"共 {len(daily_event_log)} 筆行程更新")

        if NAS_EXTERNAL_URL:
            download_url = f"{NAS_EXTERNAL_URL}/reports/{filename}"
            text_lines.append("")
            text_lines.append(f"📥 下載 Excel 報表：")
            text_lines.append(download_url)
        else:
            # 如果沒設定外部網址，改傳文字摘要
            text_lines.append("")
            for entry in daily_event_log:
                text_lines.append(
                    f"• {entry['meeting_date']} | {entry['staff'] or '-'} | "
                    f"{entry['location_city'] or '-'} | {entry['title']} | "
                    f"{entry['attendees'] or '-'}"
                )

        push_message(ADMIN_USER_ID, "\n".join(text_lines))

    except Exception as e:
        logger.error(f"Daily report error: {e}")
        push_message(ADMIN_USER_ID, f"⚠️ 今日報表產生失敗：{str(e)}")

    # 清空 log
    daily_event_log = []
    logger.info("Daily report sent, log cleared")

def report_scheduler():
    """每日報表排程器"""
    while True:
        now = datetime.now(TW_TZ)
        target = now.replace(hour=REPORT_HOUR, minute=0, second=0, microsecond=0)
        if now >= target:
            target += timedelta(days=1)

        wait_seconds = (target - now).total_seconds()
        logger.info(f"Next daily report at {target.strftime('%Y-%m-%d %H:%M')} ({int(wait_seconds)}s)")
        time_module.sleep(wait_seconds)

        try:
            send_daily_report()
        except Exception as e:
            logger.error(f"Report scheduler error: {e}")

# ===== 取得訊息時間戳 =====
def get_message_timestamp(event):
    """從 Line 事件中取得訊息時間，轉為台灣時間字串"""
    ts = event.get("timestamp", 0)
    if ts:
        dt = datetime.fromtimestamp(ts / 1000, tz=TW_TZ)
        return dt.strftime("%Y年%m月%d日 %H:%M")
    return datetime.now(TW_TZ).strftime("%Y年%m月%d日 %H:%M")

# ===== 取得來源 ID =====
def get_source_id(event):
    """取得訊息來源的 ID（群組或個人）"""
    source = event.get("source", {})
    return source.get("groupId") or source.get("roomId") or source.get("userId", "unknown")

# ===== 清理過期的 pending confirmations =====
def cleanup_pending():
    now = datetime.now(TW_TZ)
    expired = [k for k, v in pending_confirmations.items() if now > v.get("expire_time", now)]
    for k in expired:
        del pending_confirmations[k]

# ===== HTTP Handler =====
class WebhookHandler(BaseHTTPRequestHandler):

    def do_POST(self):
        logger.info(f"POST request received: {self.path} from {self.client_address}")

        if self.path != "/callback":
            self.send_response(404)
            self.end_headers()
            return

        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length)

        signature = self.headers.get("X-Line-Signature", "")
        if not verify_signature(body, signature):
            logger.warning("Invalid signature")
            self.send_response(403)
            self.end_headers()
            return

        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(b'{"status":"ok"}')

        try:
            data = json.loads(body.decode("utf-8"))
            for event in data.get("events", []):
                self.handle_event(event)
        except Exception as e:
            logger.error(f"Error processing webhook: {e}")

    def do_GET(self):
        logger.info(f"GET request received: {self.path} from {self.client_address}")

        # URL 解碼
        decoded_path = unquote(self.path)

        # 報表下載
        if decoded_path.startswith("/reports/"):
            self.serve_report()
            return

        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write("LINE Calendar Bot is running!".encode("utf-8"))

    def serve_report(self):
        """提供報表檔案下載"""
        # URL 解碼（處理中文或特殊字元被編碼的情況）
        decoded_path = unquote(self.path)
        filename = decoded_path.split("/reports/", 1)[1]

        # 安全檢查：防止路徑穿越
        if "/" in filename or "\\" in filename or ".." in filename:
            self.send_response(403)
            self.end_headers()
            return

        filepath = os.path.join(REPORT_DIR, filename)
        logger.info(f"Report download request: {filename} -> {filepath}")

        # 列出目前 reports 資料夾裡的檔案（除錯用）
        try:
            existing_files = os.listdir(REPORT_DIR)
            logger.info(f"Files in reports dir: {existing_files}")
        except Exception as e:
            logger.error(f"Cannot list reports dir: {e}")

        if not os.path.exists(filepath):
            self.send_response(404)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write("Report not found".encode("utf-8"))
            return

        with open(filepath, "rb") as f:
            data = f.read()

        logger.info(f"Serving report: {filename} ({len(data)} bytes)")
        self.send_response(200)
        self.send_header("Content-Type", "application/octet-stream")
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(data)

    def handle_event(self, event):
        if event.get("type") != "message":
            return
        if event.get("message", {}).get("type") != "text":
            return

        message_text = event["message"]["text"].strip()
        reply_token = event["replyToken"]
        source_id = get_source_id(event)

        # 清理過期的確認
        cleanup_pending()

        # 檢查是否是回覆確認/新增
        if source_id in pending_confirmations:
            if message_text == "確認":
                self.handle_confirm_replace(source_id, reply_token)
                return
            elif message_text == "新增":
                self.handle_confirm_add(source_id, reply_token)
                return

        # 只處理以「登記」開頭的訊息
        if not message_text.startswith("登記"):
            return

        message_text = message_text[2:].strip()
        if not message_text:
            reply_message(reply_token,
                "請在「登記」後面加上行程資訊，例如：\n"
                "登記 3月15日下午兩點，A棟301會議室，產品規劃會議\n\n"
                "或請假資訊：\n"
                "登記 0213同仁請假 麗如1.5小時(09-1030)")
            return

        logger.info(f"Received message: {message_text}")
        message_timestamp = get_message_timestamp(event)

        try:
            parsed_data = parse_event_with_claude(message_text)
            logger.info(f"Parsed: {json.dumps(parsed_data, ensure_ascii=False)}")

            msg_type = parsed_data.get("type", "event")

            if msg_type == "leave":
                if not parsed_data.get("dates") or not parsed_data.get("leaves"):
                    reply_message(reply_token, "⚠️ 無法從訊息中辨識出請假資訊，請確認訊息中有包含日期和請假人員。")
                    return
                results = handle_leave(parsed_data)
                logger.info(f"Leave processed: {len(results)}")
                reply_message(reply_token, format_leave_confirmation(results))

            else:
                if not parsed_data.get("date") or not parsed_data.get("start_time"):
                    reply_message(reply_token, "⚠️ 無法從訊息中辨識出日期或時間，請確認訊息中有包含行程的日期和時間。")
                    return

                # 檢查重複
                service = get_calendar_service()
                duplicate = find_duplicate_event(service, parsed_data)

                if duplicate:
                    # 暫存，等待確認
                    pending_confirmations[source_id] = {
                        "event_data": parsed_data,
                        "existing_event_id": duplicate["id"],
                        "message_timestamp": message_timestamp,
                        "expire_time": datetime.now(TW_TZ) + timedelta(minutes=10),
                    }

                    dup_title = duplicate.get("summary", "未命名")
                    dup_start = ""
                    if duplicate.get("start", {}).get("dateTime"):
                        try:
                            dt = datetime.fromisoformat(duplicate["start"]["dateTime"])
                            dup_start = dt.strftime("%H:%M")
                        except Exception:
                            pass
                    dup_end = ""
                    if duplicate.get("end", {}).get("dateTime"):
                        try:
                            dt = datetime.fromisoformat(duplicate["end"]["dateTime"])
                            dup_end = dt.strftime("%H:%M")
                        except Exception:
                            pass

                    time_str = f"{dup_start}~{dup_end}" if dup_start and dup_end else ""

                    reply_message(reply_token,
                        f"⚠️ 發現同一天已有類似行程：\n"
                        f"📌 {dup_title}\n"
                        f"🕐 {time_str}\n\n"
                        f"請回覆：\n"
                        f"「確認」→ 替換舊行程\n"
                        f"「新增」→ 另外新增一筆")
                    return

                # 沒有重複，直接建立
                created = create_calendar_event(parsed_data, message_timestamp)
                logger.info(f"Event created: {created.get('id')}")
                log_daily_event(parsed_data)
                reply_message(reply_token, format_event_confirmation(parsed_data))

        except json.JSONDecodeError as e:
            logger.error(f"JSON parse error: {e}")
            reply_message(reply_token, "⚠️ 解析行程資訊時發生錯誤，請確認訊息格式是否包含行程相關資訊。")
        except Exception as e:
            logger.error(f"Error: {e}")
            reply_message(reply_token, f"⚠️ 發生錯誤：{str(e)}")

    def handle_confirm_replace(self, source_id, reply_token):
        """處理確認替換"""
        pending = pending_confirmations.pop(source_id, None)
        if not pending:
            reply_message(reply_token, "⚠️ 找不到待確認的行程，請重新登記。")
            return

        try:
            event_data = pending["event_data"]
            event_id = pending["existing_event_id"]
            message_timestamp = pending["message_timestamp"]

            updated = update_calendar_event(event_id, event_data, message_timestamp)
            logger.info(f"Event replaced: {updated.get('id')}")
            log_daily_event(event_data)
            reply_message(reply_token, format_event_confirmation(event_data, action="替換"))
        except Exception as e:
            logger.error(f"Replace error: {e}")
            reply_message(reply_token, f"⚠️ 替換行程時發生錯誤：{str(e)}")

    def handle_confirm_add(self, source_id, reply_token):
        """處理確認新增"""
        pending = pending_confirmations.pop(source_id, None)
        if not pending:
            reply_message(reply_token, "⚠️ 找不到待確認的行程，請重新登記。")
            return

        try:
            event_data = pending["event_data"]
            message_timestamp = pending["message_timestamp"]

            created = create_calendar_event(event_data, message_timestamp)
            logger.info(f"Event added: {created.get('id')}")
            log_daily_event(event_data)
            reply_message(reply_token, format_event_confirmation(event_data))
        except Exception as e:
            logger.error(f"Add error: {e}")
            reply_message(reply_token, f"⚠️ 新增行程時發生錯誤：{str(e)}")

    def log_message(self, format, *args):
        pass

# ===== 啟動 =====
if __name__ == "__main__":
    logger.info(f"Starting LINE Calendar Bot on port {PORT}...")

    # 啟動每日報表排程
    report_thread = threading.Thread(target=report_scheduler, daemon=True)
    report_thread.start()
    logger.info(f"Daily report scheduler started (send at {REPORT_HOUR}:00)")

    server = HTTPServer(("0.0.0.0", PORT), WebhookHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        logger.info("Shutting down...")
        server.server_close()

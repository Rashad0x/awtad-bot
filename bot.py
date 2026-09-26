import os
import time
import json
import requests
from datetime import datetime
import openpyxl

CONFIG_FILE = "config.json"
EXCEL_FILE = "posts.xlsx"
GRAPH_URL = "https://graph.facebook.com/v19.0"

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

config = load_config()

def sync_excel():
    drive_url = config.get("excel_drive_download_url", "")
    if not drive_url:
        print("❌ رابط الدرايف غير موجود في الإعدادات")
        return False
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(drive_url, headers=headers)
        if response.status_code == 200:
            with open(EXCEL_FILE, 'wb') as f:
                f.write(response.content)
            print("✅ تم تحديث الشيت بنجاح")
            return True
    except Exception as e:
        print("خطأ في المزامنة:", e)
    return False

def parse_excel_datetime(val):
    if not val:
        return None
    if isinstance(val, datetime):
        return int(val.timestamp())
    val_str = str(val).strip()
    if not val_str:
        return None

    date_formats = [
        "%d/%m/%Y %I:%M %p",
        "%d/%m/%Y %I:%M:%S %p",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S"
    ]
    for fmt in date_formats:
        try:
            dt = datetime.strptime(val_str, fmt)
            return int(dt.timestamp())
        except ValueError:
            continue
    return None

def post_to_facebook(caption, schedule_timestamp=None):
    page_id = config.get('fb_page_id', '')
    token = config.get('access_token', '')
    if not page_id or not token:
        return {"error": "بيانات فيسبوك مفقودة"}
    
    url = f"{GRAPH_URL}/{page_id}/feed"
    payload = {
        'access_token': token,
        'message': caption
    }
    if schedule_timestamp:
        payload['published'] = 'false'
        payload['scheduled_publish_time'] = schedule_timestamp
        payload['unpublished_content_type'] = 'SCHEDULED'
    else:
        payload['published'] = 'true'

    res = requests.post(url, data=payload).json()
    return res

def main():
    print("--- بدء فحص السيرفر السحابي للبوستات ---")
    if not sync_excel():
        return

    if not os.path.exists(EXCEL_FILE):
        return

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    sheet = wb.active
    now_ts = int(time.time())

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            p_num = str(row[0]).strip()
            fb_cap = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
            raw_time = row[3] if len(row) > 3 else None
            st_time = parse_excel_datetime(raw_time)

            # لو البوست ميعاده جه أو تم تحديده كجدولة مستقبلية صحيحة
            if st_time and st_time > now_ts:
                # إرسال أمر الجدولة لفيسبوك
                res = post_to_facebook(fb_cap, st_time)
                if 'id' in res:
                    print(f"✅ تم جدولة بوست #{p_num} بنجاح على فيسبوك سحابياً!")
                else:
                    print(f"❌ خطأ في بوست #{p_num}:", res)

if __name__ == "__main__":
    main()

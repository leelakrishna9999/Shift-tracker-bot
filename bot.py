# Shift-tracker-bot
import logging
import re
import math
import requests
from datetime import datetime, timedelta
from telegram import Update
from telegram.ext import Application, MessageHandler, filters, ContextTypes
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import shutil
from pathlib import Path

BOT_TOKEN = "8792621305:AAFu0aKuak02DgPOoEwp3sLgpaRrmx4CMVU"
YOUR_USER_ID = 8293531859
YOUR_NAME = "Leelakrishna"
ONEDRIVE_PATH = str(Path.home() / "OneDrive" / "Desktop" / "WorkHoursTracker.xlsx")
HOME_LAT = 54.5973
HOME_LNG = -5.9301
HOME_ADDRESS = "16 Oranmore Street, Belfast, BT13 2RU"
TRAVEL_METHOD = "Bus"
BUS_SPEED_KMH = 20
BUS_WAITING_MINS = 10
WALK_TO_STOP_MINS = 5
WAKE_UP_BUFFER_MINS = 90
LEAVE_HOME_BUFFER_MINS = 15

PAY_RATES_PRE_APRIL = {
    "weekday_day": 12.50,
    "weekday_night": 13.00,
    "saturday": 13.00,
    "sunday": 15.00,
    "bank_holiday": 15.00,
}

PAY_RATES_APRIL = {
    "weekday_day": 13.00,
    "weekday_night": 14.00,
    "saturday": 14.00,
    "sunday": 16.00,
    "bank_holiday": 18.00,
}

BANK_HOLIDAYS = [
    "2026-01-01", "2026-03-17", "2026-04-10", "2026-04-13",
    "2026-05-04", "2026-05-25", "2026-07-12", "2026-08-31",
    "2026-12-25", "2026-12-28"
]

SHIFT_TYPES = {
    "long day": {"start": 8, "end": 20, "break": 1.0},
    "night duty": {"start": 20, "end": 8, "break": 1.0},
    "night": {"start": 20, "end": 8, "break": 1.0},
    "early": {"start": 6, "end": 14, "break": 1.0},
    "late": {"start": 14, "end": 22, "break": 1.0},
    "short day": {"start": 8, "end": 16, "break": 0.5},
}

logging.basicConfig(level=logging.INFO)

def geocode_location(address):
    try:
        url = "https://nominatim.openstreetmap.org/search"
        params = {"q": address, "format": "json", "limit": 1, "countrycodes": "gb"}
        headers = {"User-Agent": "ShiftTrackerBot/1.0"}
        response = requests.get(url, params=params, headers=headers, timeout=10)
        data = response.json()
        if data:
            return float(data[0]["lat"]), float(data[0]["lon"]), data[0]["display_name"]
        return None, None, address
    except:
        return None, None, address

def calculate_distance(lat1, lng1, lat2, lng2):
    R = 6371
    dlat = math.radians(lat2 - lat1)
    dlng = math.radians(lng2 - lng1)
    a = math.sin(dlat/2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlng/2)**2
    return round(R * 2 * math.asin(math.sqrt(a)), 1)

def estimate_travel_time(distance_km):
    bus_ride_mins = int((distance_km / BUS_SPEED_KMH) * 60)
    return WALK_TO_STOP_MINS + BUS_WAITING_MINS + bus_ride_mins

def get_timings(shift_date, shift_start_hour, travel_mins):
    shift_start = shift_date.replace(hour=shift_start_hour, minute=0, second=0, microsecond=0)
    leave_home = shift_start - timedelta(minutes=travel_mins + LEAVE_HOME_BUFFER_MINS)
    wake_up = leave_home - timedelta(minutes=WAKE_UP_BUFFER_MINS)
    return wake_up.strftime("%H:%M"), leave_home.strftime("%H:%M"), shift_start.strftime("%H:%M")

def get_active_rates(shift_date):
    april_2026 = datetime(2026, 4, 1)
    if shift_date < april_2026:
        return PAY_RATES_PRE_APRIL, "Feb/Mar 2026"
    else:
        return PAY_RATES_APRIL, "April 2026"

def get_pay_rate(date_obj, shift_type):
    rates, period = get_active_rates(date_obj)
    date_str = date_obj.strftime("%Y-%m-%d")
    if date_str in BANK_HOLIDAYS:
        return rates["bank_holiday"], f"Bank Holiday ({period})"
    weekday = date_obj.weekday()
    if weekday == 6:
        return rates["sunday"], f"Sunday ({period})"
    if weekday == 5:
        return rates["saturday"], f"Saturday ({period})"
    if "night" in shift_type.lower():
        return rates["weekday_night"], f"Weekday Night ({period})"
    return rates["weekday_day"], f"Weekday Day ({period})"

def calculate_hours(shift_type):
    for key, val in SHIFT_TYPES.items():
        if key in shift_type.lower():
            raw = val["end"] - val["start"]
            if raw <= 0:
                raw += 24
            return raw - val["break"], val["start"], val["end"], val["break"]
    return 11, 8, 20, 1.0

def parse_shift_message(text):
    result = {}
    months = {"january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
              "july":7,"august":8,"september":9,"october":10,"november":11,"december":12}
    
    date_match = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s+of\s+(\w+)', text, re.IGNORECASE)
    if not date_match:
        date_match = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s+(\w+)', text, re.IGNORECASE)
    if date_match:
        day = int(date_match.group(1))
        month_str = date_match.group(2).lower()
        month = months.get(month_str, datetime.now().month)
        try:
            result["date"] = datetime(2026, month, day)
        except:
            result["date"] = datetime.now()
    else:
        result["date"] = datetime.now()

    lines = [l.strip() for l in text.split('\n') if l.strip()]
    for i, line in enumerate(lines):
        if re.search(r'shift', line, re.IGNORECASE) and i + 1 < len(lines):
            result["location_name"] = lines[i + 1]
            break
    if "location_name" not in result:
        result["location_name"] = "Unknown"

    postcode_match = re.search(r'\b[A-Z]{1,2}\d{1,2}\s?\d[A-Z]{2}\b', text, re.IGNORECASE)
    result["postcode"] = postcode_match.group(0).upper() if postcode_match else ""
    result["full_address"] = f"{result['location_name']}, {result['postcode']}, NI" if result["postcode"] else f"{result['location_name']}, Belfast"

    unit_match = re.search(r'(\w+)\s+[Uu]nit', text)
    result["unit"] = unit_match.group(0) if unit_match else "Not specified"

    result["shift_type"] = "Long Day"
    for key in SHIFT_TYPES.keys():
        if key in text.lower():
            result["shift_type"] = key.title()
            break

    return result

def update_excel(parsed, hours, rate, earnings, day_type, distance_km, travel_mins, wake_up, leave_home):
    month_year = parsed['date'].strftime("%B %Y")
    
    if os.path.exists(ONEDRIVE_PATH):
        wb = openpyxl.load_workbook(ONEDRIVE_PATH)
        if month_year not in wb.sheetnames:
            ws = wb.create_sheet(month_year)
        else:
            ws = wb[month_year]
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = month_year

    if ws.max_row == 1:
        headers = ["#", "Date", "Day", "Location", "Unit", "Shift Type", "Day Type", "Start", "End", "Break", "Hours", "Rate", "Earnings", "Distance", "Travel", "Wake Up", "Leave Home"]
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="0D2137")
            c.alignment = Alignment(horizontal="center")

    next_row = ws.max_row + 1
    _, start_hr, end_hr, break_hrs = calculate_hours(parsed["shift_type"])
    values = [next_row-1, parsed['date'].strftime("%d/%m/%Y"), parsed['date'].strftime("%A"),
              parsed['location_name'], parsed['unit'], parsed['shift_type'],
              day_type, f"{start_hr:02d}:00", f"{end_hr:02d}:00", break_hrs, hours, rate, earnings, distance_km, travel_mins, wake_up, leave_home]
    
    for col, val in enumerate(values, 1):
        ws.cell(row=next_row, column=col, value=val)

    wb.save(ONEDRIVE_PATH)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != YOUR_USER_ID:
        return

    text = update.message.text.strip().lower()

    if "confirmed shift" in text or "new shift" in text:
        await update.message.reply_text("📋 Logging shift...")
        try:
            parsed = parse_shift_message(update.message.text)
            work_lat, work_lng, _ = geocode_location(parsed["full_address"])
            if work_lat and work_lng:
                distance_km = calculate_distance(HOME_LAT, HOME_LNG, work_lat, work_lng)
            else:
                distance_km = 10.0
            
            travel_mins = estimate_travel_time(distance_km)
            hours, start_hr, end_hr, break_hrs = calculate_hours(parsed["shift_type"])
            rate, day_type = get_pay_rate(parsed["date"], parsed["shift_type"])
            earnings = round(hours * rate, 2)
            wake_up, leave_home, shift_start = get_timings(parsed["date"], start_hr, travel_mins)
            
            update_excel(parsed, hours, rate, earnings, day_type, distance_km, travel_mins, wake_up, leave_home)
            
            msg = f"""✅ LOGGED
{parsed['date'].strftime('%d %B %Y')} ({parsed['date'].strftime('%A')})
{parsed['location_name']} | {parsed['unit']}
{parsed['shift_type']} | {hours}hrs | £{earnings:.2f}
🚌 {distance_km}km | {travel_mins}mins
⏰ Wake: {wake_up} | Leave: {leave_home}
📁 Saved to OneDrive"""
            
            await update.message.reply_text(msg)
        except Exception as e:
            await update.message.reply_text(f"❌ Error: {str(e)}")
    else:
        await update.message.reply_text(f"Hi {YOUR_NAME}! Send shift confirmation to log hours.")

def main():
    print("✅ Bot running on Railway - saving to OneDrive")
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.run_polling()

if __name__ == "__main__":
    main()

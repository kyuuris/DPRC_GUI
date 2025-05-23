import tkinter as tk
from tkinter import messagebox
import subprocess
import os
import time
import webbrowser
import win32com.client
import win32com.client
import calendar

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta, date

def get_calendar_folder(advisor_name):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        return outlook.Folders.Item("dprc@sfsu.edu").Folders.Item("Calendar").Folders.Item(advisor_name)
    except Exception as e:
        print(f"❌ Could not access calendar for {advisor_name}: {e}")
        return None
    
def is_ooo_or_vacation(calendar_items, date):
    start_time = datetime.combine(date, datetime.min.time()).replace(hour=8)
    end_time = datetime.combine(date, datetime.min.time()).replace(hour=18)

    restriction = (
        f"[Start] <= '{end_time.strftime('%m/%d/%Y %I:%M %p')}' AND "
        f"[End] >= '{start_time.strftime('%m/%d/%Y %I:%M %p')}'"
    )

    try:
        filtered_items = calendar_items.Restrict(restriction)
        for item in filtered_items:
            if item.AllDayEvent and hasattr(item, "Subject"):
                subject = item.Subject.lower()
                if "ooo" in subject:
                    return True
    except Exception as e:
        print(f"⚠️ Failed OOO check: {e}")

    return False

def get_busy_slots_for_day(advisor_name, day):
    calendar_folder = get_calendar_folder(advisor_name)
    if not calendar_folder:
        return []

    calendar_items = calendar_folder.Items
    calendar_items.IncludeRecurrences = True
    calendar_items.Sort("[Start]")

    day_start = datetime.combine(day, datetime.min.time()).replace(hour=8)
    day_end = datetime.combine(day, datetime.min.time()).replace(hour=18)

    busy_times = []
    try:
        for item in calendar_items:
            try:
                start = item.Start.replace(tzinfo=None)
                end = item.End.replace(tzinfo=None)
                if start < day_end and end > day_start:
                    busy_times.append((start, end))
            except Exception as e:
                print(f"⛔ Error reading event for {advisor_name}: {e}")
    except Exception as e:
        print(f"❌ Failed to loop calendar for {advisor_name}: {e}")

    print(f"✅ {advisor_name} - {len(busy_times)} actual events on {day.strftime('%A, %B %d')}")
    return busy_times

def get_busy_times_by_person(usernames, start_time, end_time):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    root_folder = outlook.Folders.Item("dprc@sfsu.edu").Folders["Calendar"]
    all_busy = {}


    for name in usernames:
        busy_times = []
        try:
            calendar = root_folder.Folders[name]
            items = calendar.Items
            items.IncludeRecurrences = True
            items.Sort("[Start]")

            restriction = f"[Start] >= '{start_time.strftime('%m/%d/%Y %I:%M %p')}' AND [End] <= '{end_time.strftime('%m/%d/%Y %I:%M %p')}'"
            restricted_items = items.Restrict(restriction)

            if is_ooo_or_vacation(items, start_time.date()):
               continue

            for item in restricted_items:
                try:
                    start = item.Start.replace(tzinfo=None)
                    end = item.End.replace(tzinfo=None)

                    sensitivity = getattr(item, "Sensitivity", 0)  # 2 = Private
                    is_private = sensitivity == 2

                    if start.year < 2100:
                        busy_times.append((start, end))

                except Exception as e:
                    print(f"⚠️ Skipped an item for {name}: {e}")

        except Exception as e:
            print(f"⚠️ Could not access {name}'s calendar: {e}")
        all_busy[name] = busy_times

    return all_busy

def get_availability_text():
    def is_weekend(date):
        return date.weekday() >= 5
    
    def load_ignored_slots(file_path="ignored_slots.txt"):
        ignored = set()
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#"):
                        try:
                            m, d, h = map(int, line.split())
                            if h in [1, 2, 3, 4]:
                                h += 12
                            ignored.add((m, d, h))
                        except:
                            continue
        except FileNotFoundError:
            pass
        return ignored

    def get_flexible_free_blocks(busy_by_person, date, min_people_free=1):
        usernames = list(busy_by_person.keys())
        start_of_day = datetime.combine(date, datetime.min.time()).replace(hour=9)
        all_slots = [start_of_day + timedelta(hours=i) for i in range(8)]
        free_slots = []

        ignored_slots = load_ignored_slots()  # ⬅️ load from file

        for slot_start in all_slots:
            m, d, h = slot_start.month, slot_start.day, slot_start.hour
            if (m, d, h) in ignored_slots:
                continue

            slot_end = slot_start + timedelta(hours=1)
            free_advisors = []

            for user in usernames:
                overlaps = any(slot_start < end and slot_end > start for start, end in busy_by_person[user])
                if not overlaps:
                    free_advisors.append(user)

            if len(free_advisors) >= min_people_free:
                free_slots.append((slot_start, free_advisors))

        return free_slots

    def find_next_5_days(usernames, min_people_free=1):
        results = []
        current_day = datetime.now().date() + timedelta(days=1)
        checked = 0

        while len(results) < 5 and checked < 30:
            if not is_weekend(current_day):
                start = datetime.combine(current_day, datetime.min.time()).replace(hour=9)
                end = datetime.combine(current_day, datetime.min.time()).replace(hour=17)
                busy = get_busy_times_by_person(usernames, start, end)

                free = get_flexible_free_blocks(busy, current_day, min_people_free)
                if free:
                    results.append((current_day, free))
            current_day += timedelta(days=1)
            checked += 1

        return results

    usernames = ["Megan Blair", "Daniel Lebrija", "Tong Kou Lor", "Kenny Adams", "Maisoon Alghethy"]
    availability = find_next_5_days(usernames)
    lines = []

    for date, slots in availability:
        date_str = date.strftime("%A, %B %d")
        for slot_time, advisors in slots:
            time_str = slot_time.strftime("%I:%M%p").lstrip("0")
        times = ", ".join(slot_time.strftime("%I:%M%p").lstrip("0") for slot_time, _ in slots)
        lines.append(f"• {date_str} - {times}")

    return "\n".join(lines)


# ──────────────────────────────────────────────────────
# Get the latest Appointment Request email from Outlook
# ──────────────────────────────────────────────────────
def get_latest_student_email():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    for msg in messages:
        if msg.Subject.startswith("Re: DPRC @ SF State - Initial Appointment Request") and msg.UnRead:
            student_email = msg.Body
            msg.UnRead = False  # mark as read if you want
            return student_email.strip()

    return None

def get_latest_email_by_subject(subject_keyword):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    for msg in messages:
        if subject_keyword in msg.Subject and msg.Class == 43:
            body = msg.Body
            return body.strip(), msg

    return None, None

def get_advisors_free_at(requested_dt):
    usernames = ["Megan Blair", "Daniel Lebrija", "Tong Kou Lor", "Kenny Adams", "Maisoon Alghethy"]

    # Define start/end of the 50-minute requested slot
    start_time = requested_dt
    end_time = requested_dt + timedelta(minutes=50)

    # Pull all busy times for all advisors during the whole day
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    root_folder = outlook.Folders.Item("dprc@sfsu.edu").Folders["Calendar"]
    all_busy = {}

    for name in usernames:
        busy_times = []
        try:
            calendar = root_folder.Folders[name]
            items = calendar.Items
            items.IncludeRecurrences = True
            items.Sort("[Start]")

            #if is_ooo_or_vacation(items):
            #    continue

            # Only pull items for the same day
            day_start = datetime.combine(requested_dt.date(), datetime.min.time()).replace(hour=8)
            day_end = datetime.combine(requested_dt.date(), datetime.min.time()).replace(hour=18)

            restriction = f"[Start] >= '{day_start.strftime('%m/%d/%Y %I:%M %p')}' AND [End] <= '{day_end.strftime('%m/%d/%Y %I:%M %p')}'"
            restricted_items = items.Restrict(restriction)

            for item in restricted_items:
                try:
                    # No skipping logic — if it has a time, we treat it as busy
                    if hasattr(item, "Start") and hasattr(item, "End"):
                        start = item.Start.replace(tzinfo=None)
                        end = item.End.replace(tzinfo=None)
                        if start.year < 2100:
                            busy_times.append((start, end))

                    if hasattr(item, "Start") and hasattr(item, "End"):
                        start = item.Start.replace(tzinfo=None)
                        end = item.End.replace(tzinfo=None)
                        if start.year < 2100:
                            busy_times.append((start, end))
                except Exception as e:
                    continue

        except Exception as e:
            print(f"⚠️ Could not access {name}'s calendar: {e}")
        all_busy[name] = busy_times

    # Now check which advisors are free during the requested time
    available_advisors = []

    for advisor, times in all_busy.items():
        print(f"\n🕒 Checking availability for {advisor} on {requested_dt.strftime('%A, %B %d at %I:%M%p')}")
        print("❌ Busy times:")
        for start, end in times:
            print(f"  - {start.strftime('%I:%M%p')} to {end.strftime('%I:%M%p')}")

        overlaps = any(start_time < end and end_time > start for start, end in times)

        if not overlaps:
            print(f"✅ {advisor} is free from {start_time.strftime('%I:%M%p')} to {end_time.strftime('%I:%M%p')}")
            available_advisors.append(advisor)
        else:
            print(f"⛔ {advisor} is busy during that time window.")

    return available_advisors

def get_day_suffix(day):
    return "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
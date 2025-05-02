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
from calendar_functions import get_advisors_free_at, get_busy_times_by_person
from datetime import datetime, timedelta, date

def get_day_suffix(day):
    return "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")

def open_scheduler_window(root, driver):
    scheduler = tk.Toplevel(root)
    scheduler.title("Schedule Appointment")
    scheduler.geometry("300x600")
    scheduler.attributes("-topmost", True)

    today = datetime.today()

    # Student Email / ID
    tk.Label(scheduler, text="Student Email or ID:").pack()
    sched_info_var = tk.StringVar()
    tk.Entry(scheduler, textvariable=sched_info_var).pack()

    # Month
    tk.Label(scheduler, text="Month:").pack()
    sched_month_var = tk.IntVar(value=today.month)
    tk.OptionMenu(scheduler, sched_month_var, *range(1, 13)).pack()

    # Year
    tk.Label(scheduler, text="Year:").pack()
    sched_year_var = tk.IntVar(value=today.year)
    tk.OptionMenu(scheduler, sched_year_var, *range(today.year, today.year + 2)).pack()

    # Day - Only weekdays
    tk.Label(scheduler, text="Day:").pack()
    sched_day_var = tk.IntVar()
    day_dropdown = tk.OptionMenu(scheduler, sched_day_var, "")
    day_dropdown.pack()

    def update_days(*args):
        day_dropdown["menu"].delete(0, "end")
        days_in_month = calendar.monthrange(sched_year_var.get(), sched_month_var.get())[1]
        for d in range(1, days_in_month + 1):
            dt = date(sched_year_var.get(), sched_month_var.get(), d)
            if dt.weekday() < 5:  # Mon-Fri only
                day_dropdown["menu"].add_command(
                    label=d,
                    command=tk._setit(sched_day_var, d)
                )

    sched_month_var.trace_add("write", update_days)
    sched_year_var.trace_add("write", update_days)
    update_days()

    # Time Selection (Hour + Minute side by side)
    tk.Label(scheduler, text="Time:").pack()

    time_frame = tk.Frame(scheduler)
    time_frame.pack()

    # Hour Dropdown
    hour_options = {
        "9 AM": "09", "10 AM": "10", "11 AM": "11",
        "12 PM": "12", "1 PM": "13", "2 PM": "14", "3 PM": "15", "4 PM": "16"
    }

    sched_hour_label = tk.StringVar(value="2 PM")
    tk.OptionMenu(
        time_frame,
        sched_hour_label,
        *hour_options.keys()
    ).pack(side="left", padx=5)

    sched_minute_var = tk.StringVar(value="00")
    tk.OptionMenu(
        time_frame,
        sched_minute_var,
        *[f"{m:02}" for m in range(0, 60, 5)]
    ).pack(side="left", padx=5)

    def update_advisors(*args):
        try:
            hour = int(hour_options[sched_hour_label.get()])
            minute = int(sched_minute_var.get())
            dt = datetime(sched_year_var.get(), sched_month_var.get(), sched_day_var.get(), hour, minute)
            available = get_advisors_free_at(dt)
            advisor_menu = advisor_dropdown["menu"]
            advisor_menu.delete(0, "end")
            for name in available:
                advisor_menu.add_command(label=name, command=tk._setit(sched_advisor_var, name))
            if available:
                sched_advisor_var.set(available[0])
            else:
                sched_advisor_var.set("")
        except Exception as e:
            print("⚠️ Advisor update failed:", e)

    # Duration selection
    tk.Label(scheduler, text="Duration:").pack()
    sched_duration_var = tk.StringVar(value="50")
    tk.OptionMenu(scheduler, sched_duration_var, "20", "30", "50").pack()

    # Advisor selection
    sched_advisor_var = tk.StringVar()
    advisor_dropdown = tk.OptionMenu(scheduler, sched_advisor_var, "")
    advisor_dropdown.pack()

    sched_hour_label.trace_add("write", update_advisors)
    sched_minute_var.trace_add("write", update_advisors)
    sched_day_var.trace_add("write", update_advisors)
    sched_month_var.trace_add("write", update_advisors)
    sched_year_var.trace_add("write", update_advisors)

    # Type selection
    tk.Label(scheduler, text="Type:").pack()
    sched_type_var = tk.StringVar(value="Zoom")  # Default value
    tk.OptionMenu(
        scheduler,
        sched_type_var,
        "Drop-in In-person",
        "Drop-in Phone/Zoom",
        "In-Person",
        "Phone",
        "Zoom"
    ).pack()

    # Appointment Category Checkboxes
    tk.Label(scheduler, text="Categories:").pack()

    first_meeting_var = tk.BooleanVar()
    initial_appt_var = tk.BooleanVar()
    follow_up_var = tk.BooleanVar()

    checkbox_frame = tk.Frame(scheduler)
    checkbox_frame.pack()

    tk.Checkbutton(checkbox_frame, text="First Meeting with Specialist", variable=first_meeting_var).grid(row=0, column=0, sticky="w")
    tk.Checkbutton(checkbox_frame, text="Follow-up", variable=follow_up_var).grid(row=0, column=1, sticky="w")
    tk.Checkbutton(checkbox_frame, text="Initial Appointment", variable=initial_appt_var).grid(row=1, column=0, sticky="w")

    # Auto-check based on duration
    def update_checkboxes(*args):
        duration = sched_duration_var.get()
        if duration == "50":
            first_meeting_var.set(True)
            initial_appt_var.set(True)
            follow_up_var.set(False)
        else:
            first_meeting_var.set(False)
            initial_appt_var.set(False)
            follow_up_var.set(True)

    sched_duration_var.trace_add("write", update_checkboxes)
    update_checkboxes()

    # Confirm button
    def confirm_appointment():
        info = sched_info_var.get()
        month = sched_month_var.get()
        day = sched_day_var.get()
        year = sched_year_var.get()
        hour = hour_options[sched_hour_label.get()]
        minute = sched_minute_var.get()
        time_str = f"{hour}:{minute}"
        duration = sched_duration_var.get()
        advisor = sched_advisor_var.get()
        appt_type = sched_type_var.get()
        first_meeting = first_meeting_var.get()
        initial_appt = initial_appt_var.get()
        follow_up = follow_up_var.get()


        # Validate inputs
        if not all([month, day, year, time_str, duration, advisor]):
            messagebox.showerror("Missing Info", "Please complete all fields.")
            return

        date_str = f"{month:02}/{day:02}/{year}"
        
        schedule_appt(driver, info, date_str, hour, minute, time_str, duration, advisor, appt_type, first_meeting, initial_appt, follow_up)
        
        scheduler.destroy()

    tk.Button(scheduler, text="Confirm Appointment", command=confirm_appointment).pack(pady=10)


def schedule_appt(driver, info, date_str, hour, minute, time_str, duration, advisor, appt_type, first_meeting, initial_appt, follow_up):
    driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$BtnScheduleAppointment").click()
    time.sleep(1)

    date_input = driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorAppointment$TxtBoxAddDate$TxtBoxInput")
    date_input.clear()
    date_input.send_keys(date_str)

    Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorAppointment$DDLAddTime$DDLCustomTime$DDLHour")).select_by_value(hour)
    Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorAppointment$DDLAddTime$DDLCustomTime$DDLMinute")).select_by_value(minute)
    Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorAppointment$DDLAddDuration$DDLInput")).select_by_value(duration)
    Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorAppointment$DDLAddCounselor$DDLInput")).select_by_visible_text(advisor)
    Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorAppointment$DDLAddAppointmentType$DDLInput")).select_by_visible_text(appt_type)

    if first_meeting:
        driver.find_element(By.ID, "ctl00_ctl00_MainContent_CphMainContent_ApplicationStudentList_FormValidatorAppointment_GunadiAppointmentListAdd_ChkBoxList_ctl00_AppointmentPurposeType_1").click()
    if initial_appt:
        driver.find_element(By.ID, "ctl00_ctl00_MainContent_CphMainContent_ApplicationStudentList_FormValidatorAppointment_GunadiAppointmentListAdd_ChkBoxList_ctl00_AppointmentPurposeType_4").click()
    if follow_up:
        driver.find_element(By.ID, "ctl00_ctl00_MainContent_CphMainContent_ApplicationStudentList_FormValidatorAppointment_GunadiAppointmentListAdd_ChkBoxList_ctl00_AppointmentPurposeType_2").click()

    driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorAppointment$BtnAppointment").click()

    button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$Button14"))
    )
    button.click()

    first = driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$TxtBoxFirstName$TxtBoxInput").get_attribute("value")
    email = driver.find_element(By.NAME,"ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$TxtBoxEmail$TxtBoxInput").get_attribute("value")

    EMAIL_TEMPLATE = """Hello {first_name},

Your appointment has been scheduled on {month_name} {day}{day_suffix} at {time_str}, with {advisor_short} Your appointment will be conducted virtually. You should have just received a calendar invite for your meeting including the zoom link to join.

Kenneth L.
Office Support Assistant
Disability Programs & Resource Center
Student Services Building, 110 | San Francisco State University
(p) (415) 338-2472 (voicemail, checked daily M–F)
"""

    # Parse the datetime from existing strings
    dt = datetime.strptime(f"{date_str} {time_str}", "%m/%d/%Y %H:%M")

    # Format values
    month_name = dt.strftime("%B")
    day = dt.day
    day_suffix = get_day_suffix(day)     
    formatted_time = dt.strftime("%I:%M%p").lstrip("0").lower()
    if " " in advisor:
        first_name, last_name = advisor.split(" ", 1)
        advisor_short = f"{first_name} {last_name[0]}."
    else:
        advisor_short = advisor


    # Format the final message
    final_email = EMAIL_TEMPLATE.format(
        first_name=first,
        month_name=month_name,
        day=day,
        day_suffix=day_suffix,
        time_str=formatted_time,
        advisor_short=advisor_short
    )

    final_email_html = f"""
    <html>
    <body>
    <p>Hello {first},</p>

    <p>Your appointment has been scheduled on <b>{month_name} {day}{day_suffix} at {formatted_time}</b>, with <b>{advisor_short}</b> 
    Your appointment will be conducted <b>virtually</b>. You should have just received a calendar invite for your meeting including the zoom link to join.</p>

    <p>
    Kenneth L.<br>
    Office Support Assistant<br>
    Disability Programs & Resource Center<br>
    Student Services Building, 110 | San Francisco State University<br>
    (p) (415) 338-2472 (voicemail, checked daily M–F)
    </p>
    </body>
    </html>
    """


    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = "DPRC @ SF State – Appointment Confirmation"
    mail.HTMLBody = final_email_html
    mail.Display()

    messagebox.showinfo("Review Email", "Press OK after sending the email.")

    Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$DDLApplicationStatus$DDLInput")).select_by_visible_text("2.2 - Initial Appointment Scheduled-Haven't met with specialist")
    Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$DDLAdvisor$DDLInput")).select_by_visible_text(advisor)
    driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$BtnUpdate2").click()
    time.sleep(1)

    Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$DDLQuickNoteTitle$DDLInput")).select_by_visible_text("Email Communication")
    Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$DDLQuickNoteType$DDLInput")).select_by_visible_text("Case Note")
    driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$TxtBoxNotesNote$TxtBoxInput").send_keys(final_email)
    driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$BtnNotes").click()

    staff_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ctl00_ctl00_MainNavigation_LblStaff"))
    )
    driver.execute_script("arguments[0].click();", staff_button)
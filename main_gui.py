import tkinter as tk
from tkinter import messagebox
import subprocess
import os
import time
import webbrowser
import win32com.client
import win32com.client
import calendar
import threading

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta, date
from outreach_functions import run_initial_outreach, run_second_outreach, run_missed_appointment
from calendar_functions import get_availability_text
from browser_setup import launch_chrome
from scheduler_window import open_scheduler_window
from calendar_openings import generate_calendar_times_txt

#pip install selenium webdriver-manager pywin32

driver = None  # Will be set after launching Chrome

def set_driver(d):
    global driver
    driver = d

def show_loading_popup(parent, message="Launching Chrome..."):
    popup = tk.Toplevel(parent)
    popup.title("Please wait")
    popup.geometry("220x80")
    popup.attributes("-topmost", True)
    tk.Label(popup, text=message).pack(expand=True, pady=20)
    popup.update()
    return popup

def threaded_chrome_launch(status_label, root):
    loading_popup = show_loading_popup(root, "Launching myDPRC...")

    def task():
        try:
            d = launch_chrome(status_label, root)
            set_driver(d)
        finally:
            loading_popup.destroy()

    threading.Thread(target=task).start()

def run_script(script_path, success_msg="✅ Script ran", fail_msg="❌ Failed"):
    try:
        subprocess.Popen(["python", script_path], shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"Script failed: {e}")

# GUI Setup
root = tk.Tk()
root.title("Kenneth's DPRC Automation")
root.geometry("300x400")

status_label = tk.Label(root, text="🟡 Chrome not running", fg="orange")
status_label.pack(pady=10)

calendar_status_var = tk.StringVar()
def get_last_updated_time():
    try:
        with open("calendar_times.txt", "r", encoding="utf-8") as f:
            lines = f.readlines()
            for line in reversed(lines):
                if line.startswith("Last updated:"):
                    return "📅 " + line.strip()
    except FileNotFoundError:
        return "📅 No update yet"

    return "📅 Update time not found"
calendar_status_var.set(get_last_updated_time())
calendar_status_label = tk.Label(root, textvariable=calendar_status_var, fg="green")

# Stay on Top Toggle
stay_on_top_var = tk.BooleanVar(value=True)
def toggle_stay_on_top():
    root.attributes("-topmost", stay_on_top_var.get())

stay_on_top_check = tk.Checkbutton(root, text="Stay on Top", variable=stay_on_top_var, command=toggle_stay_on_top)
stay_on_top_check.pack(pady=2)
root.attributes("-topmost", True)

# Buttons
tk.Button(root, text="Open myDPRC", command=lambda: threaded_chrome_launch(status_label, root)).pack(pady=2)

initial_frame = tk.Frame(root)
initial_frame.pack(pady=2)
initial_loops_var = tk.IntVar(value=1)
tk.Entry(initial_frame, textvariable=initial_loops_var, width=5).pack(side="right", padx=5)
tk.Button(initial_frame, text="Initial Outreach", command=lambda: run_initial_outreach(driver, initial_loops_var.get())).pack(side="right")

second_frame = tk.Frame(root)
second_frame.pack(pady=2)
second_loops_var = tk.IntVar(value=1)
tk.Entry(second_frame, textvariable=second_loops_var, width=5).pack(side="right", padx=5)
tk.Button(second_frame, text="Second Outreach", command=lambda: run_second_outreach(driver, second_loops_var.get())).pack(side="right")

missed_frame = tk.Frame(root)
missed_frame.pack(pady=2)
missed_loops_var = tk.IntVar(value=1)
tk.Entry(missed_frame, textvariable=missed_loops_var, width=5).pack(side="right", padx=5)
tk.Button(missed_frame, text="Missed Appointment", command=lambda: run_missed_appointment(driver, missed_loops_var.get())).pack(side="right")

tk.Button(root, text="Open Scheduler", command=lambda: open_scheduler_window(root, driver)).pack(pady=10)

def update_calendar_and_status():
    generate_calendar_times_txt()
    now = datetime.now().strftime("📅 Last updated: %A, %B %d at %I:%M %p")
    calendar_status_var.set(now)

root.after(500, lambda: threaded_chrome_launch(status_label, root))

# Bottom Frame for Calendar Update and Status
bottom_frame = tk.Frame(root)
bottom_frame.pack(side="bottom", fill="x", pady=10)

tk.Button(bottom_frame, text="Update Calendar Times", command=update_calendar_and_status).pack()
calendar_status_label = tk.Label(bottom_frame, textvariable=calendar_status_var, fg="green")
calendar_status_label.pack(pady=2)

root.mainloop()
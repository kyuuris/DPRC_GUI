import tkinter as tk
import subprocess
import os
import time
import webbrowser
import win32com.client
import calendar
import threading
import pytesseract
import pyautogui


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta, date
from outreach_functions import run_initial_outreach, run_second_outreach, run_missed_appointment, send_reminders
from calendar_functions import get_availability_text
from browser_setup import launch_chrome
from scheduler_window import open_scheduler_window
from calendar_openings import generate_calendar_times_txt
from PIL import Image
from tkinter import messagebox, ttk

#pip install selenium webdriver-manager pywin32 pyautogui pillow pytesseract


driver = None  # Will be set after launching Chrome

if not os.path.exists("ignored_slots.txt"):
    with open("ignored_slots.txt", "w", encoding="utf-8") as f:
        f.write("# Add one entry per line in the format: month day hour (24-hour time)\n")
        f.write("# Example:\n")
        f.write("# 5 12 10   # Ignores May 12 at 10AM\n")
        f.write("# 5 15 13   # Ignores May 15 at 1PM\n")

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

def save_raw_ignored_slot():
    try:
        raw = ignore_input.get().strip()
        month, day, hour = map(int, raw.split())

        if hour not in [9, 10, 11, 12, 1, 2, 3, 4]:
            raise ValueError("Hour must be one of: 9, 10, 11, 12, 1, 2, 3, 4")

        dt = datetime(datetime.now().year, month, day, hour)
        with open("ignored_slots.txt", "a", encoding="utf-8") as f:
            f.write(f"{month} {day} {hour}\n")
        ignore_input.delete(0, tk.END)

    except Exception as e:
        messagebox.showerror("Invalid Input", "Format must be: month day hour (e.g., 5 12 10)")

def update_calendar_and_status():
    generate_calendar_times_txt()
    now = datetime.now().strftime("📅 Last updated: %A, %B %d at %I:%M %p")
    calendar_status_var.set(now)

def run_script(script_path, success_msg="✅ Script ran", fail_msg="❌ Failed"):
    try:
        subprocess.Popen(["python", script_path], shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"Script failed: {e}")

# GUI Setup
root = tk.Tk()
root.title("Kenneth's DPRC Automation")
root.geometry("300x450")

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

buttons_frame = tk.Frame(root)
buttons_frame.pack(pady=5)

# Buttons
tk.Button(buttons_frame, text="Open myDPRC", command=lambda: threaded_chrome_launch(status_label, root)).pack(side="left", pady=5, padx=5)
tk.Button(buttons_frame, text="Send Daily Reminders", command=lambda: send_reminders(driver)).pack(side="left", pady=5, padx=5)

separator = ttk.Separator(root, orient="horizontal")
separator.pack(fill="x", pady=5)

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

separator = ttk.Separator(root, orient="horizontal")
separator.pack(fill="x", pady=5)

tk.Button(root, text="Open Scheduler", command=lambda: open_scheduler_window(root, driver)).pack(pady=10)

#root.after(500, lambda: threaded_chrome_launch(status_label, root))

separator = ttk.Separator(root, orient="horizontal")
separator.pack(fill="x", pady=5)

# Bottom Frame for Calendar Update and Status
bottom_frame = tk.Frame(root)
bottom_frame.pack(side="bottom", fill="x", pady=10)

tk.Button(bottom_frame, text="Update Calendar Times", command=update_calendar_and_status).pack()
calendar_status_label = tk.Label(bottom_frame, textvariable=calendar_status_var, fg="green")
calendar_status_label.pack(pady=2)

# ───── Add Ignore Slot UI ─────
ignore_simple_frame = tk.Frame(bottom_frame)
ignore_simple_frame.pack(pady=5)

tk.Label(ignore_simple_frame, text="Ignore Date (month day hour):").pack()

ignore_input = tk.Entry(ignore_simple_frame, width=15)
ignore_input.pack()
ignore_input.bind("<Return>", lambda event: save_raw_ignored_slot())

tk.Button(ignore_simple_frame, text="Add Ignored Slot", command=save_raw_ignored_slot).pack(pady=3)

root.mainloop()
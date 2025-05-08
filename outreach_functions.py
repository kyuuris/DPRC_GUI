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
from calendar_functions import get_availability_text
from calendar_openings import load_calendar_text_from_file
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains

def run_initial_outreach(driver, num_loops=1):
    try:
        for _ in range(num_loops):    
            link = driver.find_element(By.LINK_TEXT, "1 – Needs Review/Initial Reach-Out")
            link.click()
            time.sleep(1)

            view_buttons = driver.find_elements(By.LINK_TEXT, "View")
            if not view_buttons:
                messagebox.showerror("No Students", "No 'View' buttons found.")
                return

            view_buttons[0].click()
            time.sleep(1)

            first = driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$TxtBoxFirstName$TxtBoxInput").get_attribute("value")
            email = driver.find_element(By.NAME,"ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$TxtBoxEmail$TxtBoxInput").get_attribute("value")

            if not email:
                messagebox.showerror("Missing Email", "No email found.")
                return

            availability_text = load_calendar_text_from_file()

            availability_list = "".join(f"<li>{line.strip('• ').strip()}</li>" for line in availability_text.splitlines())

            case_note_email = f"""Hello {first},

Thank you for submitting your application to myDPRC. The next step would be to meet with a DPRC specialist for an initial appointment. Appointments are held virtually and last about 50 minutes. Please make sure to include at least 3 times you are available for as appointments fill up quickly!

Prescheduled 50 Minute Appointments:
{availability_list}

Kenneth L.
Office Support Assistant</b>
Disability Programs & Resource Center
Student Services Building, 110 | San Francisco State University
(415) 338-2472 (voicemail, checked daily M–F)
"""

            final_email_html = f"""
<html>
<body>
<p>Hello {first},</p>

<p>Thank you for submitting your application to myDPRC. The next step would be to meet with a DPRC specialist for an initial appointment. Appointments are held virtually and last about 50 minutes. <b>Please make sure to include <u>at least 3 times</u> you are available for as appointments fill up quickly!</b></p>

<p><b><u>Prescheduled 50 Minute Appointments:</u></b></p>
    <ul style="list-style-type: disc; margin-left: 20px;">
        {availability_list}
    </ul>

<p>
<b>Kenneth L.</b><br>
<b>Office Support Assistant</b><br>
<b>Disability Programs & Resource Center</b><br>
Student Services Building, 110 | San Francisco State University<br>
(p) (415) 338-2472 (voicemail, checked daily M–F)
</p>
</body>
</html>
"""
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = "DPRC @ SF State – Initial Appointment Request"
            mail.HTMLBody = final_email_html
            mail.send()

            Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$DDLApplicationStatus$DDLInput")).select_by_visible_text("2 - Initial Reach-Out Complete")
            driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$BtnUpdate2").click()
            time.sleep(1)

            Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$DDLQuickNoteTitle$DDLInput")).select_by_visible_text("Email Communication")
            Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$DDLQuickNoteType$DDLInput")).select_by_visible_text("Case Note")
            driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$TxtBoxNotesNote$TxtBoxInput").send_keys(case_note_email)
            driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$BtnNotes").click()

            staff_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ctl00_MainNavigation_LblStaff"))
            )
            driver.execute_script("arguments[0].click();", staff_button)

    except Exception as e:
        messagebox.showerror("Error", f"Failed during initial outreach: {e}")

def run_second_outreach(driver, num_loops=1):
    try:
        for _ in range(num_loops):
            link = driver.find_element(By.LINK_TEXT, "2 - Initial Reach-Out Complete")
            link.click()
            time.sleep(2)

            rows = driver.find_elements(By.XPATH, "//table//tr[td]")

            for row in rows:
                date_text = row.find_element(By.XPATH, ".//td[5]").text.strip()
                if date_text:
                    submitted_date = datetime.strptime(date_text, "%m/%d/%Y")
                    days_ago = (datetime.now() - submitted_date).days
                    if days_ago >= 15:
                        row.find_element(By.LINK_TEXT, "View").click()
                        time.sleep(2)
                        break

            first = driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$TxtBoxFirstName$TxtBoxInput").get_attribute("value")
            email = driver.find_element(By.NAME,"ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$TxtBoxEmail$TxtBoxInput").get_attribute("value")

            if not email:
                messagebox.showerror("Missing Email", "No email found.")
                return

            availability_text = load_calendar_text_from_file()

            availability_list = "".join(f"<li>{line.strip('• ').strip()}</li>" for line in availability_text.splitlines())

            case_note_email = f"""Hello {first},

We hope your semester has been well. We are reaching out to see if you were still interested in meeting with a DPRC specialist as the next step in our intake process. I have provided our open appointment times for the following weeks are listed below. Appointments are held virtually and last about 50 minutes. Please make sure to include <u>at least 3 times you are available for as appointments fill up quickly!

Prescheduled 50 Minute Appointments:
{availability_list}

Kenneth L.
Office Support Assistant
Disability Programs & Resource Center
Student Services Building, 110 | San Francisco State University
(p) (415) 338-2472 (voicemail, checked daily M–F)
"""
            final_email_html = f""" Hello {first},
<html>
<body>
<p>Hello {first},</p>

<p>We hope your semester has been well. We are reaching out to see if you were still interested in meeting with a DPRC specialist as the next step in our intake process. I have provided our open appointment times for the following weeks are listed below. Appointments are held virtually and last about 50 minutes. <b>Please make sure to include <u>at least 3 times</u> you are available for as appointments fill up quickly!</b></p>

<p><b><u>Prescheduled 50 Minute Appointments:</u></b></p>
    <ul style="list-style-type: disc; margin-left: 20px;">
        {availability_list}
    </ul>

<p>
<b>Kenneth L.</b><br>
<b>Office Support Assistant</b><br>
<b>Disability Programs & Resource Center</b><br>
Student Services Building, 110 | San Francisco State University<br>
(p) (415) 338-2472 (voicemail, checked daily M–F)
</p>
</body>
</html>
"""
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = "DPRC @ SF State – Initial Appointment Request Follow Up"
            mail.HTMLBody = final_email_html
            mail.send()

            Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$DDLApplicationStatus$DDLInput")).select_by_visible_text("2.1 - Second Reach-Out Complete")
            driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$BtnUpdate2").click()
            time.sleep(1)

            Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$DDLQuickNoteTitle$DDLInput")).select_by_visible_text("Email Communication")
            Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$DDLQuickNoteType$DDLInput")).select_by_visible_text("Case Note")
            driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$TxtBoxNotesNote$TxtBoxInput").send_keys(case_note_email)
            driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$BtnNotes").click()

            staff_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ctl00_MainNavigation_LblStaff"))
            )
            driver.execute_script("arguments[0].click();", staff_button)

    except Exception as e:
        messagebox.showerror("Error", f"Failed during second outreach: {e}")

def run_missed_appointment(driver, num_loops=1):
    try:
        for _ in range(num_loops):    
            link = driver.find_element(By.LINK_TEXT, "X - Student Missed Appointment | New Reach-Out Needed")
            link.click()
            time.sleep(2)

            view_buttons = driver.find_elements(By.LINK_TEXT, "View")
            if not view_buttons:
                messagebox.showerror("No Students", "No 'View' buttons found.")
                return

            view_buttons[0].click()
            time.sleep(1)

            first = driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$TxtBoxFirstName$TxtBoxInput").get_attribute("value")
            email = driver.find_element(By.NAME,"ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$TxtBoxEmail$TxtBoxInput").get_attribute("value")

            if not email:
                messagebox.showerror("Missing Email", "No email found.")
                return

            availability_text = load_calendar_text_from_file()

            availability_list = "".join(f"<li>{line.strip('• ').strip()}</li>" for line in availability_text.splitlines())

            case_note_email = f"""Hello {first},

We hope your semester has been well. We are reaching out to see if you were still interested in meeting with a DPRC specialist as the next step in our intake process. I have provided our open appointment times for the following weeks are listed below. Appointments are held virtually and last about 50 minutes. Please make sure to include <u>at least 3 times you are available for as appointments fill up quickly!

Prescheduled 50 Minute Appointments:
{availability_list}

Kenneth L.
Office Support Assistant
Disability Programs & Resource Center
Student Services Building, 110 | San Francisco State University
(p) (415) 338-2472 (voicemail, checked daily M–F)
"""

            final_email_html = f"""
<html>
<body>
<p>Hello {first},</p>

<p>We hope your semester has been well. We are reaching out to see if you were still interested in meeting with a DPRC specialist as the next step in our intake process. I have provided our open appointment times for the following weeks are listed below. Appointments are held virtually and last about 50 minutes. <b>Please make sure to include <u>at least 3 times</u> you are available for as appointments fill up quickly!</b></p>

<p><b><u>Prescheduled 50 Minute Appointments:</u></b></p>
    <ul style="list-style-type: disc; margin-left: 20px;">
        {availability_list}
    </ul>

<p>
<b>Kenneth L.</b><br>
<b>Office Support Assistant</b><br>
<b>Disability Programs & Resource Center</b><br>
Student Services Building, 110 | San Francisco State University<br>
(p) (415) 338-2472 (voicemail, checked daily M–F)
</p>
</body>
</html>
"""
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = "DPRC @ SF State – Missed Appointment Follow-Up"
            mail.HTMLBody = final_email_html
            mail.send()

            messagebox.showinfo("Review Email", "Press OK after sending the email.")

            Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$CtrlApplicationStudent$FormValidatorApps$DDLApplicationStatus$DDLInput")).select_by_visible_text("2.1 - Second Reach-Out Complete")
            driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$BtnUpdate2").click()
            time.sleep(1)

            Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$DDLQuickNoteTitle$DDLInput")).select_by_visible_text("Email Communication")
            Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$DDLQuickNoteType$DDLInput")).select_by_visible_text("Case Note")
            driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$TxtBoxNotesNote$TxtBoxInput").send_keys(case_note_email)
            driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$ApplicationStudentList$FormValidatorNotes$BtnNotes").click()

            staff_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ctl00_MainNavigation_LblStaff"))
            )
            driver.execute_script("arguments[0].click();", staff_button)

    except Exception as e:
        messagebox.showerror("Error", f"Failed during missed appointment outreach: {e}")

def send_reminders(driver):
    wait = WebDriverWait(driver, 10)

    manage_appointments = wait.until(EC.element_to_be_clickable((By.ID, "navSelected2")))
    ActionChains(driver).move_to_element(manage_appointments).click().perform()
    time.sleep(1)
    driver.find_element(By.LINK_TEXT, "Daily Appointments").click()
    time.sleep(1)
    manage_appointments = wait.until(EC.element_to_be_clickable((By.ID, "navSelected2")))
    ActionChains(driver).move_to_element(manage_appointments).click().perform()

    driver.find_element(By.ID, "ctl00_ctl00_MainContent_CphMainContent_AppointmentDuJour_DayNavigation_LinkNext").click()
    Select(driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$AppointmentDuJour$DDLType")).select_by_visible_text("For Selected")

    checkboxes = driver.find_elements(By.XPATH, "//table//input[@type='checkbox']")
    for checkbox in checkboxes:
        if not checkbox.is_selected():
            checkbox.click()

    driver.find_element(By.NAME, "ctl00$ctl00$MainContent$CphMainContent$AppointmentDuJour$BtnAction").click()

    staff_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "ctl00_ctl00_MainNavigation_LblStaff"))
    )
    driver.execute_script("arguments[0].click();", staff_button)
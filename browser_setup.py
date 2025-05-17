from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

def launch_chrome(status_label=None, root=None):
    options = Options()
    options.add_argument(r"--user-data-dir=C:/Users/924199104/AppData/Local/Google/Chrome/User Data")
    options.add_argument(r"--profile-directory=Profile 2")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-sync")
    options.add_argument("--start-maximized")
    options.add_experimental_option("detach", True)

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        driver.get("https://access.sfsu.edu/mydprc")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Current Students"))
        ).click()

        time.sleep(2)

        if password:
            try:
                password_field = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.NAME, "passwd"))
                )
                password_field.send_keys(password)
                driver.find_element(By.ID, "idSIButton9").click()
            except Exception as e:
                print(f"⚠️ Auto login failed: {e}")

        if status_label:
            status_label.config(text="✅ Chrome (Selenium) launched", fg="green")

        if root:
            poll_chrome_status(driver, status_label, root)

        return driver

    except Exception as e:
        if status_label:
            status_label.config(text="❌ Chrome launch failed", fg="red")
        print(f"Failed to launch Chrome: {e}")
        return None

def poll_chrome_status(driver, status_label, root):
    try:
        _ = driver.title  # just accessing to see if it's alive
    except:
        status_label.config(text="🟡 Chrome not running", fg="orange")
        return

    root.after(2000, lambda: poll_chrome_status(driver, status_label, root))

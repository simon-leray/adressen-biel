from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

APP_URL = "https://adressen-biel.streamlit.app/"

options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options,
)

try:
    print(f"Öffne {APP_URL} …")
    driver.get(APP_URL)

    try:
        button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//button[contains(text(), "Yes, get this app back up")]')
            )
        )
        button.click()
        print("App war im Schlaf – Wake-up-Button geklickt.")
    except Exception:
        print("App ist bereits aktiv, kein Wake-up nötig.")

finally:
    driver.quit()

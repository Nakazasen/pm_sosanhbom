import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

edge_options = EdgeOptions()
edge_options.add_argument("--headless=new")
edge_options.add_argument("--disable-gpu")
edge_options.add_argument("--no-sandbox")
edge_options.add_argument("--disable-dev-shm-usage")
edge_options.add_argument("--window-size=1920,1080")

print("Launching Edge driver...")
driver = webdriver.Edge(options=edge_options)

try:
    driver.get("http://tcmp3gwb:3000/")
    wait = WebDriverWait(driver, 15)
    
    # Login
    username_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='username']")))
    username_input.send_keys("vn_pe03")
    driver.find_element(By.CSS_SELECTOR, "input[name='password']").send_keys("vn_pe03")
    driver.find_element(By.CSS_SELECTOR, ".aw-login-signInButton button").click()
    
    # Search
    search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.aw-uiwidgets-searchBox")))
    search_input.send_keys("110C103NL0")
    search_input.send_keys(Keys.ENTER)
    
    open_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[command-id='Awp0ShowObjectCell']")))
    open_btn.click()
    time.sleep(5)
    
    # Click Reports tab
    print("Clicking Reports tab...")
    tabs = driver.find_elements(By.CSS_SELECTOR, ".sw-tab")
    for t in tabs:
        if "Reports" in t.text:
            t.click()
            time.sleep(3)
            break
            
    gen_btn = driver.find_elements(By.XPATH, "//button[contains(text(), 'Generate Report')]")
    if gen_btn:
        print("Clicking Generate Report...")
        gen_btn[0].click()
        time.sleep(3)
        
        # Look for 'PS - BOM Structure'
        ps_report = driver.find_elements(By.XPATH, "//*[contains(text(), 'PS - BOM Structure')]")
        if ps_report:
            print("Found 'PS - BOM Structure'! Clicking it...")
            ps_report[0].click()
            time.sleep(3)
            driver.save_screenshot("tc14_bom_report_dialog.png")
            
            # Inspect dialog or next steps
            dialogs = driver.find_elements(By.CSS_SELECTOR, ".sw-popup, .aw-layout-popup, .sw-dialog")
            print(f"Found {len(dialogs)} dialogs after selecting PS - BOM Structure:")
            for d in dialogs:
                print("Dialog text:", d.text.replace("\n", " | "))
                
            buttons = driver.find_elements(By.CSS_SELECTOR, ".sw-button, button")
            print("Action buttons:")
            for b in buttons:
                if b.text.strip():
                    print("  Button:", b.text.strip(), b.get_attribute("class"))
                    
            with open("tc14_bom_report_dialog_dom.html", "w", encoding="utf-8") as f:
                f.write(driver.find_element(By.TAG_NAME, "body").get_attribute("innerHTML"))
            print("Saved tc14_bom_report_dialog_dom.html")
            
finally:
    driver.quit()
    print("Driver quit.")

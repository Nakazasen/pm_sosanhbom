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
    
    # Direct search
    search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.aw-uiwidgets-searchBox")))
    search_input.send_keys("110C103NL0")
    search_input.send_keys(Keys.ENTER)
    
    open_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[command-id='Awp0ShowObjectCell']")))
    open_btn.click()
    time.sleep(5)
    
    # Click Awp0OpenGroup
    print("Clicking Awp0OpenGroup...")
    open_grp = driver.find_elements(By.CSS_SELECTOR, "button[command-id='Awp0OpenGroup']")
    if open_grp:
        open_grp[0].click()
        time.sleep(2)
        menus = driver.find_elements(By.CSS_SELECTOR, ".aw-popup-content, .sw-popup, .aw-command-popupContainer, .aw-widgets-cellListWidget")
        print(f"Found {len(menus)} popup menus after clicking Awp0OpenGroup.")
        for m in menus:
            print("Menu text:", m.text.replace("\n", " | "))
            
    # Click Reports tab
    print("\nClicking Reports tab...")
    tabs = driver.find_elements(By.CSS_SELECTOR, ".sw-tab")
    for t in tabs:
        if "Reports" in t.text:
            t.click()
            time.sleep(4)
            print("Reports tab loaded. Inspecting content...")
            r_content = driver.find_elements(By.CSS_SELECTOR, ".aw-widgets-cellListItem, .aw-report-panel, .sw-row, .aw-splm-tableRow")
            print(f"Found {len(r_content)} report elements:")
            for rc in r_content[:10]:
                print("  Report item:", rc.text.replace("\n", " | "))
            break
            
    # Now let's inspect Content tab again and check table header columns!
    print("\nNavigating back to Content tab to inspect table headers & columns...")
    for t in tabs:
        if "Content" in t.text:
            t.click()
            time.sleep(5)
            # Find table headers
            headers = driver.find_elements(By.CSS_SELECTOR, ".aw-splm-tableHeaderCell, .ui-grid-header-cell, th, .aw-jswidgets-tableHeaderCell")
            print(f"Found {len(headers)} table header columns:")
            for idx, h in enumerate(headers):
                print(f"  Col [{idx}]: {h.text.strip()}")
            break
            
    driver.save_screenshot("tc14_content_columns.png")
    
finally:
    driver.quit()
    print("Driver quit.")

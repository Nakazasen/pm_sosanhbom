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
    
    # Open item
    open_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[command-id='Awp0ShowObjectCell']")))
    open_btn.click()
    time.sleep(5)
    
    print("Item view loaded. Looking for Content tab...")
    # Click Content tab
    tabs = driver.find_elements(By.CSS_SELECTOR, ".sw-tab")
    content_tab = None
    for t in tabs:
        if "Content" in t.text:
            content_tab = t
            break
            
    if content_tab:
        print("Found Content tab! Clicking it...")
        content_tab.click()
        time.sleep(5)
        print("Content tab active!")
        driver.save_screenshot("tc14_content_tab.png")
        
        # Check elements in Content view: table rows, tree nodes, expand icons
        tree_rows = driver.find_elements(By.CSS_SELECTOR, ".aw-splm-tableRow, .ui-grid-row, .aw-jswidgets-tableRow, tr")
        print(f"Found {len(tree_rows)} rows in Content tab.")
        for idx, r in enumerate(tree_rows[:10]):
            print(f"  Row [{idx}]: {r.text.strip()[:80]}")
            
        # Check expand/collapse icons
        tree_expands = driver.find_elements(By.CSS_SELECTOR, ".aw-jswidgets-treeExpand, .aw-splm-tableTreeCommand, [icon-id*='tree'], [icon-id*='Expand'], [icon-id*='Chevron'], .aw-widgets-treeNodeToggle")
        print(f"Found {len(tree_expands)} tree expand toggles/icons.")
    else:
        print("Content tab not found!")
        
    # Check Export To Excel button
    print("Checking Export To Excel button...")
    export_btn = driver.find_elements(By.CSS_SELECTOR, "button[command-id='Awp0ExportToExcel']")
    if export_btn:
        print("Export button found! Clicking it...")
        export_btn[0].click()
        time.sleep(3)
        driver.save_screenshot("tc14_export_dialog.png")
        
        # Check if modal dialog appeared
        dialogs = driver.find_elements(By.CSS_SELECTOR, ".aw-layout-popup, .modal-dialog, .sw-popup, .aw-popup-screen")
        print(f"Found {len(dialogs)} dialogs/popups.")
        for d in dialogs:
            print("Dialog HTML snippet:", d.get_attribute("outerHTML")[:500])
            
        with open("tc14_export_dialog_dom.html", "w", encoding="utf-8") as f:
            f.write(driver.find_element(By.TAG_NAME, "body").get_attribute("innerHTML"))
        print("Saved tc14_export_dialog_dom.html")
    else:
        print("Export button not found in current view.")
        
finally:
    driver.quit()
    print("Driver quit.")

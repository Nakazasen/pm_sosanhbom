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
    
    # Content tab
    tabs = driver.find_elements(By.CSS_SELECTOR, ".sw-tab")
    for t in tabs:
        if "Content" in t.text:
            t.click()
            time.sleep(5)
            break
            
    # Check Select All command
    print("Checking Select All command...")
    sel_all = driver.find_elements(By.CSS_SELECTOR, "button[command-id='Awp0SelectAllObjectSet']")
    if sel_all:
        print("Found Select All button! Clicking it...")
        sel_all[0].click()
        time.sleep(2)
        
    # Open Export to Excel
    print("Opening Export to Excel dialog...")
    export_cmd = driver.find_element(By.CSS_SELECTOR, "button[command-id='Awp0ExportToExcel']")
    export_cmd.click()
    time.sleep(2)
    
    # Click Add Properties
    print("Clicking Add Properties button...")
    add_prop_btn = driver.find_elements(By.CSS_SELECTOR, "button[command-id='Awp0ExportSelectedColumnsAdd']")
    if add_prop_btn:
        add_prop_btn[0].click()
        time.sleep(3)
        driver.save_screenshot("tc14_add_properties_popup.png")
        
        # Check available properties
        props = driver.find_elements(By.CSS_SELECTOR, ".aw-widgets-cellListItem, .sw-property, .aw-layout-popup label")
        print(f"Found {len(props)} property elements in Add Properties popup:")
        for p in props[:20]:
            print("  Property:", p.text.replace("\n", " | "))
            
        with open("tc14_add_properties_dom.html", "w", encoding="utf-8") as f:
            f.write(driver.find_element(By.TAG_NAME, "body").get_attribute("innerHTML"))
        print("Saved tc14_add_properties_dom.html")
        
finally:
    driver.quit()
    print("Driver quit.")

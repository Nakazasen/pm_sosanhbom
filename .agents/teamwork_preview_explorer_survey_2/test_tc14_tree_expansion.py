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
            
    print("Inspecting Content table structure...")
    # Find table container
    table = driver.find_elements(By.CSS_SELECTOR, ".aw-splm-table, .aw-jswidgets-grid, [data-locator='aw-splm-table']")
    print(f"Found {len(table)} table containers.")
    
    # Inspect rows and find rows that have expand toggles
    toggles = driver.find_elements(By.CSS_SELECTOR, ".aw-jswidgets-treeExpand, .aw-splm-tableTreeCommand, [icon-id*='Chevron'], [icon-id*='tree'], .aw-theme-iconOutline")
    print(f"Found {len(toggles)} toggle/icon elements.")
    
    # Let's inspect rows with class aw-splm-tableRow
    rows = driver.find_elements(By.CSS_SELECTOR, ".aw-splm-tableRow")
    print(f"Total .aw-splm-tableRow: {len(rows)}")
    
    expandable_rows = []
    for idx, r in enumerate(rows):
        r_text = r.text.replace("\n", " | ")
        t_in_row = r.find_elements(By.CSS_SELECTOR, ".aw-jswidgets-treeExpand, [icon-id*='miscDownArrow'], [icon-id*='miscRightArrow'], svg, .aw-splm-tableTreeCommand")
        if t_in_row:
            expandable_rows.append((idx, r, r_text, t_in_row))
            print(f"Expandable Row [{idx}]: text={r_text[:60]}, toggles={len(t_in_row)}")
            
    if expandable_rows:
        target_idx, target_row, target_text, target_toggles = expandable_rows[0]
        print(f"\nAttempting to expand Row {target_idx}: {target_text}")
        initial_row_count = len(driver.find_elements(By.CSS_SELECTOR, ".aw-splm-tableRow"))
        print(f"Initial row count: {initial_row_count}")
        
        # Click the toggle
        print("Clicking tree toggle on row...")
        driver.execute_script("arguments[0].scrollIntoView(true);", target_toggles[0])
        time.sleep(1)
        target_toggles[0].click()
        
        # Wait for DOM expansion
        print("Waiting 5 seconds for expansion...")
        time.sleep(5)
        
        new_row_count = len(driver.find_elements(By.CSS_SELECTOR, ".aw-splm-tableRow"))
        print(f"New row count after click: {new_row_count} (diff: {new_row_count - initial_row_count})")
        
        # Save screenshot
        driver.save_screenshot("tc14_tree_expanded.png")
        
        # Dump table HTML
        table_html = target_row.find_element(By.XPATH, "../..").get_attribute("outerHTML")
        with open("tc14_table_tree_dom.html", "w", encoding="utf-8") as f:
            f.write(table_html[:20000])
        print("Saved snippet to tc14_table_tree_dom.html")
        
finally:
    driver.quit()
    print("Driver quit.")

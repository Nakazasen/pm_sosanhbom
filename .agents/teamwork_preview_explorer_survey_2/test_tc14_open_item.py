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
    
    # Wait for search box
    search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.aw-uiwidgets-searchBox")))
    search_input.send_keys("110C103NL0")
    search_input.send_keys(Keys.ENTER)
    
    # Wait for search results
    print("Waiting for search results...")
    open_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[command-id='Awp0ShowObjectCell']")))
    print("Found Open button! Clicking it...")
    open_btn.click()
    
    # Wait for item object view to load
    print("Waiting for item view to load...")
    time.sleep(8)
    
    print("Current URL after Open:", driver.current_url)
    print("Page Title after Open:", driver.title)
    
    driver.save_screenshot("tc14_item_view.png")
    
    # Check tabs on the item view
    tabs = driver.find_elements(By.CSS_SELECTOR, ".sw-tab, .aw-layout-tabBar li, .sw-tabContainer li, [role='tab']")
    print(f"Found {len(tabs)} tabs:")
    for idx, t in enumerate(tabs):
        print(f"  Tab [{idx}]: text='{t.text.strip()}', class='{t.get_attribute('class')}', role='{t.get_attribute('role')}'")
        
    # Check commands in command bar
    cmds = driver.find_elements(By.CSS_SELECTOR, "button[command-id], .aw-commands-commandIconButton, [data-command-id]")
    print(f"Found {len(cmds)} commands:")
    for idx, c in enumerate(cmds[:30]):
        cmd_id = c.get_attribute("command-id") or c.get_attribute("data-command-id") or c.get_attribute("button-id")
        aria = c.get_attribute("aria-label") or c.get_attribute("title") or c.text.strip()
        print(f"  Cmd [{idx}]: id='{cmd_id}', label='{aria}', class='{c.get_attribute('class')[:40]}'")
        
    with open("tc14_item_view_dom.html", "w", encoding="utf-8") as f:
        f.write(driver.find_element(By.TAG_NAME, "body").get_attribute("innerHTML"))
    print("Saved tc14_item_view_dom.html")
    
finally:
    driver.quit()
    print("Driver quit.")

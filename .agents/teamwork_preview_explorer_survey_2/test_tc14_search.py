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
    
    # Wait for search box to appear
    print("Waiting for search box...")
    search_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.aw-uiwidgets-searchBox")))
    print("Found search box!")
    
    # Let's inspect siblings/parent of search box
    parent = search_input.find_element(By.XPATH, "..")
    print("Parent HTML:", parent.get_attribute("outerHTML"))
    
    # Test typing query and pressing ENTER
    query = "110C103NL0"
    print(f"Typing search query: {query} and sending ENTER...")
    search_input.send_keys(query)
    time.sleep(1)
    search_input.send_keys(Keys.ENTER)
    
    # Wait for results or URL transition
    print("Waiting for search results transition...")
    time.sleep(8)
    
    print("Current URL after search:", driver.current_url)
    print("Page Title after search:", driver.title)
    
    driver.save_screenshot("tc14_search_results.png")
    
    # Inspect search result elements
    # Common Active Workspace result classes: aw-widgets-cellListWidget, aw-jswidgets-table, etc.
    results = driver.find_elements(By.CSS_SELECTOR, ".aw-widgets-cellListCell, .aw-widgets-cellListItem, .ui-grid-row, .aw-splm-tableRow, tr, .aw-item-properties")
    print(f"Found {len(results)} potential result rows/cells.")
    
    # Check text of first few elements
    for idx, r in enumerate(results[:10]):
        print(f"  Result [{idx}] ({r.get_attribute('class')}): {r.text.strip()[:100]}")
        
    with open("tc14_search_dom.html", "w", encoding="utf-8") as f:
        f.write(driver.find_element(By.TAG_NAME, "body").get_attribute("innerHTML"))
    print("Saved tc14_search_dom.html")
    
finally:
    driver.quit()
    print("Driver quit.")

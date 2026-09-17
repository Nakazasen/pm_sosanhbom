import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
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
    url = "http://tcmp3gwb:3000/"
    print(f"Navigating to {url}...")
    driver.get(url)
    
    # Wait up to 20 seconds for DOM rendering
    print("Waiting for page load and React SPA hydration...")
    time.sleep(5) # initial wait for bundles
    
    print("Current URL:", driver.current_url)
    print("Page Title:", driver.title)
    
    # Check cookies
    cookies = driver.get_cookies()
    print(f"Cookies ({len(cookies)}):")
    for c in cookies:
        print(f"  {c['name']} = {c['value'][:20]}... (domain={c.get('domain')}, path={c.get('path')})")
        
    # Check localStorage & sessionStorage
    local_storage = driver.execute_script("return window.localStorage;")
    session_storage = driver.execute_script("return window.sessionStorage;")
    print("localStorage keys:", list(local_storage.keys()) if local_storage else "empty")
    print("sessionStorage keys:", list(session_storage.keys()) if session_storage else "empty")
    
    # Search for input elements
    inputs = driver.find_elements(By.TAG_NAME, "input")
    print(f"Found {len(inputs)} input elements:")
    for idx, inp in enumerate(inputs):
        try:
            print(f"  [{idx}] tag=input, type={inp.get_attribute('type')}, id={inp.get_attribute('id')}, name={inp.get_attribute('name')}, class={inp.get_attribute('class')}, placeholder={inp.get_attribute('placeholder')}, aria-label={inp.get_attribute('aria-label')}")
        except Exception as e:
            print(f"  [{idx}] error: {e}")
            
    # Search for button elements
    buttons = driver.find_elements(By.TAG_NAME, "button")
    print(f"Found {len(buttons)} button elements:")
    for idx, btn in enumerate(buttons):
        try:
            print(f"  [{idx}] text='{btn.text}', id={btn.get_attribute('id')}, name={btn.get_attribute('name')}, class={btn.get_attribute('class')}, aria-label={btn.get_attribute('aria-label')}")
        except Exception as e:
            print(f"  [{idx}] error: {e}")
            
    # Dump body html snippet
    body = driver.find_element(By.TAG_NAME, "body")
    body_html = body.get_attribute("innerHTML")
    with open("tc14_login_page_dom.html", "w", encoding="utf-8") as f:
        f.write(body_html)
    print("Wrote tc14_login_page_dom.html (len:", len(body_html), ")")
    
finally:
    driver.quit()
    print("Driver quit.")

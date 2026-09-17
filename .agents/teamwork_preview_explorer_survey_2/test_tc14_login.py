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
    
    wait = WebDriverWait(driver, 15)
    
    print("Waiting for username input...")
    username_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='username']")))
    username_input.clear()
    username_input.send_keys("vn_pe03")
    print("Entered username: vn_pe03")
    
    password_input = driver.find_element(By.CSS_SELECTOR, "input[name='password']")
    password_input.clear()
    password_input.send_keys("vn_pe03")
    print("Entered password: [HIDDEN]")
    
    time.sleep(1)
    
    # Wait for login button to be clickable
    login_btn = driver.find_element(By.CSS_SELECTOR, ".aw-login-signInButton button")
    print("Login button classes:", login_btn.get_attribute("class"))
    
    print("Clicking login button...")
    login_btn.click()
    
    # Wait for URL change or main container to transition away from loginViewPanel
    print("Waiting for post-login transition...")
    time.sleep(10)
    
    print("Current URL after login:", driver.current_url)
    print("Page Title after login:", driver.title)
    
    cookies = driver.get_cookies()
    print(f"Cookies after login ({len(cookies)}):")
    for c in cookies:
        print(f"  {c['name']} = {c['value'][:25]}... (domain={c.get('domain')})")
        
    local_storage = driver.execute_script("return Object.keys(window.localStorage);")
    print("localStorage keys after login:", local_storage)
    
    # Check if there are error toasts or popups
    errors = driver.find_elements(By.CSS_SELECTOR, ".aw-layout-error, .noty_text, .sw-error, .aw-widgets-propertyError")
    if errors:
        print(f"Found {len(errors)} error elements:")
        for e in errors:
            print("  Error text:", e.text)
            
    # Search for search bar, navigation bar, command bars
    search_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='search'], input[placeholder*='Search'], input[aria-label*='Search'], input[aria-label*='search']")
    print(f"Found {len(search_inputs)} search input elements:")
    for s in search_inputs:
        print(f"  placeholder='{s.get_attribute('placeholder')}', aria-label='{s.get_attribute('aria-label')}', class='{s.get_attribute('class')}'")
        
    # Check commands/buttons
    buttons = driver.find_elements(By.CSS_SELECTOR, "button, .aw-commands-commandIconButton, .sw-command")
    print(f"Found {len(buttons)} button/command elements. Top 15:")
    for b in buttons[:15]:
        txt = b.text.strip() or b.get_attribute("aria-label") or b.get_attribute("title")
        print(f"  tag={b.tag_name}, class={b.get_attribute('class')}, label={txt}")
        
    # Save post-login snapshot
    body_html = driver.find_element(By.TAG_NAME, "body").get_attribute("innerHTML")
    with open("tc14_post_login_dom.html", "w", encoding="utf-8") as f:
        f.write(body_html)
    print(f"Saved tc14_post_login_dom.html (len: {len(body_html)})")
    
    driver.save_screenshot("tc14_post_login.png")
    print("Saved tc14_post_login.png screenshot.")

except Exception as ex:
    print("Exception during login test:", ex)
    import traceback
    traceback.print_exc()
finally:
    driver.quit()
    print("Driver quit.")

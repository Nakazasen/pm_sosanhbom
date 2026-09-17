import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

download_dir = os.path.abspath("downloads_test")
os.makedirs(download_dir, exist_ok=True)

edge_options = EdgeOptions()
edge_options.add_argument("--headless=new")
edge_options.add_argument("--disable-gpu")
edge_options.add_argument("--no-sandbox")
edge_options.add_argument("--disable-dev-shm-usage")
edge_options.add_argument("--window-size=1920,1080")
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
edge_options.add_experimental_option("prefs", prefs)

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
    print("Checking Reports tab Generate Report...")
    tabs = driver.find_elements(By.CSS_SELECTOR, ".sw-tab")
    for t in tabs:
        if "Reports" in t.text:
            t.click()
            time.sleep(3)
            break
            
    gen_btn = driver.find_elements(By.XPATH, "//button[contains(text(), 'Generate Report')]")
    if gen_btn:
        print("Found 'Generate Report' button! Clicking it...")
        gen_btn[0].click()
        time.sleep(3)
        driver.save_screenshot("tc14_generate_report_panel.png")
        report_options = driver.find_elements(By.CSS_SELECTOR, ".aw-report-tile, .aw-widgets-cellListItem, .sw-popup, .aw-layout-popup")
        print(f"Found {len(report_options)} elements in report panel.")
        for ro in report_options[:10]:
            print("  Report option text:", ro.text.replace("\n", " | "))
            
    # Go to Content tab
    print("\nGoing to Content tab...")
    tabs = driver.find_elements(By.CSS_SELECTOR, ".sw-tab")
    for t in tabs:
        if "Content" in t.text:
            t.click()
            time.sleep(5)
            break
            
    # Click Export to Excel
    print("Clicking Export To Excel button...")
    export_cmd = driver.find_elements(By.CSS_SELECTOR, "button[command-id='Awp0ExportToExcel']")
    if export_cmd:
        export_cmd[0].click()
        time.sleep(2)
        
        # Click the Export button inside the side dialog
        export_action_btn = driver.find_elements(By.XPATH, "//button[normalize-space()='Export']")
        if export_action_btn:
            print("Found Export action button in dialog! Clicking it...")
            export_action_btn[0].click()
            
            # Wait and observe toasts / notifications / downloads
            print("Waiting 15 seconds for export processing...")
            for s in range(15):
                time.sleep(1)
                files = os.listdir(download_dir)
                if files:
                    print(f"Downloaded files so far at second {s}: {files}")
                    break
                    
            toasts = driver.find_elements(By.CSS_SELECTOR, ".noty_message, .aw-layout-popup, .sw-toast, .noty_text")
            print(f"Found {len(toasts)} toasts/messages:")
            for to in toasts:
                print("  Toast:", to.text.replace("\n", " | "))
                
            driver.save_screenshot("tc14_after_export.png")
            print("Current files in download dir:", os.listdir(download_dir))
    else:
        print("Export button not found!")
        
finally:
    driver.quit()
    print("Driver quit.")

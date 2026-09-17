import os
import sys
import time
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By

download_dir = os.path.abspath("downloads_verified")
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

driver = webdriver.Edge(options=edge_options)

# Enable CDP download behavior
driver.execute_cdp_cmd("Page.setDownloadBehavior", {
    "behavior": "allow",
    "downloadPath": download_dir
})

print("CDP setDownloadBehavior enabled for:", download_dir)
driver.quit()
print("Verification complete.")

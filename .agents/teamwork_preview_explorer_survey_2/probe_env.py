import os
import sys
import subprocess
import requests

print("=== Checking Browsers ===")
edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
if not os.path.exists(edge_path):
    edge_path = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
if os.path.exists(edge_path):
    print("Found Edge at:", edge_path)
else:
    print("Edge not found in default paths")

chrome_paths = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe")
]
found_chrome = None
for cp in chrome_paths:
    if os.path.exists(cp):
        found_chrome = cp
        print("Found Chrome at:", cp)
        break
if not found_chrome:
    print("Chrome not found in default paths")

print("\n=== Testing Selenium Manager ===")
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService

try:
    edge_options = EdgeOptions()
    edge_options.add_argument("--headless=new")
    edge_options.add_argument("--disable-gpu")
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Edge(options=edge_options)
    print("Edge WebDriver initialized successfully!")
    print("Driver capabilities browserVersion:", driver.capabilities.get('browserVersion'))
    driver.quit()
except Exception as e:
    print("Edge WebDriver init failed:", e)

try:
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    chrome_options = ChromeOptions()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(options=chrome_options)
    print("Chrome WebDriver initialized successfully!")
    print("Chrome capabilities browserVersion:", driver.capabilities.get('browserVersion'))
    driver.quit()
except Exception as e:
    print("Chrome WebDriver init failed:", e)

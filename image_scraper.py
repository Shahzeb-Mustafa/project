import os
import time
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image
from io import BytesIO

def search_and_download_images(query, num_images=5):
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')  # Run in headless mode
    options.add_argument('--start-maximized')
    options.add_argument('--disable-gpu')  # Disable GPU acceleration (useful for headless mode)
    options.add_argument('--no-sandbox')  # Bypass OS security model
    options.add_argument('--disable-dev-shm-usage')  # Overcome limited resources in some environments
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64)")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    search_url = f"https://www.bing.com/images/search?q={query}&qft=+filterui:imagesize-large"
    driver.get(search_url)

    for _ in range(5):
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.END)
        time.sleep(2)

    image_elements = driver.find_elements(By.CSS_SELECTOR, "img.mimg")
    image_urls = set()

    for img in image_elements:
        src = img.get_attribute('src') or img.get_attribute('data-src')
        if src and src.startswith('http'):
            image_urls.add(src)
            if len(image_urls) >= num_images:
                break

    driver.quit()

    save_path = os.path.abspath(query.replace(' ', '_'))
    os.makedirs(save_path, exist_ok=True)

    downloaded_images = []

    for i, url in enumerate(image_urls):
        try:
            response = requests.get(url, stream=True, timeout=10)
            response.raise_for_status()
            img = Image.open(BytesIO(response.content))
            img = img.convert('RGB')
            img_path = os.path.join(save_path, f"image_{i+1}.jpg")
            img.save(img_path, quality=95)
            downloaded_images.append(img_path)
        except Exception as e:
            print(f"Failed to save image from {url}: {e}")

    return downloaded_images

from typing import Dict, List, Optional
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium_stealth import stealth


def prepare_browser() -> webdriver.Chrome:
    """Return a new Chrome webdriver with stealth settings."""
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("start-maximized")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(options=chrome_options)
    stealth(driver,
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.53 Safari/537.36',
            languages=["en-US", "en"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=False,
            run_on_insecure_origins=False,
            )
    return driver


def scrape(username: str) -> Optional[Dict[str, str]]:
    """Scrape Instagram user data for a given username and return a dictionary of the scraped data."""
    url = f'https://instagram.com/{username}/?__a=1&__d=dis'
    chrome = prepare_browser()
    chrome.get(url)

    # Check if redirected to login page
    if "login" in chrome.current_url:
        print(f"Failed to scrape {username}: Redirected to login page.")
        chrome.quit()
        return None

    # Parse user data from JSON response
    resp_body = chrome.find_element(By.TAG_NAME, "body").text
    data_json = json.loads(resp_body)
    user_data = data_json['graphql']['user']

    # Return dictionary of user data
    chrome.quit()
    return {
        'name': user_data['full_name'],
        'category': user_data['category_name'],
        'followers': user_data['edge_followed_by']['count'],
        'posts': [edge['node']['edge_media_to_caption']['edges'][0]['node']['text']
                  for edge in user_data['edge_owner_to_timeline_media']['edges']
                  if edge['node']['edge_media_to_caption']['edges']]
    }

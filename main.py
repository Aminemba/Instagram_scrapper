import json
import openpyxl
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium_stealth import stealth


def prepare_browser():
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


def get_user_data(username):
    url = f"https://www.instagram.com/{username}/?__a=1"
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()
    user_data = data["graphql"]["user"]
    return user_data


def parse_user_data(user_data):
    captions = []
    for node in user_data['edge_owner_to_timeline_media']['edges']:
        caption = node['node']['edge_media_to_caption']['edges'][0]['node']['text']
        if caption:
            captions.append(caption)
    return {
        'name': user_data['full_name'],
        'category': user_data['category_name'],
        'followers': user_data['edge_followed_by']['count'],
        'posts': captions,
    }


def main():
    username = input("Enter an Instagram username to scrape: ")
    try:
        user_data = get_user_data(username)
        parsed_data = parse_user_data(user_data)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Instagram Data"
        ws.append(["Name", "Category", "Followers", "Posts"])
        ws.append([
            parsed_data['name'],
            parsed_data['category'],
            parsed_data['followers'],
            "\n".join(parsed_data['posts']),
        ])
        wb.save(f"{username}.xlsx")
    except requests.exceptions.HTTPError as e:
        print(f"Error getting user data: {e}")
    except (KeyError, IndexError) as e:
        print(f"Error parsing user data: {e}")


if __name__ == '__main__':
    main()

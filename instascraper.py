from typing import Dict, List, Optional
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium_stealth import stealth
import openpyxl

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
    user_data = json.loads(resp_body)['graphql']['user']

    # Return dictionary of user data
    data_dict = {
        'name': user_data['full_name'],
        'category': user_data['category_name'],
        'followers': user_data['edge_followed_by']['count'],
        'following': user_data['edge_follow']['count'],
        'posts': [edge['node']['edge_media_to_caption']['edges'][0]['node']['text']
                  for edge in user_data['edge_owner_to_timeline_media']['edges']
                  if edge['node']['edge_media_to_caption']['edges']],
        'comment_count': [edge['node']['edge_media_to_comment']['count']
                          for edge in user_data['edge_owner_to_timeline_media']['edges']
                          if edge['node']['edge_media_to_comment']],
        'likes_count': [edge['node']['edge_liked_by']['count']
                        for edge in user_data['edge_owner_to_timeline_media']['edges']
                        if edge['node']['edge_liked_by']]
    }



    chrome.quit()
    return data_dict


def save_to_excel(username: str, data: dict) -> None:
    """Save scraped data to an Excel file with the given username."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Instagram Data"
    ws.append(["Name", "Category", "Followers", "Following", "Posts", "Comment Count", "Likes Count"])
    posts = "\n".join(data['posts'])  # Combine all posts into one string separated by newline
    post_list = posts.split("\n")  # Split the string into a list of separate posts
    comments = ", ".join(map(str, data['comment_count']))  # Combine comment counts into a comma-separated string
    likes = ", ".join(map(str, data['likes_count']))  # Combine likes counts into a comma-separated string
    ws.append([
        data['name'],
        data['category'],
        data['followers'],
        data['following'],
        post_list[0],  # Append the first post separately
        data['comment_count'][0] if len(data['comment_count']) > 0 else "",
        data['likes_count'][0] if len(data['likes_count']) > 0 else 0
    ])
    for i in range(1, len(post_list)):  # Append each remaining post, comment count, and likes count on a separate line
        comment_count = data['comment_count'][i] if i < len(data['comment_count']) else ""
        likes_count = data['likes_count'][i] if i < len(data['likes_count']) else 0
        ws.append(["", "", "", "", post_list[i], comment_count, likes_count])
    wb.save(f"{username}.xlsx")
    print(f"Data for {username} saved to {username}.xlsx")


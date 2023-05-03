from typing import List
import openpyxl
from instascraper import scrape


def save_to_excel(username: str, data: dict) -> None:
    """Save scraped data to an Excel file with the given username."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Instagram Data"
    ws.append(["Name", "Category", "Followers", "Posts"])
    ws.append([
        data['name'],
        data['category'],
        data['followers'],
        "\n".join(data['posts']),
    ])
    wb.save(f"{username}.xlsx")
    print(f"Data for {username} saved to {username}.xlsx")


def main() -> None:
    """Main function for running the Instagram scraper."""
    usernames: List[str] = [input("Enter a username to scrape: ")]
    for username in usernames:
        scraped_data = scrape(username)
        if scraped_data:
            save_to_excel(username, scraped_data)



if __name__ == '__main__':
    main()

from typing import List
from instascraper import scrape , save_to_excel




def main() -> None:
    """Main function for running the Instagram scraper."""
    usernames: List[str] = [input("Enter a username to scrape: ")]
    for username in usernames:
        scraped_data = scrape(username)
        if scraped_data:
            save_to_excel(username, scraped_data)



if __name__ == '__main__':
    main()

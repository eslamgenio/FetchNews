import requests
import openpyxl
from datetime import datetime
import time

# Set up API key (replace with your News API key)
NEWS_API_KEY = 'NEWS_API_KEY'

# Function to fetch top generative AI news article titles in English
def fetch_latest_ai_news_titles():
    url = f'https://newsapi.org/v2/everything?q=generative AI&language=en&sortBy=publishedAt&pageSize=30&apiKey={NEWS_API_KEY}'
    response = requests.get(url)
    articles = response.json().get("articles", [])
    
    # Filter to include only articles with "AI" in the title
    return [{"title": article["title"], "link": article["url"]}
            for article in articles if "AI" in article["title"]]

# Function to check if a title already exists in the Excel sheet
def title_exists(sheet, title):
    for row in sheet.iter_rows(min_row=2, max_col=1, values_only=True):  # Assume titles are in column A starting from row 2
        if row[0] == title:
            return True
    return False

# Retry mechanism for updating Excel in case of PermissionError
def update_excel_with_titles(titles):
    workbook_path = 'AI_Latest_Updates.xlsx'
    max_retries = 5  # Maximum retry attempts
    retry_delay = 2  # Seconds to wait between retries

    for attempt in range(max_retries):
        try:
            # Load the workbook and select the active sheet
            workbook = openpyxl.load_workbook(workbook_path)
            sheet = workbook.active  # Assuming the first sheet is "AI_Latest_Updates"

            # Find the next available row in the sheet
            next_row = sheet.max_row + 1

            # Update the sheet with new titles, links, and current date, only if the title is not a duplicate
            current_date = datetime.now().strftime("%Y-%m-%d")
            for title_data in titles:
                if not title_exists(sheet, title_data['title']):  # Check for duplicates
                    sheet[f"A{next_row}"].value = title_data['title']
                    sheet[f"B{next_row}"].value = current_date
                    sheet[f"C{next_row}"].value = title_data['link']
                    next_row += 1
                else:
                    print(f"Skipping duplicate title: {title_data['title']}")

            # Save the workbook
            workbook.save(workbook_path)
            print("Excel sheet updated successfully with article titles.")
            break  # Exit loop if successful

        except PermissionError:
            print(f"Permission denied for {workbook_path}. Retrying in {retry_delay} seconds...")
            time.sleep(retry_delay)
        
        except Exception as e:
            print(f"An error occurred: {e}")
            break
    else:
        print("Failed to access the Excel file after multiple attempts. Please close the file if it’s open and try again.")

def main():
    # Step 1: Fetch the latest generative AI news titles in English
    titles = fetch_latest_ai_news_titles()
    
    # Step 2: Update the Excel sheet with titles, links, and date
    update_excel_with_titles(titles)

# Run the main function
if __name__ == "__main__":
    main()

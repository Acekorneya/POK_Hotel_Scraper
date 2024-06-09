# POK Hotel Data Scraper 2.6

This project is a hotel data scraper that extracts hotel information from travel websites and stores it in an Excel file. It uses multithreading for efficient data fetching and provides a graphical user interface (GUI) for user interaction.

## Table of Contents
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [GUI Interface](#gui-interface)
- [Dependencies](#dependencies)
- [Contributing](#contributing)
- [License](#license)

## Features
- Scrape hotel information such as name, room count, address, contact details, year built, and year renovated.
- Use multithreading for faster data fetching.
- Handle retries for HTTP requests.
- GUI for input and displaying progress.
- Save extracted data to Excel with filtering and formatting.
- Stop scraping process gracefully.

## Installation
### Prerequisites
- Python 3.x
- pip (Python package installer)

### Steps
1. Clone this repository:
   ```bash
   git clone https://github.com/Acekorneya/POK_Hotel_Scraper.git
   cd POK_Hotel_Scraper
   ```

2. Install the required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
1. Run the script:
   ```bash
   python hotel_scraper.py
   ```

2. Enter the state or country in the GUI input field.

3. Specify the fetch delay in seconds (default is 2 seconds).

4. Click on "Start Scraping" to begin the data extraction process.

5. Monitor the progress in the GUI and stop the process anytime by clicking "Stop Scraping".

## GUI Interface
- **State or Country:** Enter the state or country you want to scrape hotel data for.
- **Fetch Delay (seconds):** Set the delay between fetch requests to avoid being blocked by the server.
- **Start Scraping:** Button to start the scraping process.
- **Stop Scraping:** Button to stop the scraping process.
- **Progress:** Displays the current status and progress of the scraping process.

## Dependencies
- requests
- BeautifulSoup
- pandas
- tkinter
- xlsxwriter

To install these dependencies, you can use:
```bash
pip install requests beautifulsoup4 pandas xlsxwriter
```

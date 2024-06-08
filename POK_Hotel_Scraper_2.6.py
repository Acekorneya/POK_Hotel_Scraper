# Check if modules are available
try:
    import requests
    from bs4 import BeautifulSoup
    import pandas as pd
    import tkinter as tk
    from tkinter import ttk
    from tkinter import messagebox
    from tkinter import filedialog
    import time
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from urllib.parse import unquote 
    import threading
    import os
    import random
    import re
except ImportError as e:
    print(f"Missing module: {e}")
    exit(1)
headers = [
    {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0'},
    {'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:61.0) Gecko/20100101 Firefox/61.0'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.3'},
    {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.3'},
    {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/11.1.2 Safari/605.1.15'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.3 Edge/17.17134'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko'},
    {'User-Agent': 'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)'},
    {'User-Agent': 'Opera/9.80 (Windows NT 6.1; WOW64; U; en) Presto/2.10.229 Version/11.62'},
    {'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 12_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/12.0 Mobile/15E148 Safari/604.1'},
    {'User-Agent': 'Mozilla/5.0 (Linux; Android 8.0.0; SM-G960U) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.84 Mobile Safari/537.36'},
    {'User-Agent': 'Mozilla/5.0 (iPad; CPU OS 11_0 like Mac OS X) AppleWebKit/604.1.34 (KHTML, like Gecko) Version/11.0 Mobile/15A5341f Safari/604.1'},
    {'User-Agent': 'Mozilla/5.0 (Windows Phone 10.0; Android 6.0.1; Microsoft; Lumia 950) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Mobile Safari/537.3 Edge/15.15254'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:47.0) Gecko/20100101 Firefox/47.0'},
    {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/601.7.7 (KHTML, like Gecko) Version/9.1.2 Safari/601.7.7'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.111 Safari/537.36'},
    {'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:15.0) Gecko/20100101 Firefox/15.0.1'},
    {'User-Agent': 'Mozilla/5.0 (CrKey armv7l 1.5.16041) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.0 Safari/537.36'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.135 Safari/537.36 Edge/12.246'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:54.0) Gecko/20100101 Firefox/54.0'},
    {'User-Agent': 'Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; Win64; x64; Trident/6.0)'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.3'},
    {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Ubuntu Chromium/58.0.3029.110 Chrome/58.0.3029.110 Safari/537.3'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 6.2; Win64; x64; rv:16.0) Gecko/16.0 Firefox/16.0'},
    {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36'},
]
state_abbreviations = {
    'Alabama': 'AL',
    'Alaska': 'AK',
    'Arizona': 'AZ',
    'Arkansas': 'AR',
    'California': 'CA',
    'Colorado': 'CO',
    'Connecticut': 'CT',
    'Delaware': 'DE',
    'Florida': 'FL',
    'Georgia': 'GA',
    'Hawaii': 'HI',
    'Idaho': 'ID',
    'Illinois': 'IL',
    'Indiana': 'IN',
    'Iowa': 'IA',
    'Kansas': 'KS',
    'Kentucky': 'KY',
    'Louisiana': 'LA',
    'Maine': 'ME',
    'Maryland': 'MD',
    'Massachusetts': 'MA',
    'Michigan': 'MI',
    'Minnesota': 'MN',
    'Mississippi': 'MS',
    'Missouri': 'MO',
    'Montana': 'MT',
    'Nebraska': 'NE',
    'Nevada': 'NV',
    'New Hampshire': 'NH',
    'New Jersey': 'NJ',
    'New Mexico': 'NM',
    'New York': 'NY',
    'North Carolina': 'NC',
    'North Dakota': 'ND',
    'Ohio': 'OH',
    'Oklahoma': 'OK',
    'Oregon': 'OR',
    'Pennsylvania': 'PA',
    'Rhode Island': 'RI',
    'South Carolina': 'SC',
    'South Dakota': 'SD',
    'Tennessee': 'TN',
    'Texas': 'TX',
    'Utah': 'UT',
    'Vermont': 'VT',
    'Virginia': 'VA',
    'Washington': 'WA',
    'West Virginia': 'WV',
    'Wisconsin': 'WI',
    'Wyoming': 'WY'
}

def to_proper_case(s):
    return ' '.join([word.capitalize() for word in s.split()])

def is_us_state(name):
    return name in state_abbreviations
    
class HotelScraper:
    def __init__(self, root, state_entry, time_delay_entry, progress_text):
        self.root = root
        self.state_entry = state_entry
        self.time_delay_entry = time_delay_entry
        self.progress_text = progress_text
        self.stop_scraping_flag = False
        self.is_stopping = False
        self.visited_hotel_urls = set()
        self.workers = os.cpu_count()
        self.visited_city_urls = set()
        self.max_retries = 5
    
    def make_request(self, url):
        retries = 0
        while retries < self.max_retries:
            try:
                response = requests.get(url, headers=random.choice(headers))
                response.raise_for_status()  # Will raise HTTPError for bad responses
                return response
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 504:  # Gateway Timeout
                    retries += 1
                    print(f"Gateway Timeout. Retrying {retries}/{self.max_retries}")
                    time.sleep(2)  # Wait for 2 seconds before retrying
                else:
                    raise  # Re-raise the exception for other types of errors
            except requests.exceptions.RequestException as e:
                print(f"An error occurred: {e}")
                retries += 1
                print(f"Retrying {retries}/{self.max_retries}")
                time.sleep(2)  # Wait for 2 seconds before retrying
        return None  # Return None if max_retries reached
      
    def stop_scraping(self):
        self.stop_scraping_flag = True
        self.is_stopping = True  # Set the new flag to True
        self.progress_text.set("Stopping... Please wait.")  # Update the progress message
        self.root.update_idletasks()  # Update Tkinter widgets
        self.root.update()
    
    def fetch_max_pages_for_city(self, city_url):
        try:
            # Use the session object for making the request
            response = self.make_request(f"{city_url}?pg=1")        
            # If max_retries are reached, make_request will return None
            if response is None:
                print("Max retries reached. Could not fetch max pages for city.")
                return 1
            soup = BeautifulSoup(response.text, 'html.parser')
            pagination = soup.find('ul', {'class': 'pagination'})
            
            if pagination:
                last_li = pagination.find_all('li')[-1]
                max_pages = int(last_li.text)
            else:
                max_pages = 1
            
            return max_pages
        except Exception as e:
            print(f"An error occurred while determining max_pages: {e}")
            return 1
    
    # Fetch URLs from directory pages
    def fetch_directory_urls(self, directory_url):
        hotel_urls = []
        try:
            response = self.make_request(directory_url)
            if response is None:
                return []
            soup = BeautifulSoup(response.text, 'html.parser')

            for link in soup.find_all('a', {'class': 'title'}):
                hotel_url = link.get('href')
                full_url = f"https://www.travelweekly.com{hotel_url}"
                if full_url not in self.visited_hotel_urls:
                    hotel_urls.append(full_url)
                    self.visited_hotel_urls.add(full_url)
        except Exception as e:
            print(f"An error occurred: {e}")
        return hotel_urls
    
    def fetch_all_directory_urls_for_city(self, city_url, workers):
        max_pages = self.fetch_max_pages_for_city(city_url)
        all_hotel_urls = []
        if self.stop_scraping_flag:
            return all_hotel_urls
        with ThreadPoolExecutor(max_workers=workers) as executor:
            results = list(executor.map(self.fetch_directory_urls, [f"{city_url}?pg={page}" for page in range(1, max_pages + 1)]))
            for result in results:
                if self.stop_scraping_flag:
                    break
                if result:
                    all_hotel_urls.extend(result)
        return all_hotel_urls
    
    def fetch_hotel_data(self, url, City_Country):
        hotel_data = []
        try:
            if self.stop_scraping_flag:
                return []
            response = self.make_request(url)
            if response is None:
                return []
            soup = BeautifulSoup(response.text, 'html.parser')

            hotel_name_tag = soup.find('h1', {'class': 'title-xxxl'})
            hotel_name = hotel_name_tag.text.strip() if hotel_name_tag else ""

            room_count_tag = soup.find('td', {'class': 'rooms'})
            room_count = room_count_tag.text.strip() if room_count_tag else ""

            address_tag = soup.find('div', {'class': 'address'})
            main_address, state, country, postal_code = "", "", "", ""
            
            # Extract address components
            if address_tag:
                address_parts = list(address_tag.stripped_strings)
                if "View Large Map" in address_parts:
                    address_parts.remove("View Large Map")
                
                main_address = address_parts[0].strip() if len(address_parts) > 0 else ""
                state_country_zip = address_parts[1].strip() if len(address_parts) > 1 else ""
                state_country_zip_parts = state_country_zip.split(',')
                state = state_country_zip_parts[0].strip() if len(state_country_zip_parts) > 0 else ""
                country = state_country_zip_parts[1].strip() if len(state_country_zip_parts) > 1 else ""
                postal_code = state_country_zip_parts[2].strip() if len(state_country_zip_parts) > 2 else ""

            # Concatenate the various address components into a single string
            full_address = []
            if main_address:
                full_address.append(main_address)
            if state:
                full_address.append(state)
            if country:
                full_address.append(country)
            if postal_code:
                full_address.append(postal_code)
            full_address_str = ', '.join(full_address)

            contact_div = soup.find('div', {'class': 'contact'})
            phone_number, fax, email = "", "", ""
            if contact_div:
                phone_number = contact_div.find('b', string='Phone:').find_next_sibling(string=True).strip().replace("+", "") if contact_div.find('b', string='Phone:') else ""
                fax = contact_div.find('b', string='Fax:').find_next_sibling(string=True).strip().replace("+", "") if contact_div.find('b', string='Fax:') else ""
                email_tag = contact_div.find('a', {'title': 'Hotel E-mail'})
                email = email_tag['href'].replace('mailto:', '').strip() if email_tag else ""

            details_div = soup.find('div', {'class': 'col-xs-6'})
            year_built, year_renovated = "", ""
            if details_div:
                details_list = details_div.find('ul')
                if details_list:
                    for li in details_list.find_all('li'):
                        label_span = li.find('span', {'class': 'label'})
                        if label_span:
                            label_text = label_span.text.strip()
                            if "Year Built:" in label_text:
                                year_built = li.text.replace("Year Built:", "").strip()
                            elif "Year Last Renovated:" in label_text:
                                year_renovated = li.text.replace("Year Last Renovated:", "").strip()

            hotel_data.append({
                'Hotel Name': hotel_name,
                'Room Count': room_count,
                'Phone Number': phone_number,
                'Fax': fax,
                'Address': full_address_str,
                'City/Country': City_Country,
                'Email': email,
                'Year Built': year_built,
                'Year Renovated': year_renovated
            })

        except Exception as e:
            print(f"An error occurred: {e}")

        return hotel_data  
    def fetch_data_from_multiple_urls(self, url_list, workers, time_delay, city):
        all_hotel_data = []

        # Chunk URLs into smaller groups for efficient fetching
        chunk_size = 100  # Adjust based on your needs
        url_chunks = [url_list[i:i + chunk_size] for i in range(0, len(url_list), chunk_size)]

        try:
            with ThreadPoolExecutor(max_workers=workers) as executor:
                for url_chunk in url_chunks:
                    if self.stop_scraping_flag:
                        break  # Stop if the flag is set
                    results = list(executor.map(lambda url: self.fetch_hotel_data(url, city), url_chunk))
                    for hotel_data in results:
                        if self.stop_scraping_flag:
                            break  # Stop if the flag is set
                        if hotel_data:
                            all_hotel_data.extend(hotel_data)
                    time.sleep(time_delay)  # Delay between chunks
        except Exception as e:
            print(f"An error occurred: {e}")

        return all_hotel_data
    # Function to fetch all city URLs for a given state
    def fetch_city_urls(self, state_url, time_delay):
        city_urls = []
        try:
            response = self.make_request(state_url)
            if response is None:
                return []
            soup = BeautifulSoup(response.text, 'html.parser')
            city_div = soup.find('div', {'class': 'search-list'})
            if city_div:
                for city_li in city_div.find_all('li'):
                    if self.stop_scraping_flag:
                        break
                    city_a_tag = city_li.find('a')
                    if city_a_tag:
                        city_url = city_a_tag.get('href')
                        if city_url:
                            full_url = f"https://www.travelweekly.com{city_url}"
                            if full_url not in self.visited_city_urls:  # Check if city URL has been visited
                                city_urls.append(full_url)
                                self.visited_city_urls.add(full_url)  # Add to visited set
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                messagebox.showwarning("Invalid Input", "Please enter a valid state or country.")
                return []
            else:
                print(f"An error occurred: {e}")
                return []
        time.sleep(time_delay)
        return city_urls

    # New function to handle scraping for a single city
    
    def scrape_city_data(self, city_url, workers, time_delay):
        try:
            if '#' in city_url:
                return []
            
            city_name_encoded = city_url.split('/')[-1]
            hidden_city = unquote(city_name_encoded).replace('-', ' ')

            print(f"City URL before fetching directory: {city_url}")

            all_hotel_urls = self.fetch_all_directory_urls_for_city(city_url, workers)
            city_hotel_data = self.fetch_data_from_multiple_urls(all_hotel_urls, workers, time_delay, hidden_city)

            return city_hotel_data
        except Exception as e:
            print(f"An error occurred while scraping city: {e}")
            return []
    
    def validate_time_delay(self, time_delay_str):
        try:
            time_delay = float(time_delay_str)
            if time_delay < 0:
                messagebox.showwarning("Invalid Input", "Please enter a valid time delay in seconds.")
                return None
        except ValueError:
            messagebox.showwarning("Invalid Input", "Please enter a valid time delay in seconds.")
            return None
        return time_delay
    
    def save_data_to_excel(self, all_hotel_data, target_state=None):
        df = pd.DataFrame(all_hotel_data).drop_duplicates()

        if target_state:
            df['Extracted State'] = df['Address'].apply(lambda x: re.search(r'\b[A-Z]{2}\b', str(x)).group() if re.search(r'\b[A-Z]{2}\b', str(x)) else '')
            df = df[df['Extracted State'].str.contains(target_state, case=True, na=False)]

        if df.empty:
            messagebox.showwarning("Data Issue", "No hotels were found.")
            return

        df['Room Count'] = pd.to_numeric(df['Room Count'], errors='coerce')
        df = df[df['Room Count'] > 40]
        
        if df.empty:
            messagebox.showwarning("Data Issue", "No hotels were found after filtering.")
            return

        column_order = ['Hotel Name', 'Room Count', 'Phone Number', 'Fax', 'City/Country', 'Address', 'Email', 'Year Built', 'Year Renovated']
        df = df[column_order]

        while True:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if not file_path:
                break
            try:
                with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                    worksheet = writer.sheets['Sheet1']
                    worksheet.set_column('J:J', None, None, {'hidden': 1})
                break
            except PermissionError:
                retry = messagebox.askretrycancel("Permission Denied", "The file is currently open. Please close it before proceeding.")
                if not retry:
                    break
            except Exception as e:
                messagebox.showerror("An error occurred", str(e))
                break  
    def start_scraping(self):
        self.stop_scraping_flag = False  # Reset the flag
        self.visited_hotel_urls.clear()
        self.visited_city_urls.clear()
        self.progress_text.set("Initializing...")  # Set initial progress text
        self.root.update_idletasks()  # Update Tkinter widgets
        self.root.update()

        start_time = time.time()

        try:
            state = self.state_entry.get().strip()
            state = to_proper_case(state)  # Convert to proper case
            state_url = f"https://www.travelweekly.com/Hotels/{state.replace(' ', '-')}"

            time_delay_str = self.time_delay_entry.get().strip()
            time_delay = self.validate_time_delay(time_delay_str)

            if not state:
                messagebox.showwarning("Invalid Input", "Please enter valid input for all fields.")
                return

            initial_city_urls = self.fetch_city_urls(state_url, time_delay)
            
            if not initial_city_urls:
                messagebox.showwarning("Data Issue", "No cities were found for the given state or country.")
                return
            
            all_city_urls = initial_city_urls.copy()

            all_hotel_data = []

            self.progress_text.set("Initializing...")
            self.root.update_idletasks()
            
            with ThreadPoolExecutor(max_workers=self.workers) as executor:
                future_to_city = {}
                for city in all_city_urls:
                    if self.stop_scraping_flag:
                        executor.shutdown(wait=False, cancel_futures=True)
                        break
                    future = executor.submit(self.scrape_city_data, city, self.workers, time_delay)
                    future_to_city[future] = city

                for future in as_completed(future_to_city):
                    if self.stop_scraping_flag:
                        executor.shutdown(wait=False, cancel_futures=True)
                        break
                    city_data = future.result()
                    if city_data:
                        all_hotel_data.extend(city_data)

                    completed_city = future_to_city[future]
                    completed_city_name = unquote(completed_city.split('/')[-1]).replace('-', ' ')
                    self.progress_text.set(f"Hotel data for {completed_city_name} has been fetched.")
                    self.root.update_idletasks()
                    self.root.update()

            # Determine if the input is a U.S. state
            if is_us_state(state):
                target_state = state_abbreviations.get(state, state)
            else:
                target_state = None

            # Save data to Excel
            if target_state:
                threading.Thread(target=self.save_data_to_excel, args=(all_hotel_data, target_state)).start()
            else:
                threading.Thread(target=self.save_data_to_excel, args=(all_hotel_data, )).start()

            end_time = time.time()
            total_time = round(end_time - start_time, 2)
            messagebox.showinfo("Success", f"Data has been successfully scraped.\nTotal time taken: {total_time} seconds.")

            self.progress_text.set("Ready to Fetch Hotels")

        except Exception as e:
            messagebox.showerror("An error occurred", str(e))
            self.progress_text.set("Ready to Fetch Hotels")
            
# GUI setup
root = tk.Tk()
root.title("POK Hotel Data Scraper 2.6")

frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# State or Country Entry
state_label = ttk.Label(frame, text="State or Country:")
state_label.grid(row=0, column=0, sticky=tk.W)
state_entry = ttk.Entry(frame, width=30)
state_entry.grid(row=0, column=1)

# Time Delay Entry
time_delay_label = ttk.Label(frame, text="Fetch Delay (seconds):")
time_delay_label.grid(row=1, column=0, sticky=tk.W)
time_delay_entry = ttk.Entry(frame, width=10)
time_delay_entry.grid(row=1, column=1)
time_delay_entry.insert(0, "2")

# Progress text
progress_text = tk.StringVar()
progress_text.set("Ready to Fetch Hotels")

# Initialize HotelScraper object
scraper = HotelScraper(root, state_entry, time_delay_entry, progress_text)  

# Start Scraping Button
scrape_button = ttk.Button(frame, text="Start Scraping", command=lambda: threading.Thread(target=scraper.start_scraping).start())
scrape_button.grid(row=2, columnspan=2)

# Stop Scraping Button
stop_button = ttk.Button(frame, text="Stop Scraping", command=scraper.stop_scraping)
stop_button.grid(row=3, columnspan=2)
# Progress Label
progress_label = ttk.Label(frame, textvariable=progress_text)
progress_label.grid(row=4, columnspan=2)

root.mainloop()
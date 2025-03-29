from bs4 import BeautifulSoup
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import inspect
import pandas as pd
import re

class StockCrawler:
    debug = True

    def __init__(self, stock_code):
        self.raw_stock_code = stock_code  # Raw stock code without .TW
        # Set up Selenium Chrome driver
        chrome_options = Options()
        if not self.debug:
            chrome_options.add_argument('--headless')  # Run in headless mode (no GUI)
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
        # Specify path to chromedriver if not in PATH
        # service = Service('/path/to/chromedriver')
        self.driver = webdriver.Chrome(options=chrome_options)  # Use service=service if specifying path
    
    def _fetch_page(self, url):
        """
        Helper method to fetch page source with Selenium
        """
        try:
            self.driver.get(url)
            
            # Attempt to close advertisement if present
            try:
                # Option 1: Look for common close button (adjust XPath/ID as needed)
                close_button_xpath = "//button[@id='ats-interstitial-button']"
                WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, close_button_xpath))
                )
                close_button = self.driver.find_element(By.XPATH, close_button_xpath)
                close_button.click()
                print(f"Advertisement closed for {url}")
            except (TimeoutException, NoSuchElementException):
                print(f"No advertisement found or unable to close for {url}")
            
            # Wait for main content (table) to load after closing ad
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'table'))
            )
            return self.driver.page_source
        except Exception as e:
            print(f"Error fetching {url}: {e}")
            return None
    
    def get_revenue(self):
        """
        Fetch the last 36 months of revenue data from ShowSaleMonChart.asp
        """
        print(f"Start to fetch {inspect.currentframe().f_code.co_name.split('_')[1]} data")

        url = f"https://goodinfo.tw/tw/ShowSaleMonChart.asp?STOCK_ID={self.raw_stock_code}"
        html = self._fetch_page(url)
        if not html:
            return pd.DataFrame()

        soup = BeautifulSoup(html, 'html.parser')
        revenue_data = []
        revenue_table = soup.find('table', {'id': 'tblDetail'})
        if revenue_table:
            # Get rows from row0 to row35 (36 months)
            for i in range(36):  # 0 to 35 inclusive
                row = revenue_table.find('tr', {'id': f'row{i}'})
                if row:
                    cols = row.find_all('td')
                    if len(cols) >= 8:  # Ensure there are at least 8 <td> elements
                        year_month = cols[0].text.strip()  # First column is year/month
                        revenue = cols[7].text.strip().replace(',', '')  # 8th column (index 7) is revenue
                        if re.match(r'\d{4}/\d{2}', year_month):
                            year, month = map(int, year_month.split('/'))
                            revenue_data.append({
                                'Year': year,
                                'Month': month,
                                'Revenue': float(revenue) if revenue else None
                            })
                        else:
                            print(f"Invalid year/month format in row{i}: {year_month}")
                else:
                    print(f"Row {i} not found for stock {self.raw_stock_code}")
                    break  # Stop if a row is missing
        else:
            print(f"Revenue table 'tblDetail' not found for stock {self.raw_stock_code}")
        
        revenue_df = pd.DataFrame(revenue_data)
        if len(revenue_df) > 36:
            # Ensure only the latest 36 months are returned
            revenue_df = revenue_df.tail(36)
        return revenue_df
    
    def get_profit_ratio(self):
        """
        Fetch net profit margin data for the past 5 years from StockBzPerformance.asp
        """
        print(f"Start to fetch {inspect.currentframe().f_code.co_name.split('_')[1]} data")

        url = f"https://goodinfo.tw/tw/StockBzPerformance.asp?STOCK_ID={self.raw_stock_code}"
        html = self._fetch_page(url)
        if not html:
            return pd.DataFrame()

        soup = BeautifulSoup(html, 'html.parser')
        profit_data = []
        profit_table = soup.find('table', {'id': 'tblDetail'})
        if profit_table:
            # Iterate through rows (row0 to row5)
            for i in range(6):  # 0 to 5 inclusive
                row = profit_table.find('tr', {'id': f'row{i}'})
                if row:
                    cols = row.find_all('td')
                    if len(cols) >= 16:  # Ensure there are at least 16 <td> elements
                        year = cols[0].text.strip()  # First column is year
                        net_profit_margin = cols[15].text.strip().replace('%', '')  # 16th column (index 15) is net profit margin
                        if re.match(r'\d{4}', year):
                            # Skip if net profit margin is empty or invalid
                            if not net_profit_margin or net_profit_margin == '-' or net_profit_margin == '':
                                print(f"Skipping row{i} (Year {year}) due to empty or invalid profit margin")
                                continue
                            year = int(year)
                            profit_data.append({
                                'Year': year,
                                'Month': 12,  # Annual data, default to December
                                'Net Profit Margin': float(net_profit_margin)
                            })
                        else:
                            print(f"Invalid year format in row{i}: {year}")
                else:
                    print(f"Row {i} not found for stock {self.raw_stock_code}")
                    break
        
        profit_df = pd.DataFrame(profit_data)
        if profit_df.empty:
            print(f"No valid profit data parsed for stock {self.raw_stock_code}")
            return profit_df
        
        # Ensure we have exactly 5 years of valid data
        if len(profit_df) > 5:
            profit_df = profit_df.tail(5)  # Take the latest 5 years
        elif len(profit_df) < 5:
            print(f"Warning: Only {len(profit_df)} years of valid profit data found for stock {self.raw_stock_code}")
        
        return profit_df
    
    def get_pe_ratio(self):
        """
        Fetch the last 180 weeks of P/E ratio data from ShowK_ChartFlow.asp
        """
        print(f"Start to fetch {inspect.currentframe().f_code.co_name.split('_')[1]} data")        

        url = f"https://goodinfo.tw/tw/ShowK_ChartFlow.asp?RPT_CAT=PER&STOCK_ID={self.raw_stock_code}"
        
        try:
            self.driver.get(url)
            
            # Click the Expand year button
            five_years_button_xpath = "//input[@value='查5年']"
            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, five_years_button_xpath))
            )
            self.driver.find_element(By.XPATH, five_years_button_xpath).click()
            print(f"Clicked Expand Year button for stock {self.raw_stock_code}")
            
            # Wait for the table to update with 5 years of data
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'row180'))
            )
            html = self.driver.page_source
        except Exception as e:
            print(f"Error loading page or clicking button for {self.raw_stock_code}: {e}")
            return pd.DataFrame()

        soup = BeautifulSoup(html, 'html.parser')
        pe_data = []
        pe_table = soup.find('table', {'id': 'tblDetail'})
        if pe_table:
            for i in range(180):  # 0 to 179 inclusive
                row = pe_table.find('tr', {'id': f'row{i}'})
                if row:
                    cols = row.find_all('td')
                    if len(cols) >= 6:  # Ensure there are at least 6 <td> elements
                        week_str = cols[0].text.strip()  # First column is week (e.g., "25W13")
                        pe_ratio = cols[5].text.strip()  # 6th column (index 5) is P/E ratio
                        if re.match(r'\d{2}W\d{1,2}', week_str):
                            year, week = week_str.split('W')
                            year = "20" + year
                            pe_data.append({
                                'Year': year,
                                'Week': week,
                                'P/E Ratio': float(pe_ratio) if pe_ratio and pe_ratio != '-' else None
                            })
                        else:
                            print(f"Invalid week format in row{i}: {week_str}")
                else:
                    print(f"Row {i} not found for stock {self.raw_stock_code}")
                    break
        else:
            print(f"PE ratio table 'tblDetail' not found for stock {self.raw_stock_code}")
        
        pe_df = pd.DataFrame(pe_data)
        if len(pe_df) > 180:
            pe_df = pe_df.tail(180)  # Ensure only the latest 180 weeks
        return pe_df
    
    def get_current_stock_price(self):
        """
        Fetch the latest transaction price from StockDetail.asp using XPath
        """
        print(f"Start to fetch {inspect.currentframe().f_code.co_name.split('_')[1]} data")

        url = f"https://goodinfo.tw/tw/StockDetail.asp?STOCK_ID={self.raw_stock_code}"
        html = self._fetch_page(url)
        if not html:
            return pd.DataFrame([{'Year': datetime.now().year, 'Month': datetime.now().month, 'Price': None}])
        
        try:
            price_xpath = "/html/body/table[2]/tbody/tr[2]/td[3]/main/table/tbody/tr/td[1]/section/table/tbody/tr[3]/td[1]"
            price_element = self.driver.find_element(By.XPATH, price_xpath)
            price = price_element.text.strip().replace(',', '')
            current_date = datetime.now()
            return pd.DataFrame([{
                'Year': current_date.year,
                'Month': current_date.month,
                'Price': float(price) if price else None
            }])
        except Exception as e:
            print(f"Error finding price element for {self.raw_stock_code}: {e}")
            return pd.DataFrame([{
                'Year': datetime.now().year,
                'Month': datetime.now().month,
                'Price': None
            }])
    
    def __del__(self):
        """
        Clean up by closing the driver
        """
        self.driver.quit()
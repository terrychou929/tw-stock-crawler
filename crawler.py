import pandas as pd
from bs4 import BeautifulSoup
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class StockCrawler:
    def __init__(self, stock_code):
        self.raw_stock_code = stock_code  # Raw stock code without .TW
        # Set up Selenium Chrome driver
        chrome_options = Options()
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
                WebDriverWait(self.driver, 5).until(
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
        Fetch all monthly revenue data from the default view of ShowSaleMonChart.asp
        """
        url = f"https://goodinfo.tw/tw/ShowSaleMonChart.asp?STOCK_ID={self.raw_stock_code}"
        html = self._fetch_page(url)
        if not html:
            return pd.DataFrame()
        
        soup = BeautifulSoup(html, 'html.parser')
        revenue_data = []
        revenue_table = soup.find('table', {'id': 'tblChart'})
        if revenue_table:
            rows = revenue_table.find_all('tr')[1:]  # Skip header
            for row in rows:
                cols = row.find_all('td')
                if len(cols) >= 2:
                    year_month = cols[0].text.strip()
                    revenue = cols[1].text.strip().replace(',', '')
                    if re.match(r'\d{4}/\d{2}', year_month):
                        year, month = map(int, year_month.split('/'))
                        revenue_data.append({
                            'Year': year,
                            'Month': month,
                            'Revenue': float(revenue) if revenue else None
                        })
        else:
            print(f"Revenue table not found for stock {self.raw_stock_code}")
        
        return pd.DataFrame(revenue_data)
    
    def get_profit_ratio(self):
        """
        Fetch net profit margin data for the past 10 years from StockBzPerformance.asp
        """
        url = f"https://goodinfo.tw/tw/StockBzPerformance.asp?STOCK_ID={self.raw_stock_code}"
        html = self._fetch_page(url)
        if not html:
            return pd.DataFrame()
        
        soup = BeautifulSoup(html, 'html.parser')
        profit_data = []
        profit_table = soup.find('table', {'class': 'solid_1_padding_4_2_tbl'})
        if profit_table:
            rows = profit_table.find_all('tr')[2:]  # Skip headers
            for row in rows:
                cols = row.find_all('td')
                if len(cols) >= 11:
                    year = cols[0].text.strip()
                    net_profit_margin = cols[10].text.strip().replace('%', '')
                    if re.match(r'\d{4}', year):
                        year = int(year)
                        profit_data.append({
                            'Year': year,
                            'Month': 12,
                            'Net Profit Margin': float(net_profit_margin) if net_profit_margin and net_profit_margin != '-' else None
                        })
        else:
            print(f"Profit table not found for stock {self.raw_stock_code}")
        
        profit_df = pd.DataFrame(profit_data)
        if profit_df.empty:
            print(f"No profit data parsed for stock {self.raw_stock_code}")
            return profit_df
        
        ten_years_ago = datetime.now().year - 10
        profit_df = profit_df[profit_df['Year'] >= ten_years_ago]
        return profit_df
    
    def get_pe_ratio(self):
        """
        Fetch all 'Current PER' data from ShowK_ChartFlow.asp
        """
        url = f"https://goodinfo.tw/tw/ShowK_ChartFlow.asp?RPT_CAT=PER&STOCK_ID={self.raw_stock_code}"
        html = self._fetch_page(url)
        if not html:
            return pd.DataFrame()
        
        soup = BeautifulSoup(html, 'html.parser')
        pe_data = []
        pe_table = soup.find('table', {'id': 'tblFlowChart'})
        if pe_table:
            rows = pe_table.find_all('tr')[1:]  # Skip header
            for row in rows:
                cols = row.find_all('td')
                if len(cols) >= 2:
                    year_month = cols[0].text.strip()
                    pe_ratio = cols[1].text.strip()
                    if re.match(r'\d{4}/\d{2}', year_month):
                        year, month = map(int, year_month.split('/'))
                        pe_data.append({
                            'Year': year,
                            'Month': month,
                            'P/E Ratio': float(pe_ratio) if pe_ratio and pe_ratio != '-' else None
                        })
        else:
            print(f"PE ratio table not found for stock {self.raw_stock_code}")
        
        return pd.DataFrame(pe_data)
    
    def get_current_stock_price(self):
        """
        Fetch the latest transaction price from StockDetail.asp using XPath
        """
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
import sys
import pandas as pd
from crawler import StockCrawler
from concurrent.futures import ThreadPoolExecutor
import os

def main():
    # Check command line arguments
    if len(sys.argv) != 2:
        print("Please provide a stock code, e.g., python main.py 2330")
        sys.exit(1)
    
    stock_code = sys.argv[1]
    
    # Initialize crawler
    crawler = StockCrawler(stock_code)
    
    # Use multi-threading to fetch data
    with ThreadPoolExecutor(max_workers=3) as executor:
        # Execute three crawler tasks in parallel
        revenue_data = executor.submit(crawler.get_revenue)
        profit_ratio_data = executor.submit(crawler.get_profit_ratio)
        pe_ratio_data = executor.submit(crawler.get_pe_ratio)
        current_stock_price_data = executor.submit(crawler.get_current_stock_price)
        
        # Wait for all tasks to complete and retrieve results
        revenue = revenue_data.result()
        profit_ratio = profit_ratio_data.result()
        pe_ratio = pe_ratio_data.result()
        current_stock_price = current_stock_price_data.result()


if __name__ == "__main__":
    main()
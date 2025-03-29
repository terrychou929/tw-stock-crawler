# Taiwan Stock Crawler

## Overview
This project is designed to crawl financial data for individual Taiwan stocks over the past five years in one click

## Technique Utilized
- Python Threading
- Python Web Scrapping with Selenium Automation
- Python Openpyxl Excel File Generation
- Python Class Development
- Python Unit Testing
- Peer Coding with AI Model: Grok3

## Index to be crawled
- Monthly Revenue
- After Tax Gross Profit Ratio
- P/E Ratio
- Current Stock Price

Eventually it aims at calculating the upper and lower price bounds based on latest revenue, historic P/E ratio and net profit ratio, then saves the results as an Excel file.

## Installation and Usage
1. Install Python 3.8 or higher.
2. Create and activate a virtual environment:
3. Execute the main.py with stock code as the first argument
4. Check the generated file in output folder

```bash
python -m venv venv
# Windows
venv\Scripts\activate
# Mac/Linux
source venv/bin/activate
# Install dependency
pip install -r requirements.txt
# Execution
python main.py {stock_code}
```
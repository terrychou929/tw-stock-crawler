# Taiwan Stock Crawler

## Overview
This project is designed to crawl financial data for individual Taiwan stocks over the past five years in one click

## Index to be crawled
- Monthly Revenue
- After Tax Gross Profit Ratio
- P/E Ratio
- Current Stock Price

Eventually it aims at calculating the upper and lower price bounds based on the historic P/E ratio and saves the results as an Excel file.

## Installation
1. Install Python 3.8 or higher.
2. Create and activate a virtual environment:

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
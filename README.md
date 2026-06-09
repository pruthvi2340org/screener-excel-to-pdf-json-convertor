# Excel Financial Data Parser

A Python-based tool to extract and convert financial data from Excel files into structured PDF and JSON formats. Designed specifically for parsing financial statements and company data from standardized Excel templates.

## Features

- **Dual Output Formats**: Extract data to both PDF and JSON formats
- **Intelligent Section Detection**: Automatically identifies and categorizes different financial sections
- **Clean Data Processing**: Removes URLs, filters empty rows, and formats dates consistently
- **Professional PDF Output**: Creates well-formatted, landscape-oriented PDF reports with styled tables
- **Structured JSON Output**: Generates hierarchical JSON with organized sections and metrics
- **Combined Sections**: Intelligently merges related data (e.g., Company Info + Metadata, Price + Derived data)

## Supported Financial Sections

The parser extracts and organizes the following sections:

1. **Company Information & Metadata**
   - Company name, current price, market capitalization
   - Basic company metadata

2. **Profit & Loss Statement**
   - Sales, expenses, operating profit
   - Net profit, tax, dividends

3. **Quarterly Results**
   - Quarterly sales, expenses, operating profit

4. **Balance Sheet**
   - Assets, liabilities, equity
   - Inventory, receivables, cash & bank

5. **Cash Flow Statement**
   - Operating, investing, and financing activities

6. **Price & Market Data**
   - Stock price information
   - Market-derived metrics

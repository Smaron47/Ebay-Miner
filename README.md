**Web Data Aggregator & Excel Updater – Documentation**

---

## Table of Contents

1. Project Overview
2. Key Features
3. Technology Stack & Dependencies
4. Environment Setup
5. Directory & Configuration Files
6. Script Workflow & Data Flow
7. Function Descriptions

   * `webdata(inp)`
8. Excel Integration Details
9. Rate Limiting & Delays
10. Error Handling & Logging
11. Customization & Extensibility
12. Troubleshooting & FAQs
13. SEO Keywords
14. Author & License

---

## 1. Project Overview

This Python script automates the process of:

1. Reading search terms from an Excel file (column C).
2. Performing price and link extraction from Google UK and eBay UK via Selenium.
3. Filtering eBay listings to only those shipped from the United Kingdom.
4. Writing the extracted links and prices back into designated columns (F–I) in a new Excel workbook.

It is designed to help e-commerce analysts and market researchers quickly gather price comparisons across platforms.

---

## 2. Key Features

* **Batch Processing**: Reads multiple search queries from an Excel column.
* **Multi-Platform Scraping**: Extracts data from both Google search results and eBay listings.
* **Location Filtering**: Only includes eBay items shipped from the UK.
* **Excel Integration**: Reads and writes data using `openpyxl`.
* **Configurable Output**: Saves results into a user-specified output file.

---

## 3. Technology Stack & Dependencies

* **Python 3.7+**
* **Selenium**: Browser automation for web scraping.
* **openpyxl**: Reading and writing Excel `.xlsx` files.

Install dependencies:

```bash
pip install selenium openpyxl
```

Additionally, download the **ChromeDriver** matching your Chrome version and ensure it is in your `PATH`.

---

## 4. Environment Setup

1. **Chrome Browser**: Install latest stable release.
2. **ChromeDriver**: Place `chromedriver` executable in your system `PATH`.
3. **Python Packages**: Install via pip as above.
4. **Excel File**: Create an input workbook where column C contains search terms.

---

## 5. Directory & Configuration Files

```
/WebScraperExcel/
├── scraper.py           # Main script
├── input.xlsx           # Example input file
├── output.xlsx          # Generated output file
└── README.md            # This documentation
```

No additional configuration files are required.

---

## 6. Script Workflow & Data Flow

1. **User Input**: Prompts for the input file path and desired output file name.
2. **Excel Load**: Opens the input workbook and reads all values from column C, skipping header.
3. **Loop Through Queries**:

   * Calls `webdata(query)` to scrape data.
   * Receives four lists: `li`/`pr` for eBay UK links and prices, `liuk`/`pruk` for Google UK results.
4. **Data Formatting**: Joins lists into newline-delimited strings for each row.
5. **Excel Write**: Writes Google links to column F, Google prices to G, eBay links to H, eBay prices to I.
6. **Save Workbook**: Saves the modified data into the output file.

---

## 7. Function Descriptions

### `webdata(inp)`

Performs web scraping for a single search query `inp`:

1. **Google UK Search**

   * Navigates to `https://www.google.co.uk/search?q={inp}`.
   * Finds result blocks containing “£”.
   * Extracts the link and euro price ("€" marker) from each block, storing in `liuk`, `pruk`.
2. **eBay UK Search**

   * Navigates to the eBay search URL with 240 results per page.
   * Filters `<li>` elements having location text “from United Kingdom”.
   * Extracts each item’s link (`<a>` href) and price text.
   * Appends to `li`, `pr`.
3. **Return Values**: Four lists:

   * `li`: eBay UK links
   * `pr`: eBay UK prices
   * `liuk`: Google UK links
   * `pruk`: Google UK prices

---

## 8. Excel Integration Details

* Uses `openpyxl.load_workbook` to open the input file.
* Reads `ws['C']` values into a Python list, dropping the header.
* Iterates rows via index `o` starting at 2 for writing.
* Writes results:

  * Column F (6): Google UK links
  * Column G (7): Google UK prices
  * Column H (8): eBay UK links
  * Column I (9): eBay UK prices
* Saves workbook with new filename.

---

## 9. Rate Limiting & Delays

* Built-in `time.sleep` calls ensure pages load before scraping.
* Google search: 2s delay after page load.
* eBay search: 3s delay before locating elements.
* 5s pause before closing driver to allow any JavaScript to finish.

Consider adjusting delays according to network speed.

---

## 10. Error Handling & Logging

* **No Results**: If no Google matches, inserts "Not Found" placeholders.
* **Element Not Found**: Try/except around eBay filtering ensures robust operation.
* **Driver Cleanup**: Always calls `driver.close()` to free resources.

Enhancement: integrate Python’s `logging` module for detailed runtime logs.

---

## 11. Customization & Extensibility

* **Search URL**: Modify regional domains (`.co.uk`) to suit other markets.
* **Result Filters**: Adapt XPath expressions or filtering logic for different page layouts.
* **Excel Columns**: Change target columns by adjusting `ws.cell(row, column)` indices.
* **Batch Size**: Loop through multiple sheets or extend to CSV inputs.

---

## 12. Troubleshooting & FAQs

* **ChromeDriver Mismatch**: Confirm driver version matches browser.
* **XPath Changes**: Web layouts may change; update XPaths accordingly.
* **Slow Networks**: Increase `time.sleep` durations.
* **Unicode Errors**: Ensure Excel file encoding is UTF-8.

---

## 13. SEO Keywords

```
Python Excel scraper
Selenium openpyxl tutorial
eBay price extractor
Google search scraper
Excel web automation
data aggregator script
web scraping Python
eBay Selenium example
Python Excel scraper
Selenium openpyxl tutorial
eBay price extractor
Google search scraper
Excel web automation
data aggregator script
web scraping Python
eBay Selenium example
```

---

## 14. Author & License

**Author:** Smaron Biswas
**Date:** 2025
**License:** MIT License

Feel free to reuse and adapt this script for your web data collection and Excel reporting needs.

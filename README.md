# ATK Business Data Scraper (GUI)

This is a Python-based desktop tool that allows users to scrape business data from the **ATK (Tax Administration of Kosovo)** website:  
[https://apps.atk-ks.org/BizPasiveApp/VatRegist/Index](https://apps.atk-ks.org/BizPasiveApp/VatRegist/Index)

The tool provides a simple and user-friendly **graphical interface** built with `Tkinter`, making it easy for anyone—even non-technical users—to input a business number and retrieve the related business details. The extracted data is then saved to an Excel file for easy access and reuse.

---

## Features

- Simple and clean **GUI interface**
- Retrieves data directly from the official ATK business registry page
- Saves all scraped data to an Excel file (`business_data.xlsx`)
- CAPTCHA is completed **manually by the user** to comply with anti-bot protection
- Works without requiring users to manually navigate a browser

---

## How It Works

1. The user opens the application and enters a **business number**.
2. The tool launches a Chrome browser using Selenium and navigates to the ATK VAT registration lookup page.
3. The tool pre-fills the business number into the search field.
4. A popup appears prompting the user to manually complete the CAPTCHA in the browser window.
5. After solving the CAPTCHA and clicking "OK" in the app, the tool automatically proceeds to fetch business data from the results.
6. Data is saved (or appended) to an Excel file in the project folder.

---

## Data Extracted

The scraper collects the following details:

- Nr. Fiskal
- Emri
- Adresa
- Qyteti
- Komuna
- Qendra Tatimore
- Nr. TVSh
- Statusi në TVSh
- Tjetër

---

## Technologies Used

- **Python**
- **Tkinter** – for the graphical interface
- **Selenium** – for browser automation and web scraping
- **OpenPyXL** – for handling Excel file creation and updates
- **ChromeDriver** – to drive the browser via Selenium

---

## Requirements

Make sure you have the following installed:

- Python 3.x
- Google Chrome browser
- ChromeDriver (compatible with your Chrome version)
- Python libraries:
  - `selenium`
  - `openpyxl`

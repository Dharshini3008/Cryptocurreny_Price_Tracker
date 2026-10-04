# crypto_tracker.py - Cryptocurrency Price Tracker
#
# 1. Selenium opens CoinMarketCap and reads the Top 10 coins
# 2. Pandas cleans the data
# 3. Every run is added to a history file (crypto_prices.csv and crypto_prices.xlsx)
# 4. You type a minimum price and the program shows the coins above it, sorted by price
#
# Install:  python -m pip install selenium pandas openpyxl
# Run:      python crypto_tracker.py            (add  --headless  to hide Chrome)

import os
import sys
from datetime import datetime

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.support.ui import WebDriverWait

URL = "https://coinmarketcap.com/"
CSV_FILE = "crypto_prices.csv"
EXCEL_FILE = "crypto_prices.xlsx"
COLUMNS = ["Date & Time", "Name", "Symbol", "Price (USD)", "Market Cap", "24h Change (%)"]

# ---------------------------------------------------------------------------
# JavaScript that runs INSIDE Chrome and returns the coin rows.
# It finds the table by its headings "Name" and "Price" (the website's class
# names change often, so we do not use them).
# ---------------------------------------------------------------------------
JS = r"""
// 1. find the table that has "Name" and "Price" headings
const table = [...document.querySelectorAll('table')].find(t =>
  t.tHead && /name/i.test(t.tHead.textContent) && /price/i.test(t.tHead.textContent));
if (!table) return [];

// 2. find which column is which, using the heading text
const heads = [...table.tHead.querySelectorAll('th')].map(th => th.textContent.trim().toLowerCase());
const col = test => heads.findIndex(test);
const RANK = col(h => h === '#'),        NAME = col(h => h.startsWith('name')),
      PRICE = col(h => h.startsWith('price')), CHANGE = col(h => h.startsWith('24h')),
      CAP = col(h => h.startsWith('market'));

// 3. read every row that has a rank number (the pinned top row has none, so it is skipped)
const result = [];
for (const row of table.tBodies[0].rows) {
  const rank = row.cells[RANK] ? row.cells[RANK].textContent.trim() : '';
  if (!/^\d+$/.test(rank)) continue;

  // name cell contains the coin name and its symbol as the last small texts
  const words = [...row.cells[NAME].querySelectorAll('*')]
    .filter(e => e.children.length === 0 && e.textContent.trim())
    .map(e => e.textContent.trim());

  // is the 24h change going down? (minus sign, red arrow icon, or red colour)
  const cell = row.cells[CHANGE];
  const text = cell.textContent.trim();
  const colour = getComputedStyle(cell.querySelector('span') || cell).color.match(/\d+/g);
  const down = /^[-−]/.test(text) || /caret-?down|arrow-?down|icon-down/i.test(cell.innerHTML)
               || (+colour[0] > 150 && +colour[1] < 120);

  result.push({Rank: +rank, Name: words[0], Symbol: words[words.length - 1],
               Price: row.cells[PRICE].textContent.trim(),
               Change_24h: (down ? '-' : '') + text.replace(/^[-−+]/, ''),
               Market_Cap: row.cells[CAP].textContent.trim()});
}
return result;
"""


def to_number(text):
    """'$1,321.36' -> 1321.36    '-0.78%' -> -0.78    '$1.7T' -> 1700000000000"""
    text = text.replace("$", "").replace(",", "").replace("%", "").replace("−", "-")
    multiplier = {"K": 1e3, "M": 1e6, "B": 1e9, "T": 1e12}
    if text[-1] in multiplier:
        return round(float(text[:-1]) * multiplier[text[-1]])
    return float(text)


def get_top_10(headless=False):
    """Open CoinMarketCap and return the Top 10 coins as a DataFrame."""
    options = webdriver.ChromeOptions()
    options.page_load_strategy = "eager"          # do not wait for every ad and image
    if headless:
        options.add_argument("--headless=new")    # Chrome without a window
    driver = webdriver.Chrome(options=options)    # Selenium downloads ChromeDriver by itself
    try:
        driver.set_page_load_timeout(40)
        try:
            driver.get(URL)
        except TimeoutException:
            driver.execute_script("window.stop();")   # page is slow: use what has loaded
        # explicit wait: keep checking (up to 30 seconds) until 10 coins are on the page
        rows = WebDriverWait(driver, 30).until(
            lambda d: len(d.execute_script(JS)) >= 10 and d.execute_script(JS))
    finally:
        driver.quit()                             # always close Chrome

    data = pd.DataFrame(rows).drop_duplicates("Rank").sort_values("Rank").head(10)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return pd.DataFrame({
        "Date & Time": now,
        "Name": data["Name"].values,
        "Symbol": data["Symbol"].values,
        "Price (USD)": [to_number(x) for x in data["Price"]],
        "Market Cap": [to_number(x) for x in data["Market_Cap"]],
        "24h Change (%)": [to_number(x) for x in data["Change_24h"]],
    })


def save_history(new_data):
    """Add the new rows to the end of the CSV and Excel files. Returns the full history."""
    if os.path.exists(CSV_FILE):
        history = pd.concat([pd.read_csv(CSV_FILE), new_data], ignore_index=True)
    else:
        history = new_data
    history.to_csv(CSV_FILE, index=False)
    try:
        history.to_excel(EXCEL_FILE, index=False)
    except PermissionError:
        print("Excel file is open. Close it to update it (the CSV file was saved).")
    return history


def main():
    try:
        new_data = get_top_10("--headless" in sys.argv)
    except TimeoutException:
        print("The coin table did not appear in time. Check your internet and try again.")
        return
    except WebDriverException as error:
        print("Could not read CoinMarketCap:", str(error).strip().splitlines()[0])
        print("If it says ERR_CONNECTION_TIMED_OUT, your network cannot reach the website.")
        return

    try:
        history = save_history(new_data)
    except PermissionError:
        print("Cannot write crypto_prices.csv. Close it if it is open in Excel.")
        return
    print("Historical crypto data saved successfully!")

    print("\nLatest 10 crypto records:")
    print(history.tail(10).to_string())

    # ask for a minimum price and filter the latest 10 coins
    try:
        minimum = float(input("\nEnter minimum price in USD: "))
    except ValueError:
        print("Please type a number, for example 1000.")
        return
    latest = history.tail(10)
    above = latest[latest["Price (USD)"] > minimum]
    print("\nCryptocurrencies above the selected price:")
    if above.empty:
        print("No cryptocurrency found above", minimum)
        return
    print(above.to_string())

    print("\nSorted cryptocurrencies by price:")
    print(above.sort_values("Price (USD)", ascending=False).to_string())
    print("\nTotal cryptocurrencies found:", len(above))


if __name__ == "__main__":
    main()
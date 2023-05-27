import numpy as np
import pandas as pd
import xlsxwriter
import requests
import math
from secret import IEX_CLOUD_API_TOKEN

stocks = pd.read_csv("constituents_csv.csv")
symbol = "AAPL"
api_url = f"https://cloud.iexapis.com/stable/stock/{symbol}/quote?token={IEX_CLOUD_API_TOKEN}"
data = requests.get(api_url).json()

dictionary = {
    "Ticker": symbol,
    "Stock Price": data["latestPrice"],
    "Market Capitalization": data["marketCap"],
    "Number of Shares to Buy": "N/A"
}

price = data["latestPrice"]
market_cap = data["marketCap"]
print(market_cap / 1000000000000)

my_columns = ["Ticker", "Stock Price", "Market Capitalization", "Number of Shares to Buy"]

final_dataframe = pd.DataFrame(columns=my_columns)

final_dataframe = pd.concat([
    final_dataframe,
    pd.Series([symbol, price, market_cap, "N/A"], index=my_columns)
], ignore_index=True)

print(final_dataframe)

for stock in stocks["Symbol"][:5]:
    print(stock)

for stock in stocks["Symbol"][:5]:
    api_url = f"https://cloud.iexapis.com/stable/stock/{stock}/quote?token={IEX_CLOUD_API_TOKEN}"
    data = requests.get(api_url).json()
    final_dataframe = pd.concat([
        final_dataframe,
        pd.Series([stock, data["latestPrice"], data["marketCap"], "N/A"], index=my_columns)
    ], ignore_index=True)

def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

symbol_groups = list(chunks(stocks["Symbol"], 100))
symbol_strings = []
for i in range(0, len(symbol_groups)):
    symbol_strings.append(",".join(symbol_groups[i]))

for symbol_string in symbol_strings[:1]:
    batch_api_call_url = f"https://cloud.iexapis.com/stable/stock/market/batch/?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}"
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(","):
        final_dataframe = pd.concat([
            final_dataframe,
            pd.Series(
                [
                    symbol,
                    data[symbol]["quote"]["latestPrice"],
                    data[symbol]["quote"]["marketCap"],
                    "N/A"
                ],
                index=my_columns
            )
        ], ignore_index=True)

print(final_dataframe)

portfolio_size = input("Enter the value of your portfolio:")
try:
    val = float(portfolio_size)
    print(val)
except ValueError:
    print("That's not a number! \n Try again:")
    portfolio_size = input("Enter the value of your portfolio:")
    val = float(portfolio_size)

position_size = 0
if len(final_dataframe) > 0:
    position_size = val / len(final_dataframe)
    for i in range(len(final_dataframe)):
        stock_price = final_dataframe["Stock Price"][i]
        if pd.notnull(stock_price):  # Skip stocks with missing prices
            final_dataframe.loc[i, "Number of Shares to Buy"] = math.floor(position_size / stock_price)
else:
    print("No data available in final_dataframe. Please check the data retrieval process.")


writer = pd.ExcelWriter("recommended_trades.xlsx", engine="xlsxwriter")
final_dataframe.to_excel(writer, "Recommended Trades", index=False)

background_color = "#0a0a23"
font_color = "#ffffff"

string_format = writer.book.add_format({
    "font_color": font_color,
    "bg_color": background_color,
    "border": 1
})

dollar_format = writer.book.add_format({
    "num_format": "$0.00",
    "font_color": font_color,
    "bg_color": background_color,
    "border": 1
})

integer_format = writer.book.add_format({
    "num_format": "0",
    "font_color": font_color,
    "bg_color": background_color,
    "border": 1
})

percent_format = writer.book.add_format({
    "num_format": "0.0%",
    "font_color": font_color,
    "bg_color": background_color,
    "border": 1
})

column_formats = {
    "A": ["Ticker", string_format],
    "B": ["Stock Price", dollar_format],
    "C": ["Market Capitalization", dollar_format],
    "D": ["Number of Shares to Buy", integer_format]
}

for column in column_formats.keys():
    writer.sheets["Recommended Trades"].set_column(f"{column}:{column}", 20, column_formats[column][1])
    writer.sheets["Recommended Trades"].write(f"{column}1", column_formats[column][0], string_format)

writer.close()
print("Recommended trades are saved in recommended_trades.xlsx")

import requests
import pandas as pd

# Fetch top 50 cryptocurrencies from CoinGecko
url = "https://api.coingecko.com/api/v3/coins/markets"
params = {
    'vs_currency': 'usd',
    'order': 'market_cap_desc',
    'per_page': 50,
    'page': 1,
    'sparkline': 'false'
}

response = requests.get(url, params=params)
data = response.json()

# Prepare the data for Excel
crypto_data = []
for coin in data:
    crypto_data.append({
        'Name': coin['name'],
        'Symbol': coin['symbol'],
        'Current Price (USD)': coin['current_price'],
        'Market Capitalization': coin['market_cap'],
        '24h Trading Volume': coin['total_volume'],
        'Price Change (24h, %)': coin['price_change_percentage_24h']
    })

# Convert to a DataFrame for easier manipulation
df = pd.DataFrame(crypto_data)
print(df.head())  # Check the first few rows

#Identify the top 5 cryptocurrencies by market cap
top_5_by_market_cap = df.nlargest(5, 'Market Capitalization')
print("Top 5 Cryptocurrencies by Market Cap:")
print(top_5_by_market_cap[['Name', 'Market Capitalization']])

#Calculate the average price of the top 50 cryptocurrencies.
avg_price = df['Current Price (USD)'].mean()
print(f"Average Price of the Top 50 Cryptocurrencies: ${avg_price:.2f}")

#Analyze the highest and lowest 24-hour percentage price change among the top 50
highest_change = df.loc[df['Price Change (24h, %)'].idxmax()]
lowest_change = df.loc[df['Price Change (24h, %)'].idxmin()]

print("Highest 24-hour Price Change:")
print(highest_change[['Name', 'Price Change (24h, %)']])

print("Lowest 24-hour Price Change:")
print(lowest_change[['Name', 'Price Change (24h, %)']])


#Create and Write to Excel:
from openpyxl import Workbook

# Create a new workbook and add a sheet
wb = Workbook()
ws = wb.active
ws.title = "Cryptocurrency Data"

# Add headers
headers = ['Name', 'Symbol', 'Current Price (USD)', 'Market Capitalization', '24h Trading Volume', 'Price Change (24h, %)']
ws.append(headers)

# Write data to the sheet
for row in crypto_data:
    ws.append([row['Name'], row['Symbol'], row['Current Price (USD)'], row['Market Capitalization'], row['24h Trading Volume'], row['Price Change (24h, %)']])

# Save the file
wb.save("crypto_data.xlsx")

#Set up Auto-Refresh:
import time

while True:
    # Fetch and save new data every 5 minutes
    response = requests.get(url, params=params)
    data = response.json()
    crypto_data = []
    for coin in data:
        crypto_data.append({
            'Name': coin['name'],
            'Symbol': coin['symbol'],
            'Current Price (USD)': coin['current_price'],
            'Market Capitalization': coin['market_cap'],
            '24h Trading Volume': coin['total_volume'],
            'Price Change (24h, %)': coin['price_change_percentage_24h']
        })

    df = pd.DataFrame(crypto_data)

    # Create and write to Excel as before
    wb = Workbook()
    ws = wb.active
    ws.title = "Cryptocurrency Data"
    ws.append(headers)
    for row in crypto_data:
        ws.append([row['Name'], row['Symbol'], row['Current Price (USD)'], row['Market Capitalization'], row['24h Trading Volume'], row['Price Change (24h, %)']])

    # Save file
    wb.save("crypto_data.xlsx")
    
    # Wait for 5 minutes before updating again
    time.sleep(300)

import streamlit as st
import requests
import pandas as pd
from openpyxl import Workbook
import time

# Function to fetch data from CoinGecko API
def fetch_data():
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
    
    return pd.DataFrame(crypto_data)

# Streamlit App Layout
st.title('Live Cryptocurrency Data')

# Display last updated timestamp
last_updated = st.empty()

# Button to trigger manual refresh
refresh_button = st.button("Refresh Data")

# Display the data table
data_placeholder = st.empty()

# Display analysis results
analysis_placeholder = st.empty()

# Function to display analysis
def display_analysis(df):
    # Identify the top 5 cryptocurrencies by market cap
    top_5_by_market_cap = df.nlargest(5, 'Market Capitalization')
    st.subheader("Top 5 Cryptocurrencies by Market Cap:")
    st.write(top_5_by_market_cap[['Name', 'Market Capitalization']])

    # Calculate the average price of the top 50 cryptocurrencies
    avg_price = df['Current Price (USD)'].mean()
    st.subheader(f"Average Price of the Top 50 Cryptocurrencies: ${avg_price:.2f}")

    # Analyze the highest and lowest 24-hour price change
    highest_change = df.loc[df['Price Change (24h, %)'].idxmax()]
    lowest_change = df.loc[df['Price Change (24h, %)'].idxmin()]
    
    st.subheader("Highest 24-hour Price Change:")
    st.write(highest_change[['Name', 'Price Change (24h, %)']])

    st.subheader("Lowest 24-hour Price Change:")
    st.write(lowest_change[['Name', 'Price Change (24h, %)']])

# Fetch and display live data initially
df = fetch_data()

# Show the data table and analysis
data_placeholder.dataframe(df)
display_analysis(df)

# Save data to Excel (optional)
def save_to_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Cryptocurrency Data"
    headers = ['Name', 'Symbol', 'Current Price (USD)', 'Market Capitalization', '24h Trading Volume', 'Price Change (24h, %)']
    ws.append(headers)
    for row in df.itertuples(index=False, name=None):
        ws.append(row)
    wb.save("crypto_data.xlsx")

# Save the file initially
save_to_excel(df)

# Live Update every 5 minutes (optional)
if refresh_button:
    while True:
        # Wait for 5 minutes before refreshing data
        time.sleep(300)

        # Fetch the latest data
        df = fetch_data()

        # Display the updated data
        data_placeholder.dataframe(df)
        display_analysis(df)

        # Save to Excel
        save_to_excel(df)

        # Update the timestamp
        last_updated.text(f"Last Updated: {time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime())}")

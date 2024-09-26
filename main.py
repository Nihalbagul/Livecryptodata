import requests
import pandas as pd
import xlwings as xw
import time



def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": "false"
    }

    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Failed to fetch data: {response.status_code}")
        return None


def create_dataframe(data):
    df = pd.DataFrame(data)
    df = df[['name', 'symbol', 'current_price', 'market_cap', 'total_volume', 'price_change_percentage_24h']]
    df.columns = ['Cryptocurrency Name', 'Symbol', 'Current Price (USD)', 'Market Cap', '24h Trading Volume', '24h Price Change (%)']
    return df



def update_excel(dataframe):
    # Write dataframe to Excel using pandas without needing the Excel app
    with pd.ExcelWriter("crypto_analysis.xlsx", engine="openpyxl", mode='w') as writer:
        dataframe.to_excel(writer, sheet_name="Sheet1", index=False)
    print("Data successfully written to crypto_analysis.xlsx")



def perform_analysis(df):
    top_5_by_market_cap = df.nlargest(5, 'Market Cap')
    avg_price = df['Current Price (USD)'].mean()
    highest_24h_change = df.nlargest(1, '24h Price Change (%)')
    lowest_24h_change = df.nsmallest(1, '24h Price Change (%)')

    return top_5_by_market_cap, avg_price, highest_24h_change, lowest_24h_change



def main_loop():
    while True:
        data = fetch_crypto_data()
        if data:
            crypto_df = create_dataframe(data)

           
            update_excel(crypto_df)

        
            top_5, avg_price, highest_change, lowest_change = perform_analysis(crypto_df)

            print("\nTop 5 Cryptocurrencies by Market Cap:\n", top_5)
            print("\nAverage Price of Top 50 Cryptocurrencies: $", avg_price)
            print("\nHighest 24h Percentage Price Change:\n", highest_change)
            print("\nLowest 24h Percentage Price Change:\n", lowest_change)

        # Wait 5 minutes before next update
        time.sleep(300)


if __name__ == "__main__":
    main_loop()

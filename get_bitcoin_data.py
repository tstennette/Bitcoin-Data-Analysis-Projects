import json
import pandas as pd
import pprint
import pytz
from datetime import datetime
from dateutil import parser
from requests import Session

assets = ['bitcoin'] # only asset(s) data is retrieved for 
api = '935cfaa3-eb85-490c-9a9c-aba1bbc065b8' # Coinmarketcap API key
url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest' # Coinmarketcap API URL

def update_invest_file(data):
    # Read the DataFrames
    df1 = pd.read_excel('Current BTC Holdings.xlsx')
    df2 = pd.read_excel('BTC Performance Data.xlsx')
    
    # Filter data to only list total values
    df_list = df1[df1['Location'] == 'TOTAL']
    
    # Get individual values
    ID = len(df2)
    asset = df_list['Asset'].values[0]
    quantity = df_list['Quantity'].values[0]
    price = data[2]
    total_value = price * quantity
    total_invested = df_list['Total Invested'].values[0]
    avg_purch_price = total_invested / quantity
    unrealized_return = total_value - total_invested
    unrealized_ROI = (total_value - total_invested) / total_invested
    timestamp = data[3]
    
    # Create a flat list
    final_list = [ID, asset, quantity, price, total_value, total_invested, avg_purch_price, unrealized_return, unrealized_ROI, timestamp]
    
    # Append the list to the DataFrame
    df2.loc[len(df2)] = final_list
    
    # Write the DataFrame to the Excel file
    df2.to_excel('BTC Performance Data.xlsx', index=False)

def get_json_response(asset): # returns json text containing specific BTC data 
    
    parameters = { 'slug': asset, 'convert': 'USD' } # API parameters to pass in 
    headers = {
        'Accepts': 'application/json',
        'X-CMC_PRO_API_KEY': api
    } # Headers for API request

    session = Session() # Create new session object to manage API requests
    session.headers.update(headers) # Update with specified headers created above
    response = session.get(url, params=parameters) # getting response from the API

    info = json.loads(response.text)
    return info

def convert_time(time): 
    
    # Convert timestamp to Pacific Standard Time (PST)
    timestamp_local = parser.parse(time).astimezone(pytz.timezone('US/Pacific'))
    
    # Format timestamp to display year-month-day hour:min:sec
    formatted_timestamp = timestamp_local.strftime('%Y-%m-%d %H:%M:%S')
    return formatted_timestamp

def get_new_record(info, id): # returns new record in list form
  
    # extract data from json text response
    data = info['data']['1']
    name = data['name']
    symbol = data['symbol']
    rank = data['cmc_rank']
    total_supply = data['total_supply']
    circulating_supply = data['circulating_supply']
    market_cap = data['quote']['USD']['market_cap']
    price = data['quote']['USD']['price']
    market_cap_dominance = data['quote']['USD']['market_cap_dominance']
    percent_change_1h = data['quote']['USD']['percent_change_1h']
    percent_change_24h = data['quote']['USD']['percent_change_24h']
    volume_24h = data['quote']['USD']['volume_24h']
    volume_change_24h = data['quote']['USD']['volume_change_24h']
    timestamp = convert_time(info['status']['timestamp'])


    # update BTC performance and Current BTC holdings files 
    invest_data = [id, symbol, price, timestamp]
    update_invest_file(invest_data)
          
    record_list = [id, name, symbol, rank, total_supply, circulating_supply, market_cap, price, market_cap_dominance, percent_change_1h, percent_change_24h, volume_24h, volume_change_24h, timestamp]   
    return record_list
    
def update_data (df): # Function to append new row to dataframe

    i = 0
    while i < len(assets):

        info = get_json_response(assets[i])
        
        # Extract the data from json response
        df.loc[len(df.index)] = get_new_record(info, len(df)) # append new row to dataframe  
        i = i + 1
    return df
    
if __name__ == "__main__":

    # read in excel file containing previously recorded BTC data
    df = pd.read_excel('BTC_Data.xlsx')
    df_new = update_data(df) # adds new row to dataframe containing BTC data
    df_new.to_excel('BTC_Data.xlsx', index = False) ## add new row data to BTC data file
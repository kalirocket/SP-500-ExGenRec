import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
from secrets_ import IEX_CLOUD_API_TOKEN as ICAT


def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def enterPortfolioSize():
    global portfolio_size
    try:
        portfolio_size = input('Enter value of your portfolio: ')
        portfolio_size = float(portfolio_size)
    except ValueError:
        print(f"'{portfolio_size}' is not a valid number!")
        enterPortfolioSize()


stocks = pd.read_csv('sp_500_stocks.csv')
api_url = 'https://cloud.iexapis.com/stable'
columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
df = pd.DataFrame(columns=columns)
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_string = []

for i in range(0, len(symbol_groups)):
    symbols = ','.join(symbol_groups[i])
    batch_data = requests.get(f"{api_url}/stock/market/batch?symbols={symbols}&types=quote&token={ICAT}").json()
    for symbol in symbols.split(','):
        try:
            data = batch_data[symbol]
            new_row = {
                'Ticker': symbol,
                'Stock Price': data['quote']['latestPrice'],
                'Market Capitalization': data['quote']['marketCap'],
                'Number of Shares to Buy': 'N/A'
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        except KeyError:
            print(f"Error occurred with : {symbol}")

portfolio_size = float()
enterPortfolioSize()
position_size = portfolio_size / len(df.index)

for i in range(0, len(df.index)):
    df.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / df.loc[i, 'Stock Price'])

writer = pd.ExcelWriter('Recommended trades.xlsx', engine='xlsxwriter')
df.to_excel(writer, 'Recommended Trades', index=False)

background_color = '#FAAB78'
font_color = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)
dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)
integer_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Price', dollar_format],
    'C': ['Market Capitalization', dollar_format],
    'D': ['Number of Shares to Buy', integer_format]
}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)

writer._save()
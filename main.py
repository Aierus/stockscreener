# Import data from polygon.io or from yahoo finance
from pandas_datareader import data as pdr
from yahoo_fin import stock_info as si
from pandas import ExcelWriter
import yfinance as yf
import pandas as pd
import numpy as np
import datetime
import time
import os

# Variables
spy = si.tickers_sp500()
# naz = si.tickers_nasdaq()
# dow = si.tickers_dow()
# naz_name = '^IXIC'   #NASDAQ
spy_name = '^GSPC' #SPX
# dow_name = '^DJI'    #DOW
# indicies = [spy, naz, dow]
# index_name = [spy_name, naz_name, dow_name]
indicies = [spy]
index_name = [spy_name]
for i in range(len(indicies)):
  indicies[i] = [item.replace(".", "-") for item in indicies[i]] #formatting
start_date = datetime.datetime.now() - datetime.timedelta(days=365)
yester_date = datetime.datetime.now() - datetime.timedelta(days=2)
end_date = datetime.date.today()
exportList = pd.DataFrame(columns=['Stock', 'Volume', 'Average Volume', 'Weighted Volume', 'Percent_Deviation', 'Weighted_Percent_Deviation'])

spy_data = pdr.get_data_yahoo('SPY', start_date, end_date)
spy_data.to_csv('SPY.csv')

spy_yester_data = pdr.get_data_yahoo('SPY', yester_date - datetime.timedelta(days=1), end_date)
spy_yester_data.to_csv('SPY_YESTERDAY.csv')

spy_today_data = pdr.get_data_yahoo('SPY', yester_date, end_date)
spy_today_data.to_csv('SPY_TODAY.csv')

# spooz_list = list(spooz_today_data)
# print(spooz_list)

# spooz_today = pd.read_csv('SPY_TODAY.csv')
# print(spooz_today.columns.values.tolist())
# print(spooz_today['Volume'])

# spooz_data.columns.values.tolist()
# for i in range len(spooz_data.columns):
#   spooz_volume = spooz_today.loc[spooz_today.index, 'Volume'].iat[i]
tickers = []
total_spy_vol = 0
j = 0
# Write data to CSV
for i in range(len(index_name)):
  if i == 0:
    rows = len(spy_data)
    for j in range(rows):
      total_spy_vol += spy_data.loc[spy_data.index, 'Volume'].iat[j]
      average_spy_vol = total_spy_vol / rows

  for ticker in indicies[i]:
    df = pdr.get_data_yahoo(ticker, start_date, end_date)
    rows = len(df.index)
    tickers.append(ticker)

    weighted_vol = []
    for i in range(rows):
      weighted_vol.append(df.loc[df.index, 'Volume'].iat[i])
      weighted_vol[i] = round(weighted_vol[i] / (spy_data.loc[spy_data.index, 'Volume'].iat[i]), 10)
    df['SPY Weighted Volume'] = weighted_vol

    # Create a 365 day average volume
    total_vol = 0
    for i in range(rows):
      total_vol += df.loc[df.index, 'Volume'].iat[i]
    average_vol = total_vol / rows
    
    # Check to see if day's volume is 20% greater than previous day
    volume_deviation = []        # Historical Y/N Volume Deviation > 50%
    volume_deviation_values = [] # Historical Volume Deviation Values
    w_vol_dev_values = []        # Today's Volume Deviations
    for i in range(rows):
      vol_dev = ((df.loc[df.index, 'Volume'].iat[i] - df.loc[df.index, 'Volume'].iat[i-1]) / df.loc[df.index, 'Volume'].iat[i-1]) * 100
      w_vol_dev = ((df.loc[df.index, 'SPY Weighted Volume'].iat[i] - df.loc[df.index, 'SPY Weighted Volume'].iat[i-1]) / df.loc[df.index, 'SPY Weighted Volume'].iat[i-1]) * 100
      vol_dev = round(vol_dev, 10)
      w_vol_dev = round(w_vol_dev, 10)
      volume_deviation_values.append(vol_dev)
      w_vol_dev_values.append(w_vol_dev)
      if ((df.loc[df.index, 'Volume'].iat[i] > 1.2 * df.loc[df.index, 'Volume'].iat[i-1])):
        volume_deviation.append('+')
        # print('Ticker %s on day %d; Had a volume deviation of %.4f from the previous day\n' % (ticker, i, vol_dev))
      else:
        volume_deviation.append('-')
        # volume_deviation_values.append('-')

      if (i == (rows - 1)):
        exportList.at[j, 'Stock'] = ticker
        exportList.at[j, 'Volume'] = df.loc[df.index, 'Volume'].iat[i]
        exportList.at[j, 'Average Volume'] = average_vol
        exportList.at[j, 'Weighted Volume'] = w_vol_dev_values[i]
        exportList.at[j, 'Percent_Deviation'] = vol_dev
        exportList.at[j, 'Weighted_Percent_Deviation'] = w_vol_dev
        j = j + 1
        exportList = exportList.sort_values(by=['Weighted_Percent_Deviation'], ascending=False)
        print(exportList)

    df['Volume Deviation'] = volume_deviation
    df['Volume Deviation Value'] = volume_deviation_values
    df['Weighted Volume Deviation Values'] = w_vol_dev_values

    os.makedirs('csv_output', exist_ok=True)
    df.to_csv(f'csv_output/{ticker}.csv')
    print(f'Writing {ticker} to csv')

# Create DataFrame of top 10%
vd_df = pd.DataFrame(list(zip(tickers, volume_deviation_values)), columns=['Ticker', 'Percent_Deviation'])
vd_df['Volume_Deviation_Rating'] = vd_df.Percent_Deviation.rank(pct=True) * 100
vd_df = vd_df[vd_df.Volume_Deviation_Rating >= vd_df.Volume_Deviation_Rating.quantile(.90)]
vd_df.to_csv('Top_10_Percent_Volume_Deviation_Today.csv')
print("##### Top 10%\ of volume deviation stocks written to csv ######")
""" 
vd_stocks = vd_df['Ticker']
for stock in vd_stocks:
  try:
    df = pd.read_csv(f'{stock}.csv', index_col=0)
    sma = [50,100,200]
    for x in sma:
      df["SMA_"+str(x)] = round(df['Close'].rolling(window=x).mean(), 2)

    currentClose = df["Close"][-1]
    mva_50 = df["SMA_50"][-1]
    mva_100 = df["SMA_100"][-1]
    mva_200 = df["SMA_200"][-1]
    mva_200_20 = df["SMA_200"][-20]
    low_52_wk = round(min(df["Low"][-260:]), 2)
    high_52_wk = round(min(df["High"][-260:]), 2)

    condition_1 = currentClose > mva_100 > mva_200
    condition_2 = mva_100 > mva_200
    condition_3 = mva_200 > mva_200_20
    condition_4 = mva_50 > mva_100 > mva_200
    condition_5 = currentClose > mva_50
    condition_6 = currentClose >= (1.3*low_52_wk)
    condition_7 = currentClose >= (.75*high_52_wk)
    if(condition_1 and condition_2 and condition_3 and condition_4 and condition_5 and condition_6 and condition_7):
      print (stock + " made the Minervini requirements")

  except Exception as e:
    print(e)
    print(f'Could not gather data on {stock}')
"""
# Sort, Trim, and Export to Excel
exportList.sort_values(by=['Percent_Deviation'], ascending=True)
exportList = exportList[exportList.Percent_Deviation >= exportList.Percent_deviation.quantile(.70)]
print('\n', exportList)
writer = ExcelWriter("OverallPercentVolumeDeviation.xlsx")
exportList.to_excel(writer, "Sheet1")
writer.save()
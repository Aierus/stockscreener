import pandas as pd
from pandas import ExcelWriter

exportList = []

vd_df = pd.read_csv('Top_10_Percent_Volume_Deviation_Today.csv', index_col=0)
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
      exportList.append(stock)
      print (stock + " made the Minervini requirements")

  except Exception as e:
    print(e)
    print(f'Could not gather data on {stock}')

print('\n', exportList)
writer = ExcelWriter("Minervi.xlsx")
exportList.to_excel(writer, "Sheet1")
writer.save()
import pandas as pd
import numpy as np
import matplotlib as plot
import datetime
import openpyxl

# Load Data Frame 1 with the Seed Sheet
df_SeedSheet = pd.read_excel("Indian Stocks Dump.xlsx", "Seed Stock Sheet")
# Load Data Frame 2 with data from Web ( Twitter & MoneyControl)
df_WebData = pd.read_excel("Indian Stocks Dump.xlsx", "Web & Social Data")
# Load Data Frame 3 with Nifty Data ( NSE1)
df_Nifty50Stocks = pd.read_excel("Indian Stocks Dump.xlsx", "Nifty50")
# Load Data Frame 4 with additional NSE Stock Parameters
df_AdditionalData = pd.read_excel("Indian Stocks Dump.xlsx", "Additional Stock Metrics")
# Load Data Frame 5 with additional NSE ISIN
df_NseISIN = pd.read_csv("EQUITY_L.csv")

# Merge the Web Data - df_WebData into the main Seed sheet
df_mergedWebData = pd.merge(df_SeedSheet, df_WebData[
    ['Company',
     'Moneycontrol Link',
     'Website',
     'Email',
     'Twitter Link',
     'Twitter Handle',
     'Twitter Followers',
     'Twitter Posts',
     'Twitter Created Date',
     'Twitter Account Age (Years)']],  on = 'Company', how = 'left')



# Map the obsolete ISIN with latest ISIN
key_list = list(df_mergedWebData[~df_mergedWebData['NSE code'].isnull()]['NSE code'])
dict_lookup = dict_lookup = dict(zip(df_NseISIN['SYMBOL'], df_NseISIN['ISIN']))
df_mergedWebData['ISIN'] = df_mergedWebData['NSE code'].map(dict_lookup).fillna(df_mergedWebData['ISIN'])

# Drop NaN ISIN values from the Merged Sheet Data
list_IndexToRemove = df_mergedWebData.index[df_mergedWebData['ISIN'].isna()]
df_mergedWebData.drop(list_IndexToRemove, 0, inplace = True)

# Cast Ticker & ISIN to String column types in Additional Data Sheet
df_AdditionalData = df_AdditionalData.astype({'Ticker':'string', 'ISIN':'string'})

# Drop duplicate Tickers / ISIN from the Additional Sheet 
df_AdditionalData.drop_duplicates(subset = ['ISIN'], keep = 'first', inplace = True )

# Drop NaN ISIN from the Additional Sheet
list_IndexToRemove_NA = df_AdditionalData.index[df_AdditionalData['ISIN'].isna()]
df_AdditionalData.drop(list_IndexToRemove_NA, 0, inplace = True)

# Remove the Funds & ETF data ( starts with INF***** instead of INE*** )from the Additional Data Sheet
list_IndexOfNonStocks = df_AdditionalData.index[~df_AdditionalData['ISIN'].str.startswith('INE')]
df_AdditionalData.drop(list_IndexOfNonStocks, 0, inplace = True)

# Merge the Nifty50 - df_Nifty50Stocks into the merged web Sheet
df_mergedWebData = pd.merge(df_mergedWebData, df_Nifty50Stocks[
    ['ISIN',
     'Nifty 50 Stock']],  on = 'ISIN', how = 'left')

# Merge the Additional metrics data with the merged data frame ( seed and Web)
df_WebAndMoreData = pd.merge(df_mergedWebData, df_AdditionalData[
    ['ISIN',
     'Ticker',
     '5Y Avg ROE',
     '5Y Revenue Growth',
     'Promoter Holding',
     'No. of Shareholders',
     'Pledged Promoter Holdings',
     'Rating agency Buy Reco',
     'Volatility', 
     'Total Debt']], on = 'ISIN', how = 'outer', indicator='true')

# Cleanup Stocks that are not actively traded. These are the stocks that were not found in the Seed Sheet
List_indexOfInactiveStocks = df_WebAndMoreData.index[df_WebAndMoreData['Company'].isnull()]
df_WebAndMoreData.drop(List_indexOfInactiveStocks, 0, inplace = True)

# Remove the Stocks for whom the Earnings & Book Values is not available as it would be difficult to evaluate such companies
List_indexOfStksWithMissingBookVal = df_WebAndMoreData.index[(df_WebAndMoreData['Earning Per Share'].isnull())
                                                            | (df_WebAndMoreData['Book Value Per Share'].isnull())
                                                            | (df_WebAndMoreData['Cash Flow Per Share'].isnull())]

df_WebAndMoreData.drop(List_indexOfStksWithMissingBookVal, 0, inplace = True)

# Fill in the missing values of Price to Earning given that the Price and EPS info is already present
df_WebAndMoreData['Price to Earnings'] = df_WebAndMoreData['Price'] / df_WebAndMoreData['Earning Per Share']

# Round off the Numeric values to a standard 2 decimal digit format
df_WebAndMoreData = df_WebAndMoreData.round(decimals = 2)

# Convert Data Collection date and Twitter Handle creation date to standard datetime
df_WebAndMoreData['Date'] = pd.to_datetime(df_WebAndMoreData['Date'],  format = '%d-%m-%Y')
df_WebAndMoreData['Twitter Created Date'] = pd.to_datetime(df_WebAndMoreData['Twitter Created Date'],  format = '%d-%m-%Y')

# Final Excel and JSON Creation
df_WebAndMoreData.sort_values('Market Cap(Cr)', ascending = True).to_json("StockRefinedData.json", orient = 'index')

df_WebAndMoreData.to_excel("Final Processed Stocks Data.xlsx")
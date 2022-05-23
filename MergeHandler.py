import pandas as pd
import numpy as np
import matplotlib as plot
import datetime
import openpyxl

# Load Data Frame 1 with the Seed Sheet
df_SeedSheet = pd.read_excel("Indian Stocks Dump.xlsx", "Seed Stock Sheet")

# Load Data Frame 2 with data from Web ( Twitter & MoneyControl)
df_WebData = pd.read_excel("Indian Stocks Dump.xlsx", "Web & Social Data")

# Load Data Frame 3 with additional Stock Parameters
df_AdditionalData = pd.read_excel("Indian Stocks Dump.xlsx", "Additional Stock Metrics")

# Cast Ticker & ISIN to String column types
df_AdditionalData = df_AdditionalData.astype({'Ticker':'string', 'ISIN':'string'})

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

     # Capture all the Duplicated data per sheet
df_duplicate_SeednWebData = df_mergedWebData[df_mergedWebData.duplicated('ISIN')]
df_duplicate_AddData_ISIN = df_AdditionalData[df_AdditionalData.duplicated('ISIN')]
df_duplicate_AddData_Ticker = df_AdditionalData[df_AdditionalData.duplicated('Ticker')]
with pd.ExcelWriter('Duplicate Data.xlsx') as writer:  
    df_duplicate_SeednWebData.to_excel(writer, sheet_name='Seed Data')
    df_duplicate_AddData_ISIN.to_excel(writer, sheet_name='Add Data ISIN')
    df_duplicate_AddData_Ticker.to_excel(writer, sheet_name='Add Data Ticker')

    # Drop NaN ISIN values from the Seed Data
list_IndexToRemove = df_mergedWebData.index[df_mergedWebData['ISIN'].isna()]
df_mergedWebData.drop(list_IndexToRemove, 0, inplace = True)

# Drop duplicate Tickers / ISIN from the Additional Sheet 
df_AdditionalData.drop_duplicates(subset = ['ISIN'], keep = 'first', inplace = True )
# Drop NaN ISIN from the Additional Sheet
list_IndexToRemove_NA = df_AdditionalData.index[df_AdditionalData['ISIN'].isna()]
df_AdditionalData.drop(list_IndexToRemove_NA, 0, inplace = True)

# Remove the Funds & ETF data ( starts with INF***** instead of INE*** )from the Additional Data Sheet
list_IndexOfNonStocks = df_AdditionalData.index[~df_AdditionalData['ISIN'].str.startswith('INE')]
df_AdditionalData.drop(list_IndexOfNonStocks, 0, inplace=True)

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

df_WebAndMoreData.to_excel("Merged Social & Additional Data.xlsx")
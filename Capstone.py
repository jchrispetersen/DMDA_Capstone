import pandas as pd
import os
import sys
import itertools
from openpyxl import load_workbook


#read in excel file containing data
wb = load_workbook('Capstone data.xlsx')

ws = wb['Sheet1']

data = ws.values

columns = next(data)[0:]

#create dataframe to hold data
df = pd.DataFrame(data, columns=columns)

#set card numbers to string for later comparisons
df['Card_Number'] = df['Card_Number'].astype(str)

#compare column headers to ensure proper data format
proper_col = ['Card_Number','Merchant_Name','Date','Merchant_ID','Terminal_ID','Is_Fraud']

fraud_col = []

for col in df.columns:
    fraud_col.append(col)

#print to user to show if format matches
if fraud_col == proper_col:
    print("Data Format Matches")
else:
    print("Data format does not match, please reformat data to match:")
    print(proper_col)
    print(col)
    input('Press Enter to Exit')
    sys.exit()

# prompt for fraud card numbers
card_num = input("Enter card numbers, separated by comma plus single space ', ':")

cards = []

cards = card_num.split(", ")

# convert input cards to string
cards = map(str,cards)

# Create dataframe of only transactions that match card inputs
new_df = df[(df['Card_Number'].isin(cards) == True)]

#Find distinct merchants used by cardlist
dist_merch = new_df['Merchant_Name'].unique()

dist_merch.sort()

#remove fraud from dataset
no_fraud = new_df[new_df.Is_Fraud != 1]

#group by merchant/card count
dist_count = no_fraud.groupby(['Merchant_Name','Card_Number'])['Card_Number'].count()

#select merchant that has most unique card visits
dist_count2 = pd.DataFrame(dist_count.groupby(['Merchant_Name']).count().nlargest(1))

comp_merch = dist_count2.index[0]

#output compromised merchant
print('Compromised Merchant is:', comp_merch)
#print(dist_count2)

#pull dates where each card went to compromised merchant
poss_dates = no_fraud[no_fraud.Merchant_Name == comp_merch]

dates = pd.DataFrame(poss_dates.groupby(['Card_Number', 'Date']))

dates2 = pd.DataFrame(dates[0].to_list(),index=dates.index)

dates2.columns = ['Card_Num','Date']

dict_dates = dates2.to_dict()

dist_cards = dates2['Card_Num'].unique().tolist()

#create dictionary of unique cards and dates they went to compromised merchant
card_dict = dates2.groupby('Card_Num').agg({'Date': lambda x: x.tolist()})['Date'].to_dict()

#create dataframe of every possible date combination for the set of cards
date_optimizer = pd.DataFrame(list(itertools.product(*card_dict.values())), columns=card_dict.keys())

#create new column for min date for each iteration
date_optimizer['Min_Date'] = date_optimizer.min(axis = 1)

#create new column for max date for each iteration
date_optimizer['Max_Date'] = date_optimizer.max(axis = 1)

#convert new columns to datetimes
date_optimizer['Min_Date'] = pd.to_datetime(date_optimizer['Min_Date'])
date_optimizer['Max_Date'] = pd.to_datetime(date_optimizer['Max_Date'])

#create new column as difference between min and max date
date_optimizer['Window'] = date_optimizer['Max_Date'] - date_optimizer['Min_Date']

#find smallest window of compromise
low_date = date_optimizer['Window'].min()

low_date = date_optimizer[date_optimizer['Window'] == low_date]

date_range = low_date[['Min_Date','Max_Date']]

comp_min = date_range['Min_Date'][0]

comp_max = date_range['Max_Date'][0]

#output min and max dates of smallest window as compromise window
print('Compromise Window starts:',comp_min)

print('Compromise Window ends:',comp_max)

#modify dataframe based on merchant and compromise window from above
#to leave us with only transactions at merchant during compromise window
card_list = pd.DataFrame(df[(df['Merchant_Name'] == comp_merch)])

card_list2 = pd.DataFrame(card_list[(card_list['Date'] >= comp_min )])

card_list3 = pd.DataFrame(card_list2[(card_list2['Date'] <= comp_max)])

#create list of compromised cards 
comp_cards = card_list3['Card_Number'].unique().tolist()

#output compromised cards
print('Compromised cards are:',comp_cards)

#create input to prevent app from auto closing before user can access data returned
input('Press Enter to Exit')

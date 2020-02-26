import pandas as pd
import os
import numpy as np

path='C:\\Users\\asayeed\\Desktop\\Python Projects\\Pipeline Cognos Report Cleanup\\Run'

dirListing = os.listdir(path)
editFiles = []
for item in dirListing:
    if ".xlsx" in item:
        editFiles.append(item)


editFiles= editFiles.copy()
editFiles.remove('002. NSE Pipeline - Partner Sales Team.xlsx')
editFiles.remove('002. NSE Pipeline - Product Sales Team.xlsx')

print("EditFiles List after removing Partner/Product Sales Team: ",editFiles)

salesFiles=['002. NSE Pipeline - Partner Sales Team.xlsx','002. NSE Pipeline - Product Sales Team.xlsx']


i=1
for item in editFiles:
    df=pd.read_excel(item)

    df.columns = df.columns.str.replace('_', ' ')
    df.replace('null', np.nan, inplace=True)
    df.replace(' ', np.nan, inplace=True)
    df.dropna(inplace=True)
    df.drop_duplicates(subset=None, keep='first', inplace=False)
    writer = pd.ExcelWriter(item, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Page1_1',header=True,index=False)
    writer.save()
    print(str(i) + '/14 completed')
    i = int(i) + 1

i=13
for item in salesFiles:
    df=pd.read_excel(item)

    df.columns = df.columns.str.replace('_', ' ')
    df = df.rename(columns={'Pend Off CYTD': 'Pend_Off_CYTD'})
    df = df.rename(columns={'Pend Off LYTD': 'Pend_Off_LYTD'})
    #df.replace('null', np.nan, inplace=True)
    #df.dropna(inplace=True)
    df.drop_duplicates(subset=None, keep='first', inplace=False)
    writer = pd.ExcelWriter(item, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Page1_1',header=True,index=False)
    writer.save()
    print(str(i) + '/13 completed')
    i = int(i) + 1







#starting thr record of sums

editFiles.remove('002. NSE Pipeline by Centre.xlsx') # remove this file so that it will be checked individually


for item in editFiles:
    df=pd.read_excel(item)
    offers = df['Offers'].sum()
    print(item + "_offers:", offers)
    arrived = df['Arrived'].sum()
    print(item + "_arrived:", arrived)
    confirmed = df['Confirmed'].sum()
    print(item + "_confirmed:", confirmed)
    PO_ASS = df['Place Offered'].sum() + df['Assessment'].sum()
    print(item + "_PO_ASS:", PO_ASS)
    Y_offers = df['YAGO Offers'].sum()
    print(item + "_Y_offers:", Y_offers)
    Y_arrived = df['YAGO Arrived'].sum()
    print(item + "_Y_arrived:", Y_arrived)
    Y_confirmed = df['YAGO Confirmed'].sum()
    print(item + "_Y_confirmed:", Y_confirmed)
    Y_PO_ASS = df['YAGO Place Offered'].sum()+ df['YAGO Assessment'].sum()
    print(item + "_Y_PO_ASS:", Y_PO_ASS)
    print(' ')
    print(' ')



'''
here we check the numbers for the partner & product sales sheet
'''


for item in salesFiles:
    df = pd.read_excel(item)
    per_year = df.groupby('Year')['Arrived'].sum()
    PO_ASS = df.groupby('Year')['Place Offered'].sum()+df.groupby('Year')['Assessment'].sum()

    PO_ASS2020 = PO_ASS[2020]
    #PO_ASS2019 = PO_ASS[2019]
    #PO_ASS2021 = PO_ASS[2021]
    print(item + "_offers_2020:", PO_ASS2020)
    #print(item + "_offers_2019:", PO_ASS2019)
    #print(item + "_offers_2021:", PO_ASS2021)

    arrived2020=per_year[2020]
    #arrived2019 = per_year[2019]
    #arrived2021 = per_year[2021]
    print(item + "_arrived_2020:", arrived2020)
    #print(item + "_arrived_2019:", arrived2019)
    #print(item + "_arrived_2021:", arrived2021)

    print(' ')
    print(' ')

'''
we individually check the NSE Pipeline by Centre sheet
'''
df=pd.read_excel('002. NSE Pipeline by Centre.xlsx')
per_year_offers = df.groupby('Year')['Offers'].sum()
per_year_offers_2020=per_year_offers[2020]
per_year_offers_2019=per_year_offers[2019]
per_year_offers_2021=per_year_offers[2021]
print('002. NSE Pipeline by Centre.xlsx' + "_offers_2020:", per_year_offers_2020)
print('002. NSE Pipeline by Centre.xlsx' + "_offers_2019:", per_year_offers_2019)
print('002. NSE Pipeline by Centre.xlsx' + "_offers_2021:", per_year_offers_2021)

per_year_arrived = df.groupby('Year')['Arrived'].sum()
per_year_arrived_2020=per_year_arrived[2020]
per_year_arrived_2019=per_year_arrived[2019]
per_year_arrived_2021=per_year_arrived[2021]
print('002. NSE Pipeline by Centre.xlsx' + "_arrived_2020:", per_year_arrived_2020)
print('002. NSE Pipeline by Centre.xlsx' + "_arrived_2019:", per_year_arrived_2019)
print('002. NSE Pipeline by Centre.xlsx' + "_arrived_2021:", per_year_arrived_2021)

per_year_confirmed = df.groupby('Year')['Confirmed'].sum()
per_year_confirmed_2020=per_year_confirmed[2020]
per_year_confirmed_2019=per_year_confirmed[2019]
per_year_confirmed_2021=per_year_confirmed[2021]
print('002. NSE Pipeline by Centre.xlsx' + "_confirmed_2020:", per_year_confirmed_2020)
print('002. NSE Pipeline by Centre.xlsx' + "_confirmed_2019:", per_year_confirmed_2019)
print('002. NSE Pipeline by Centre.xlsx' + "_confirmed_2021:", per_year_confirmed_2021)

per_year_PO_ASS=  df.groupby('Year')['Place Offered'].sum()+ df.groupby('Year')['Assessment'].sum()
per_year_PO_ASS_2020=per_year_PO_ASS[2020]
per_year_PO_ASS_2019=per_year_PO_ASS[2019]
per_year_PO_ASS_2021=per_year_PO_ASS[2021]
print('002. NSE Pipeline by Centre.xlsx' + "_PO_ASS_2020:", per_year_PO_ASS_2020)
print('002. NSE Pipeline by Centre.xlsx' + "_PO_ASS_2019:", per_year_PO_ASS_2019)
print('002. NSE Pipeline by Centre.xlsx' + "_PO_ASS_2021:", per_year_PO_ASS_2021)

per_year_y_offers = df.groupby('Year')['YAGO Offers'].sum()
per_year_y_offers_2020=per_year_y_offers[2020]
per_year_y_offers_2019=per_year_y_offers[2019]
per_year_y_offers_2021=per_year_y_offers[2021]
print('002. NSE Pipeline by Centre.xlsx' + "__y_offers_2020:", per_year_y_offers_2020)
print('002. NSE Pipeline by Centre.xlsx' + "__y_offers_2019:", per_year_y_offers_2019)
print('002. NSE Pipeline by Centre.xlsx' + "__y_offers_2021:", per_year_y_offers_2021)

per_year_y_arrived = df.groupby('Year')['YAGO Arrived'].sum()
per_year_y_arrived_2020=per_year_y_arrived[2020]
per_year_y_arrived_2019=per_year_y_arrived[2019]
per_year_y_arrived_2021=per_year_y_arrived[2021]
print('002. NSE Pipeline by Centre.xlsx' + "_y_arrived_2020:", per_year_y_arrived_2020)
print('002. NSE Pipeline by Centre.xlsx' + "_y_arrived_2019:", per_year_y_arrived_2019)
print('002. NSE Pipeline by Centre.xlsx' + "_y_arrived_2021:", per_year_y_arrived_2021)


per_year_y_confirmed = df.groupby('Year')['YAGO Confirmed'].sum()
per_year_y_confirmed_2020=per_year_y_confirmed[2020]
per_year_y_confirmed_2019=per_year_y_confirmed[2019]
per_year_y_confirmed_2021=per_year_y_confirmed[2021]
print('002. NSE Pipeline by Centre.xlsx' + "_offers_2020:", per_year_offers_2020)
print('002. NSE Pipeline by Centre.xlsx' + "_offers_2019:", per_year_offers_2019)
print('002. NSE Pipeline by Centre.xlsx' + "_offers_2021:", per_year_offers_2021)

per_year_y_PO_ASS=  df.groupby('Year')['YAGO Place Offered'].sum()+ df.groupby('Year')['YAGO Assessment'].sum()
per_year_y_PO_ASS_2020=per_year_y_PO_ASS[2020]
per_year_y_PO_ASS_2019=per_year_y_PO_ASS[2019]
per_year_y_PO_ASS_2021=per_year_y_PO_ASS[2021]
print('002. NSE Pipeline by Centre.xlsx' + "_y_PO_ASS_2020:", per_year_y_PO_ASS_2020)
print('002. NSE Pipeline by Centre.xlsx' + "_y_PO_ASS_2019:", per_year_y_PO_ASS_2019)
print('002. NSE Pipeline by Centre.xlsx' + "_y_PO_ASS_2021:", per_year_y_PO_ASS_2021)

print()
print(' ')
print(' ')



print('all completed')



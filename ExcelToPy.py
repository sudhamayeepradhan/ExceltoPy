#!/usr/bin/env python

"""ExcelToPy.py: Migrate an excel to a simplified one."""

__author__      = "Sudhamayee Pradhan"

import pandas as pd
import numpy as np
pd.options.mode.chained_assignment = None  # default='warn'

Data_1=pd.read_excel('C:/Users/dv914285/Desktop/Camion Switzerland/Camion Switzerland Invoice Validation 2021 V2.xlsx',sheet_name = 'Data')
Ratecard=pd.read_excel('C:/Users/dv914285/Desktop/Camion Switzerland/Camion Switzerland Invoice Validation 2021 V2.xlsx',sheet_name = 'RateCard & Zone')

Data_1.rename(columns={'Destination':'PostCode'},inplace=True)
df3=pd.merge(Data_1,Ratecard[['PostCode','Zones']],on='PostCode',how='left')

df3['Consolidate Wt'] = (df3.apply(lambda r: df3.loc[(df3['Type of transport'] == 'Standard') & (df3['Shipment Date'] == r['Shipment Date']) & (df3['Customer'] == r['Customer'])  & (df3['Sender ZIP'] == r['Sender ZIP']) & (df3['Type of transport'] == r['Type of transport'])& (df3['PostCode'] == r['PostCode']), 'Chargeable weight'].sum(), axis = 1))

val = (df3.apply(lambda r: df3.loc[(df3['Type of transport'] == 'Standard') & (df3['Shipment Date'] == r['Shipment Date']) & (df3['Customer'] == r['Customer'])  & (df3['Sender ZIP'] == r['Sender ZIP']) & (df3['Type of transport'] == r['Type of transport'])& (df3['PostCode'] == r['PostCode']), 'Chargeable weight'].count(), axis = 1))
val = [(1/elem) if elem !=0 else 0 for elem in val]

df3["% Shipment"] = val
rate = Ratecard.iloc[0:10,4:13] 


new_val = []
df3["New Base Rate"] = None
for i in range(len(df3)):
    for j in range(len(rate)):
        if df3.loc[i]['Type of transport'] == 'Standard' and df3.loc[i]['Sender ZIP'] == 9542:
            if df3.loc[i]['Consolidate Wt'] > 7000:
                if df3.loc[i]['Zones'] == rate.loc[j][rate.columns[0]]:
                    for col in range(2, len(rate.loc[j])):
                        if df3.loc[i]['Consolidate Wt'] < rate.loc[j].keys()[col]:
                            new_val.append((rate.loc[j][rate.loc[j].keys()[col-1]]) * df3.loc[i]["% Shipment"])
                            df3["New Base Rate"][i] = (rate.loc[j][rate.loc[j].keys()[col-1]]) * df3.loc[i]["% Shipment"]
                            break
                    break
            else:
                df3["New Base Rate"][i] = (((df3.loc[i]['Consolidate Wt'] /100) * 8.5 ) + 44) * df3.loc[i]["% Shipment"]
                break
        else:
            df3["New Base Rate"][i] = 0
            break

mask1 = ((df3['Type of transport'] == 'Standard') & (df3['Sender ZIP'] == 9542))
df3.loc[mask1, 'Diff'] = df3['Final BaseRate'] - df3["New Base Rate"]
df3.loc[~mask1, 'Diff'] = 0 

new_columns = df3[['Shipping Point', 'Month', 'Customer', 'PostCode', 'Type of transport', 'Consolidate Wt', '% Shipment',
                    'Sender ZIP', 'New Base Rate', 'Diff']]

# create excel writer object
writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')
# write dataframe to excel
new_columns.to_excel(writer, sheet_name='Data', startrow=1, header=False, index= False)

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Data']

# column_settings = [{'header': column} for column in new_columns.columns]
(max_row, max_col) = new_columns.shape

# Add a header format.
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#6eeb2a',
    'border': 1})

# Write the column headers with the defined format.
for col_num, value in enumerate(new_columns.columns.values):
    worksheet.write(0, col_num, value, header_format)

worksheet.set_column(0, 0, 16)
worksheet.set_column(2, 2, 40)
worksheet.set_column(3, 3, 11)
worksheet.set_column(4, 4, 18)
worksheet.set_column(5, 5, 17)
worksheet.set_column(6, 6, 14)
worksheet.set_column(8, 8, 16)

worksheet.autofilter(0, 0, max_row, max_col - 1)

df = pd.pivot_table(new_columns,index=["Month"],values=["Diff","New Base Rate"],aggfunc=np.sum, fill_value=0)
df.to_excel(writer,sheet_name='Pivot')

# save the excel
writer.save()
print('DataFrame is written successfully to Excel File.')

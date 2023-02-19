import xlrd, xlwt
import os
import pandas as pd
import csv
# rb = xlrd.open_workbook('../ArticleScripts/ExcelPython/xl.xls',formatting_info=True)


file_prod = 'traff_ga.xlsx'
file_adv = 'Adv2.xlsx'
file_rek = 'rek.xlsx'
prod_xl = pd.ExcelFile(file_prod)
adv_xl = pd.ExcelFile(file_adv)
rek_xl = pd.ExcelFile(file_rek)
#print(adv_xl.sheet_names)
prod = prod_xl.parse('traff_ga')
# print(prod)
budj = adv_xl.parse('РК Бюджет')
chan = adv_xl.parse('Каналы')
rek = rek_xl.parse('Каналы')
# print(rek)
#print(budj)
# prod_xl.add
New_Data = dict()

# print(budj.shape[0])
for ch in budj['Канал трафика']:
    New_Data[ch] = 0
#print(New_Data)
df = pd.DataFrame(New_Data, index=[0])
#print(df)
channels = []
for index, row in prod.iterrows():
    utm_source = row['source']
    utm_medium = row['medium']
    # print(utm_source)
    # if index> 15:
    #     break
    for index_rek, row_rek in rek.iterrows():
        if row_rek['Оператор'] == 'и':
            if row_rek['UTM_Source'] == utm_source and row_rek['UTM Medium'] == utm_medium:
                channels.append(row_rek['Канал трафика'])
                break

        if row_rek['Оператор'] == 'или':
            # print('!!!')
            if row_rek['UTM_Source'] == utm_source or row_rek['UTM Medium'] == utm_medium:
                channels.append(row_rek['Канал трафика'])
                break
        if row_rek['Оператор'] == '-':
            if row_rek['UTM_Source'] == utm_source or row_rek['UTM Medium'] == utm_medium:
                channels.append(row_rek['Канал трафика'])
                break
    if len(channels)< index+1:
        channels.append('-')
    #print(channels[index], row)

#print(channels)
prod['Канал трафика'] = channels
print(prod)
prod.to_excel('./traf_new.xlsx', sheet_name='1', index=False)

    # utm_medium =
    #print(type(row))
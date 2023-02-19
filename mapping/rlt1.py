import xlrd, xlwt
import os
import pandas as pd
# rb = xlrd.open_workbook('../ArticleScripts/ExcelPython/xl.xls',formatting_info=True)


file_prod = 'Prod.xlsx'
file_adv = 'Adv.xlsx'
file_rek = 'rek.xlsx'
file_prod_new = 'prod_new.xlsx'
prod_xl = pd.ExcelFile(file_prod)
prod_new_xl = pd.ExcelFile(file_prod_new)
adv_xl = pd.ExcelFile(file_adv)
rek_xl = pd.ExcelFile(file_rek)

prod = prod_xl.parse('Продажи')
prod_new = prod_new_xl.parse('1')
# print(prod)
budj = adv_xl.parse('РК Бюджет')
chan = adv_xl.parse('Каналы')
rek = rek_xl.parse('Каналы')

New_Data = dict()

# print(budj.shape[0])
data = {}
for ch in chan['Канал трафика']:
    New_Data[ch] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
New_Data['Месяц'] = ['янв', 'фев', 'март', 'апр', 'май', 'июнь', 'июль', 'авг', 'сен', 'окт', 'ноя', 'дек', 'Год']

#print(New_Data)
df = pd.DataFrame(New_Data, index=[0,1,2,3,4,5,6,7,8,9,10,11,12])
#df['Месяц'] = ['янв', 'фев', 'март', 'апр', 'май', 'июнь', 'июль', 'авг', 'сен', 'окт', 'ноя', 'дек', 'Год']
#print(df)
for index, row in prod_new.iterrows():
    month = int(row['Время'][3:5])

    #df[row['Канал трафика']] += row['Стоимость услуги'] - row['Сумма НДС']
    df[row['Канал трафика']][month-1] += float(row['Стоимость услуги'])
    df[row['Канал трафика']][12] += float(row['Стоимость услуги'])
    # data[row['Канал трафика']] -= float(row['Сумма НДС'])
    #print(data[row['Канал трафика']])

print(df)
df.to_excel('./df.xlsx', sheet_name='1', index=False)
# print(df)

import xlrd, xlwt
import os
import pandas as pd
# rb = xlrd.open_workbook('../ArticleScripts/ExcelPython/xl.xls',formatting_info=True)


file_prod = 'Prod.xlsx'
file_adv = 'Adv.xlsx'
file_rek = 'rek.xlsx'
file_prod_new = 'prod_new.xlsx'
file_traff = 'traff.xlsx'
file_traff_new = 'traf_new.xlsx'
file_df = 'df.xlsx'
prod_xl = pd.ExcelFile(file_prod)
prod_new_xl = pd.ExcelFile(file_prod_new)
adv_xl = pd.ExcelFile(file_adv)
rek_xl = pd.ExcelFile(file_rek)
traff_xl = pd.ExcelFile(file_traff)
traff_new_xl = pd.ExcelFile(file_traff_new)
df_xl = pd.ExcelFile(file_df)

prod = prod_xl.parse('Продажи')
prod_new = prod_new_xl.parse('1')
# print(prod)
budj = adv_xl.parse('РК Бюджет')
chan = adv_xl.parse('Каналы')
rek = rek_xl.parse('Каналы')
traff = traff_xl.parse('traff_ga')
traff_new = traff_new_xl.parse('1')
df = df_xl.parse('1')

New_Data = dict()

# print(budj.shape[0])
data = {}
for ch in chan['Канал трафика']:
    New_Data[ch] = [0, 0, 0]
# New_Data['Месяц'] = ['янв', 'фев', 'март', 'апр', 'май', 'июнь', 'июль', 'авг', 'сен', 'окт', 'ноя', 'дек', 'Год']

#print(New_Data)
df1 = pd.DataFrame(New_Data, index=[0, 1, 2])
df1['данные'] = ['сумма стоим', 'колво сессий', 'стоимость 1 сессия']
# print(df1)
for index, row in traff_new.iterrows():
    # month = int(row['Время'][3:5])

    # df[row['Канал трафика']] += row['Стоимость услуги'] - row['Сумма НДС']
    df1[row['Канал трафика']][0] = df[row['Канал трафика']][12]
    df1[row['Канал трафика']][1] += int(row['sessions'])
    if df1[row['Канал трафика']][0] != 0:
        df1[row['Канал трафика']][2] = df1[row['Канал трафика']][1] / int(df1[row['Канал трафика']][0])


    # df[row['Канал трафика']][12] += float(row['Стоимость услуги'])
    # data[row['Канал трафика']] -= float(row['Сумма НДС'])
    #print(data[row['Канал трафика']])

print(df1)
df1.to_excel('./df1.xlsx', sheet_name='1', index=False)
# print(df)

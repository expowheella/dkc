import pandas as pd

df = pd.read_excel('C:\\Users\Petr\Downloads\Otchet_po_neudovl_sprosu_24.03.2021__1100.xlsx', sheet_name='Sheet1')
new_header = df.iloc[0] #grab the first row for the header
df = df[1:] #take the data except the header row
df.columns = new_header
df = df[['Первая подтвержденная дата доставки на день размещения заказа', 'Код материала', 'Номер и название проекта',
         'Сумма  кол-ва по заказу, руб с НДС']]

# df['Сумма  кол-ва по заказу, руб с НДС'] = [x.replace('.', ',') for x in df['Сумма  кол-ва по заказу, руб с НДС']]
# df['Сумма  кол-ва по заказу, руб с НДС'] = [x.replace(' ', '') for x in df['Сумма  кол-ва по заказу, руб с НДС']]
# df[['Сумма  кол-ва по заказу, руб с НДС']] = df['Сумма  кол-ва по заказу, руб с НДС'].astype(float)

# searching for the part of a word in column
df_pta = df[df['Код материала'].str.contains(r'PTA(?!$)')]
df_ptc = df[df['Код материала'].str.contains(r'PTC(?!$)')]
df_ptn = df[df['Код материала'].str.contains(r'PTN(?!$)')]

df_dta = df[df['Код материала'].str.contains(r'DTA(?!$)')]
df_dtc = df[df['Код материала'].str.contains(r'DTC(?!$)')]
df_dtn = df[df['Код материала'].str.contains(r'DTN(?!$)')]

df_ltc = df[df['Код материала'].str.contains(r'LTC(?!$)')]
df_ltn = df[df['Код материала'].str.contains(r'LTN(?!$)')]
# df = df.sort_values('Номер и название проекта')

# the sum of money in orders by project name
df_pta = df_pta.groupby(['Номер и название проекта'])['Сумма  кол-ва по заказу, руб с НДС'].sum()
df_ptc = df_ptc.groupby(['Номер и название проекта'])['Сумма  кол-ва по заказу, руб с НДС'].sum()
df_ptn = df_ptn.groupby(['Номер и название проекта'])['Сумма  кол-ва по заказу, руб с НДС'].sum()

df_dta = df_dta.groupby(['Номер и название проекта'])['Сумма  кол-ва по заказу, руб с НДС'].sum()
df_dtc = df_dtc.groupby(['Номер и название проекта'])['Сумма  кол-ва по заказу, руб с НДС'].sum()
df_dtn = df_dtn.groupby(['Номер и название проекта'])['Сумма  кол-ва по заказу, руб с НДС'].sum()

df_ltc = df_ltc.groupby(['Номер и название проекта'])['Сумма  кол-ва по заказу, руб с НДС'].sum()
df_ltn = df_ltn.groupby(['Номер и название проекта'])['Сумма  кол-ва по заказу, руб с НДС'].sum()

df_merged = pd.concat([df_pta, df_ptc, df_ptn, df_dta, df_dtc, df_dtn, df_ltc, df_ltn])
print(df_merged)

df_merged.to_excel('data.xlsx')


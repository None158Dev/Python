import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re


path = '1.xlsx'

sh_name = pd.ExcelFile(path).sheet_names

listJob = ['ФЕР', 'ТЕР', 'ТССЦпг', 'ФССЦпг', 'ФСЭМ', 'PRDX']
filter_sum = ['ЗП', 'ЭМ', 'НР от ФОТ', 'СП от ФОТ']
filter_df = ['всего по позиции', 'ЗП', 'ЭМ', 'МР', 'нр от фот', 'сп от фот']
PRDX = [0, 'PRDX', 'Вспомогательная расценка', 'шт', 1, 1]

def KeyGen():
    daysDelta = datetime.now() - datetime(1899, 12, 30)
    today = timedelta(hours=datetime.now().hour, 
                        minutes=datetime.now().minute, 
                        seconds=datetime.now().second)

    maxhours = timedelta(hours=24, 
                        minutes=0, 
                        seconds=0)

    return int(round((daysDelta.days + today / maxhours)*10000000000, 0))

def Table_sm(df_temp, i):
    temp_sp = pd.DataFrame()
    
# Верхняя граница таблицы
    tb_lbound = list(df_temp.index[df_temp.iloc[:,0].str.contains('раздел', na=False, flags=re.IGNORECASE)].tolist())
    
# Нижняя граница таблицы
    tb_col2 = df_temp.index[df_temp.iloc[:,1].str.contains('', na=False, flags=re.IGNORECASE)].tolist()
    tb_col3 = df_temp.index[df_temp.iloc[:,2].str.contains('всего по позиции', na=False, flags=re.IGNORECASE)].tolist()

# Выбираем последнюю строку для ubound
    if len(tb_lbound) == 0:
        print('Верхняя граница в смете № ' + i + ' не найдена! Нужно в начало сметы добавить "Раздел 1."'); quit()
    elif len(tb_col3) == 0:
        print('Нижняя граница в смете № ' + i + ' не найдена! Нужно в конец сметы добавить "Всего по позиции"'); quit()
    else:
        if tb_col2[len(tb_col2)-3] < tb_col3[len(tb_col3)-1]:
            tb_ubound = [tb_col3[len(tb_col3)-1]]
        else:
            tb_ubound = [tb_col2[len(tb_col2)-3]+1]

        df_temp.rename(columns = {0:'ПП', 1:'Код ресурса по смете', 2:'Наименование позиции', 3:'Ед. Изм.', 4:'Кол-во', 9:'Заказчик сумма'}, inplace = True)
# сплит раздела
        df_temp['Наименование раздела'] = np.where(df_temp['ПП'].str.contains('раздел', na=False, flags=re.IGNORECASE), df_temp['ПП'], np.nan)
        temp_sp[['Наименование раздела', 'Ключ раздела']] =  df_temp['Наименование раздела'].str.split('.', n=1, expand=True, regex=False)
        temp_sp['Наименование раздела'] =  temp_sp['Наименование раздела'].str.findall(r'\d+').explode().str.strip()
        temp_sp['Ключ раздела'] =  temp_sp['Ключ раздела'].str.replace(r'\s+', ' ', regex=True).str.strip()
        df_temp[['Ключ раздела', 'Наименование раздела']] = temp_sp
# остальное
        df_temp['Наименование группы'] = np.where((df_temp['Наименование позиции'].isna() == True) & (df_temp['ПП'].str.contains('раздел', na=False, flags=re.IGNORECASE) == False), df_temp['ПП'].str.strip(), np.nan)
        df_temp[['Наименование раздела', 'Ключ раздела', 'Наименование группы', 'ПП']] = df_temp[['Наименование раздела', 'Ключ раздела', 'Наименование группы', 'ПП']].fillna(method="ffill").astype({'Наименование раздела': 'str', 'Ключ раздела': 'str', 'Наименование группы': 'str'})
        df_temp['Ключ cметы'] = i.replace('-', '_').replace( 'ЛН', '').strip()
        return df_temp.iloc[tb_lbound[0]:tb_ubound[0], [0, 1, 2, 3, 4, 9, 10, 11, 12, 13]].reset_index(drop=True)

def check_float(row):
    try:
        float(row)
        return row
    except ValueError:
        return float(''.join(filter(str.isdigit, row)))

def count_entry(row):
    if row.index[0] == 'Ключ раздела':
        sign = '-'
    else:
        sign = '='

    if len(str(row[0])) == 1:
        return row[1] + sign + '00' + str(row[0])
    elif  len(str(row[0])) == 2:
        return row[1] + sign + '0' + str(row[0])
    else:
        return row[1] + sign + str(row[0])

def sp_Unit(Unit):
# Проверка на пустую ячеку в Ед. Изм.
    if Unit[0] != Unit[0]:
        Unit[0] = '1 шт'
        temp_sp = Unit[0].split(' ', 1)
    else:
        temp_sp = Unit[0].split(' ', 1)
# Делим строки
    if len(temp_sp) == 1:
        temp_sp = [1 * Unit[1] , temp_sp[0]]
    else:
        if temp_sp[0].isnumeric() == False:
            temp_sp[0] = 1; temp_sp[1] = Unit[0]
        temp_sp = [float(temp_sp[0]) * Unit[1], temp_sp[1]]
    return temp_sp

df = pd.DataFrame()
for i in sh_name:
    df_temp = pd.read_excel(path, sheet_name=i, header=None, usecols=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9])
    df = pd.concat([df, Table_sm(df_temp, i)], ignore_index=True)

df.loc[(df.iloc[:,1].shift(-1).str.contains('|'.join(listJob), na=False, flags=re.IGNORECASE) != True & (df.iloc[:,1].shift(-1).isna() != True)) & df.iloc[:,2].isna(), ['ПП', 'Код ресурса по смете', 'Наименование позиции', 'Ед. Изм.', 'Кол-во', 'Заказчик сумма']] = PRDX
df.reset_index(inplace=True)

df['Job'] = df['Код ресурса по смете'].str.contains('|'.join(listJob), na=False, flags=re.IGNORECASE)
df['ind_ForGR'] = np.where(df['Код ресурса по смете'].str.contains('|'.join(listJob), na=False, flags=re.IGNORECASE), df['index'], np.nan)
df['ForGr'] = df['Наименование позиции'].isin(filter_sum) | df['Код ресурса по смете'].str.contains('|'.join(listJob), na=False, flags=re.IGNORECASE)

df_Key = df.loc[df['Job'] == 1].reset_index(drop=True)
df[['ind_ForGR']] = df[['ind_ForGR']].fillna(method="ffill").astype({'ind_ForGR': 'Int64'})
df['Заказчик сумма'] = df['Заказчик сумма'].apply(check_float)

# Сумма в работу
dfsumm = df.groupby(['ind_ForGR', 'ForGr'])['Заказчик сумма'].sum().reset_index(drop=False)
dfsumm = dfsumm.loc[dfsumm.iloc[:,1] == 1]
dfsumm.rename(columns = {'ind_ForGR':'index', 'Заказчик сумма':'Temp_Summ'}, inplace = True)

df_Key['Ключ расценки'] = 'Р_' + (KeyGen() + df_Key.index).astype('str')

df = df.merge(df_Key[['Ключ расценки', 'index']], how='left', on='index')
df = df.merge(dfsumm[['index', 'Temp_Summ']], how='left', on='index')

df = df.loc[((df.iloc[:,2].isna() == False) & (df.iloc[:,2].str.contains('|'.join(filter_df), na=False, flags=re.IGNORECASE) == False))].reset_index(drop=True)

df['Заказчик сумма'] = np.where(df['Job'] == 1, df['Temp_Summ'], df['Заказчик сумма'])
df[['Ключ расценки']] = df[['Ключ расценки']].fillna(method="ffill")
df['Тип р/м/о'] = np.where(df['Job'] == 1,'р', 'м')

df['Ключ раздела'] = df[['Ключ раздела', 'Ключ cметы']].apply(count_entry, axis=1)
df = df.drop(['ind_ForGR', 'ForGr', 'Temp_Summ', 'Job'], axis=1)


df['Кол-во'], df['Ед. Изм.'] = zip(*df[['Ед. Изм.', 'Кол-во']].apply(sp_Unit, axis=1))

Key_GR = pd.pivot_table(df, index=['Наименование раздела', 'Наименование группы'], values='index', aggfunc={'index':min}, sort=False).reset_index(drop=False)
Key_GR['Ключ группы'] = Key_GR.groupby(['Наименование раздела']).cumcount()+1
df = df.merge(Key_GR[['Ключ группы', 'index']], how='left', on='index')

df['Ключ группы'] = df['Ключ группы'].fillna(method="ffill").astype({'Ключ группы': 'Int64'})

df['Ключ группы'] = df[['Ключ группы', 'Ключ раздела']].apply(count_entry, axis=1)
df[['Ключ договора', 'Ключ бюджета', 'Ключ позиции (порядковый)', 'Ключ позиции', 'Ключ ресурса(порядковый)', 'Примечание']] = 'С_ТС', '—', '', '', '—', ''

df = df[['Ключ договора', 'Ключ бюджета', 'Ключ cметы', 'Ключ раздела', 'Ключ группы', 'Ключ позиции (порядковый)', 
         'Наименование раздела', 'Наименование группы', 'Ключ расценки', 'Ключ позиции', 'Ключ ресурса(порядковый)', 
         'Примечание', 'Тип р/м/о', 'Код ресурса по смете', 'Наименование позиции', 'Ед. Изм.', 'Кол-во', 'Заказчик сумма']]

df.to_excel('Data.xlsx')

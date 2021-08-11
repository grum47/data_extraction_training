import pandas as pd
import numpy as np

"""Выборка значений за текущий месяц этого года"""

df_initial = pd.read_excel('C:/Users/User/Downloads/06/1029.xlsx', sheet_name='Н-П_23')
df_initial = df_initial.rename(columns={'Произведено важнейших видов продукции за июнь 2021 года': 'production',
                                        'Unnamed: 1': 'code',
                                        'Unnamed: 2': 'accounting_month',
                                        'Unnamed: 3': 'previous_month',
                                        'Unnamed: 4': 'month_of_the_previous_year'})
df = df_initial[['production',
                 'code',
                 'accounting_month',
                 'previous_month',
                 'month_of_the_previous_year']]
df.head(10)
df = df.iloc[7:]
df['index'] = pd.isna(df['accounting_month'])
#  df.index[df['index'] != True]
df['country'] = 0
df['country'] = df.loc[df['index'] is False]
df['index_code'] = df['code'].str.find('.')
df['new_code'] = np.where(df.index_code == 2, df.code, np.NaN)
df['new_code'].fillna(method='ffill', inplace=True)
df_transition = df[['country', 'accounting_month', 'new_code']]
df_transition['Код ОКПД'] = ''
df_transition['Параметр'] = 'Производство'
df_transition['Месяц'] = 6
df_transition['Год'] = 2021
df_transition['1, 2, 3'] = 1
df_transition = df_transition.rename(columns={'country': 'Субъект',
                                              'accounting_month': 'Значение',
                                              'new_code': 'Код ОКПД-2'})
df_accounting_month = pd.DataFrame(df_transition, columns=['Субъект',
                                                           'Код ОКПД',
                                                           'Параметр',
                                                           'Значение',
                                                           'Месяц',
                                                           'Год',
                                                           '1, 2, 3',
                                                           'Код ОКПД-2'])
df_accounting_month.dropna()
writer = pd.ExcelWriter('D:/Work/python/Преобразования через pandas/по работе/2021/06/Текущий месяц.xlsx')
df_accounting_month.to_excel(writer)
writer.save()
print('Текущий месяц загружен в файл ===C-Пользователи-User===')

# """Выборка значений за предыдущий месяц этого года"""
#
# df_initial = pd.read_excel('C/Users/User/Downloads/06/1029.xlsx', sheet_name='Н-П_23')
# df_initial.head()
# df_initial = df_initial.rename(columns={'Произведено важнейших видов продукции за июнь 2021 года': 'production',
#                                         'Unnamed: 1': 'code',
#                                         'Unnamed: 2': 'accounting_month',
#                                         'Unnamed: 3': 'previous_month',
#                                         'Unnamed: 4': 'month_of_the_previous_year'})
# df = df_initial[['production',
#                  'code',
#                  'accounting_month',
#                  'previous_month',
#                  'month_of_the_previous_year']]
# df.head(10)
# df = df.iloc[7:]
# df['index'] = pd.isna(df['accounting_month'])
# #  df.index[df['index'] != True]
# df['country'] = 0
# df['country'] = df.loc[df['index'] is False]
# df['index_code'] = df['code'].str.find('.')
# df['new_code'] = np.where(df.index_code == 2, df.code, np.NaN)
# df['new_code'].fillna(method='ffill', inplace=True)
# df_transition = df[['country', 'previous_month', 'new_code']]
# df_transition['Код ОКПД'] = ''
# df_transition['Параметр'] = 'Производство'
# df_transition['Месяц'] = 5
# df_transition['Год'] = 2021
# df_transition['1, 2, 3'] = 2
# df_transition = df_transition.rename(columns={'country': 'Субъект',
#                                               'previous_month': 'Значение',
#                                               'new_code': 'Код ОКПД-2'})
# df_accounting_month = pd.DataFrame(df_transition, columns=['Субъект',
#                                                            'Код ОКПД',
#                                                            'Параметр',
#                                                            'Значение',
#                                                            'Месяц',
#                                                            'Год',
#                                                            '1, 2, 3',
#                                                            'Код ОКПД-2'])
# df_accounting_month.dropna()
# writer = pd.ExcelWriter('Предыдущий месяц.xlsx')
# df_accounting_month.to_excel(writer)
# writer.save()
# print('Предыдущий месяц загружен в файл ===C-Пользователи-User===')
#
# """Выборка значений за аналогичный текущий месяц прошлого года"""
#
# df_initial = pd.read_excel('C/Users/User/Downloads/06/1029.xlsx', sheet_name='Н-П_23')
# df_initial = df_initial.rename(columns={'Произведено важнейших видов продукции за июнь 2021 года': 'production',
#                                         'Unnamed: 1': 'code',
#                                         'Unnamed: 2': 'accounting_month',
#                                         'Unnamed: 3': 'previous_month',
#                                         'Unnamed: 4': 'month_of_the_previous_year'})
# df = df_initial[['production',
#                  'code',
#                  'accounting_month',
#                  'previous_month',
#                  'month_of_the_previous_year']]
# df.head(10)
# df = df.iloc[7:]
# df['index'] = pd.isna(df['accounting_month'])
# #  df.index[df['index'] != True]
# df['country'] = 0
# df['country'] = df.loc[df['index'] is False]
# df['index_code'] = df['code'].str.find('.')
# df['new_code'] = np.where(df.index_code == 2, df.code, np.NaN)
# df['new_code'].fillna(method='ffill', inplace=True)
# df_transition = df[['country', 'month_of_the_previous_year', 'new_code']]
# df_transition['Код ОКПД'] = ''
# df_transition['Параметр'] = 'Производство'
# df_transition['Месяц'] = 6
# df_transition['Год'] = 2020
# df_transition['1, 2, 3'] = 3
# df_transition = df_transition.rename(columns={'country': 'Субъект',
#                                               'month_of_the_previous_year': 'Значение',
#                                               'new_code': 'Код ОКПД-2'})
# df_accounting_month = pd.DataFrame(df_transition, columns=['Субъект',
#                                                            'Код ОКПД',
#                                                            'Параметр',
#                                                            'Значение',
#                                                            'Месяц',
#                                                            'Год',
#                                                            '1, 2, 3',
#                                                            'Код ОКПД-2'])
# df_accounting_month.dropna()
# writer = pd.ExcelWriter('Месяц прошлого года.xlsx')
# df_accounting_month.to_excel(writer)
# writer.save()
# print('Месяц прошлого года загружен в файл ===C-Пользователи-User===')

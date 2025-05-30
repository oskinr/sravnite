#module.py
import pandas as pd
from tkinter import messagebox
def some_function(df1,df2,st1):

    # df1 = pd.read_excel('file1.xls')
    # df2 = pd.read_excel('file2.xls')
# Объединяем в одну таблицу и вычисляем различия
    
    # print('Файл1', df1)
    # print('Файл2', df2)
    
# Имя основного столбца для подсчета частот
    target_column = st1
    
# Оставшиеся столбцы кроме целевой колонки
    other_columns = [col for col in df1.columns if col != target_column]

# Подсчет частот в каждом файле
    freq1 = df1[target_column].value_counts().sort_index()
    freq2 = df2[target_column].value_counts().sort_index()
    #print('Первый:',freq1, 'ВТОРОЙ:',freq2)
# Объединяем две серии с частотами и находим разницу
    comparison = pd.concat([freq1, freq2], axis=1, keys=["Файл_1", "Файл_2"])
    comparison.fillna(0, inplace=True)  # Меняем NaN на 0

# Рассчитываем разницу
    comparison["Разница"] = comparison["Файл_1"] - comparison["Файл_2"]

# Преобразование индекса в отдельную колонку
    comparison.reset_index(inplace=True)
    comparison.rename(columns={'index': target_column}, inplace=True)

# Присоединяем остальные столбцы из df1 и df2
    all_data = []
    for _, row in comparison.iterrows():
        value = row[target_column]
    # Берем строки из df1 и df2, соответствующие данному значению в целевом столбце
        rows_from_df1 = df1.loc[df1[target_column] == value][other_columns]
        rows_from_df2 = df2.loc[df2[target_column] == value][other_columns]
    
    # Объединяем строки из df1 и df2
        combined_rows = pd.concat([rows_from_df1.assign(Source='Файл_1'),
                               rows_from_df2.assign(Source='Файл_2')],
                              ignore_index=True)
    
    # Присваиваем всем записям соответствующую частоту и разницу
        combined_rows[target_column] = value
        combined_rows['Файл_1'] = row['Файл_1']
        combined_rows['Файл_2'] = row['Файл_2']
        combined_rows['Разница'] = row['Разница']
    
        all_data.append(combined_rows)

# Собираем всю таблицу вместе
        final_result = pd.concat(all_data, ignore_index=True)


# Переименовываем индекс в качестве отдельной колонки
        comparison.reset_index(inplace=True)
        comparison.rename(columns={'index': target_column}, inplace=True)



# Если нужно подсчитать сумму значений второго файла для каждого уникального ID
    #final_df = df.groupby('Номер ФЛС').agg({'Номер ФЛС': ['sum']})
    #final_df = df1.merge(df2, on='Номер ФЛС', how='outer')
    #final_df = pd.merge(df1, df2, left_on=['Номер ФЛС'], right_on=['Номер ФЛС'], suffixes=('_Файл_1', '_Файл_2'),  how='inner')
  
    #final_df['compare'] = final_df['Наименование услуги_Файл_1'] == final_df['Наименование услуги_Файл_2']
    
# Сохраняем результат в новый файл
    output_file = 'merged_file.xlsx'
    final_result.to_excel(output_file, index=False)
    messagebox.showinfo("Title",f"Файлы успешно объединены. Результат сохранён в '{output_file}'.")
    


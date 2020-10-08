import pandas as pd
import os
import datetime as dt
import time

# Получаем адрес до папки с файлами
path_to_files = str(input('Path to files: '))
path_to_files = os.path.normpath(path_to_files)

start_time = time.time()

# Получаем список файлов в папке и если там был txt удаляем его, чтобы не мешался
files_list = []
for file in os.listdir(path_to_files):
    extension = os.path.splitext(file)[1]
    if extension == '.txt':
        os.unlink(os.path.join(path_to_files, file))
        continue
    files_list.append(os.path.join(path_to_files, file))


# Описываем типы данных в столбцах и заголовки
data_types = {
    "Sch. GI Date Order Item": str,
    "Billing Date": str,
    "Billed Quantity": int,
    "Sold-to party": int,
    "Material": int,
    "Last Cust expected qty": str,
    "Billed Quantity": str,
    "Plant": str,
}
header = [
    'Material',	
    'Last Cust expected qty',
    'Billed Quantity',	
    'Plant', 
    'Sch. GI Date Order Item',	
    'Billing Date',	
    'Ship-to party'
]
index_name = 'LineN'

# Создаем главный фрейм, в котором будет собираться инофрмация из файлов
main_df = pd.DataFrame()
for index, file in enumerate(files_list):
    df = pd.read_excel(file)
    # Удаляем последние две строчки, где по формату SAP выводится итог
    df.drop(df.tail(2).index, inplace=True)
    # Удаляем строчки, где MRDR не было определен, ибо он все равно не будет ни к чему прикреплен
    df = df[df['Material'] != 'DUMMY']
    # Меняем формат дат
    df['Sch. GI Date Order Item'] = pd.to_datetime(df['Sch. GI Date Order Item'])
    df['Sch. GI Date Order Item'] =  df['Sch. GI Date Order Item'].dt.strftime('%d.%m.%Y')
    df['Billing Date'] = pd.to_datetime(df['Billing Date'])
    df['Billing Date'] =  df['Billing Date'].dt.strftime('%d.%m.%Y')
    # Убираем NaN, которые появились при создании фрейма
    df = df.fillna('')
    # Устанавливаем типы значений в столбцах
    df = df.astype(data_types)
    # Прикрепляем обработанные строки к основному фрейму
    main_df = main_df.append(df)

# Пересчитываем индексы
main_df.reset_index(inplace=True, drop=True)
# Делаем так, чтобы индексы начинались с 1
main_df.index += 1
# Сохраняем фрейм в txt файл с табуляциями в качестве разделитеоей
main_df.to_csv(os.path.join(path_to_files, 'res.txt'), mode='w', sep='\t', index_label=index_name)

end_time = time.time()
print('Completed in ' + str(int(end_time - start_time)) + ' sec.')
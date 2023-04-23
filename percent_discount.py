import pandas as pd    # by Kirill Kasparov, 2023
import os
import time
from pyxlsb import open_workbook as open_xlsb

vesrion = 'Программа "Дисконт" (Percent_discount). Версия: 1.05'

def updade_base(update_input):
    if update_input == '1':  # файл кривой, собираем его через модуль pyxlsb
        df = []
        count = 0
        persent = 0
        time_bstart = time.time()
        print('Обновление базы займет около 5 минут. Начинаем загрузку ...')
        with open_xlsb(new_base) as wb:
            with wb.get_sheet(1) as sheet:
                for row in sheet.rows():

                    df.append([item.v for item in row])
                    count += 1
                    if count % 20000 == 0:
                        persent += 5
                        print('Загружено', persent, '%')
                    if count == 400000:
                        break
        df = pd.DataFrame(df[1:], columns=df[0])
        df = df[['Артикул', 'Наименование товара', 'Название ТР (ТС)', 'ТК', 'ТГ', 'АГ', 'Себестоимость, руб',
                 'Признак-категория', 'Доступное кол-во', 'Цена 1', 'Процент НДС']]
        df = cleaner_art(df)
        df.to_csv(data_import, sep=';', encoding='windows-1251', index=False, mode='w')
        time_end = time.time()
        print('База обновлена. Время выполнения: ' + str(int(time_end - time_bstart)), 'сек.')
    elif update_input.lower() == '2' and os.path.exists('import.csv'):
        print('-'* 50)
        print('Товарный классификатор выводит все категории Артикулов из списка "Заказ от партнера.csv", чтобы можно было настроить файл "Спецификация.csv"')
        print('Загружаем базу данных...')
        df = pd.read_csv('import.csv', sep=';', encoding='windows-1251', dtype='unicode', nrows=400000)
        print('База данных загружена.')
    elif os.path.exists('import.csv'):  # проверяем наличие базы данных
        print('Загружаем базу данных...')
        df = pd.read_csv('import.csv', sep=';', encoding='windows-1251', dtype='unicode', nrows=400000)
        print('База данных загружена.')
    else:
        print('Файл import.csv не найден. Перезагрузите программу и подтвердите обновление базы, тогда он появится.')
        print(input())
    return df
def cleaner_art(df):    # Чистим столбцы
    for i in range(len(df.columns)):
        df['Артикул'] = df['Артикул'].fillna('0')
        df['Артикул'] = df['Артикул'].astype('str')
        df['Артикул'] = df['Артикул'].str.replace('.0', '', regex=False)
    return df
def cleaner_offer(df):    # Чистим столбцы
    if df_offer['Артикул'].count() > 0:
        df = cleaner_art(df)
        df['Скидка'] = df['Скидка'].astype('str')
        df['Скидка'] = df['Скидка'].str.replace(',', '.', regex=False)
        df['Скидка'] = df['Скидка'].astype('float64')
    if df_offer['АГ'].count() > 0:
        df['Скидка по АГ'] = df['Скидка по АГ'].astype('str')
        df['Скидка по АГ'] = df['Скидка по АГ'].str.replace(',', '.', regex=False)
        df['Скидка по АГ'] = df['Скидка по АГ'].astype('float64')
    if df_offer['ТГ'].count() > 0:
        df['Скидка по ТГ'] = df['Скидка по ТГ'].astype('str')
        df['Скидка по ТГ'] = df['Скидка по ТГ'].str.replace(',', '.', regex=False)
        df['Скидка по ТГ'] = df['Скидка по ТГ'].astype('float64')
    if df_offer['ТК'].count() > 0:
        df['Скидка по ТК'] = df['Скидка по ТК'].astype('str')
        df['Скидка по ТК'] = df['Скидка по ТК'].str.replace(',', '.', regex=False)
        df['Скидка по ТК'] = df['Скидка по ТК'].astype('float64')
    if df_offer['Название ТР (ТС)'].count() > 0:
        df['Скидка по ТР'] = df['Скидка по ТР'].astype('str')
        df['Скидка по ТР'] = df['Скидка по ТР'].str.replace(',', '.', regex=False)
        df['Скидка по ТР'] = df['Скидка по ТР'].astype('float64')
    if df_offer['Скидка на всё'].count() > 0:
        df['Скидка на всё'] = df['Скидка на всё'].astype('str')
        df['Скидка на всё'] = df['Скидка на всё'].str.replace(',', '.', regex=False)
        df['Скидка на всё'] = df['Скидка на всё'].astype('float64')
    return df
def cleaner_to_excel(df):
    df['Себестоимость, руб'] = df['Себестоимость, руб'].astype('str')
    df['Себестоимость, руб'] = df['Себестоимость, руб'].str.replace('.', ',', regex=False)
    df['Цена 1'] = df['Цена 1'].astype('str')
    df['Цена 1'] = df['Цена 1'].str.replace('.', ',', regex=False)
    df['Процент НДС'] = df['Процент НДС'].astype('str')
    df['Процент НДС'] = df['Процент НДС'].str.replace('.', ',', regex=False)
    df['Скидка'] = df['Скидка'].astype('str')
    df['Скидка'] = df['Скидка'].str.replace('.', ',', regex=False)
    df['Экономия в руб.'] = df['Экономия в руб.'].astype('str')
    df['Экономия в руб.'] = df['Экономия в руб.'].str.replace('.', ',', regex=False)
    df['Цена со скидкой'] = df['Цена со скидкой'].astype('str')
    df['Цена со скидкой'] = df['Цена со скидкой'].str.replace('.', ',', regex=False)
    df['КТН'] = df['КТН'].astype('str')
    df['КТН'] = df['КТН'].str.replace('.', ',', regex=False)
    df['ВП'] = df['ВП'].astype('str')
    df['ВП'] = df['ВП'].str.replace('.', ',', regex=False)

    return df
def merge_main_data(df, df_art, df_offer, mode):
    df_exp = df_art.merge(df, on='Артикул')
    df_exp['Цена 1'] = df_exp['Цена 1'].fillna('0')
    df_exp['Цена 1'] = df_exp['Цена 1'].astype('float64')
    df_exp['Себестоимость, руб'] = df_exp['Себестоимость, руб'].fillna('0')
    df_exp['Себестоимость, руб'] = df_exp['Себестоимость, руб'].astype('float64')
    df_exp['Доступное кол-во'] = df_exp['Доступное кол-во'].fillna('0')
    df_exp['Доступное кол-во'] = df_exp['Доступное кол-во'].astype('str')
    df_exp['Доступное кол-во'] = df_exp['Доступное кол-во'].str.replace('.0', '', regex=False)
    if df_offer['АГ'].count() == 0 and mode != '2':
        del df_exp['АГ']
    if df_offer['ТГ'].count() == 0 and mode != '2':
        del df_exp['ТГ']
    if df_offer['ТК'].count() == 0 and mode != '2':
        del df_exp['ТК']
    return df_exp
def total_discount(df_exp, df_offer):
    if df_offer['Артикул'].count() > 0:
        df_offer_art = df_offer[['Артикул', 'Скидка']]
        df_exp = df_exp.merge(df_offer_art, on='Артикул', how='outer')
        df_exp = df_exp[~pd.isnull(df_exp['Наименование товара'])]  # чистим пустые значения

    if df_offer['АГ'].count() > 0:
        df_offer_ag = df_offer[['АГ', 'Скидка по АГ']]
        df_exp = df_exp.merge(df_offer_ag, on='АГ', how='outer')  # добавляем скидки по АГ
        df_exp = df_exp[~pd.isnull(df_exp['Наименование товара'])]  # чистим пустые значения
        df_exp['Скидка'] = df_exp['Скидка'].fillna(df_exp['Скидка по АГ'])  # добавляем скидку
        del df_exp['Скидка по АГ']  # прячем лишнее

    if df_offer['ТГ'].count() > 0:
        df_offer_tg = df_offer[['ТГ', 'Скидка по ТГ']]
        df_exp = df_exp.merge(df_offer_tg, on='ТГ', how='outer')  # добавляем скидки по ТГ
        df_exp = df_exp[~pd.isnull(df_exp['Наименование товара'])]  # чистим пустые значения
        df_exp['Скидка'] = df_exp['Скидка'].fillna(df_exp['Скидка по ТГ'])  # добавляем скидку
        del df_exp['Скидка по ТГ']  # прячем лишнее

    if df_offer['ТК'].count() > 0:
        df_offer_tk = df_offer[['ТК', 'Скидка по ТК']]
        df_exp = df_exp.merge(df_offer_tk, on='ТК', how='outer')  # добавляем скидки по ТК
        df_exp = df_exp[~pd.isnull(df_exp['Наименование товара'])]  # чистим пустые значения
        df_exp['Скидка'] = df_exp['Скидка'].fillna(df_exp['Скидка по ТК'])  # добавляем скидку
        del df_exp['Скидка по ТК']  # прячем лишнее

    if df_offer['Название ТР (ТС)'].count() > 0:
        df_offer_tr = df_offer[['Название ТР (ТС)', 'Скидка по ТР']]
        df_exp = df_exp.merge(df_offer_tr, on='Название ТР (ТС)', how='outer')  # добавляем скидки по ТР
        df_exp = df_exp[~pd.isnull(df_exp['Наименование товара'])]  # чистим пустые значения
        df_exp['Скидка'] = df_exp['Скидка'].fillna(df_exp['Скидка по ТР'])  # добавляем скидку
        del df_exp['Скидка по ТР']  # прячем лишнее

    if df_offer['Скидка на всё'].count() == 1:
        df_offer_all = df_offer['Скидка на всё'][df_offer['Скидка на всё'].notna()]  # забираем скидку (из любой ячейки)
        df_exp['Скидка'] = df_exp['Скидка'].fillna(list(df_offer_all)[0])  # добавляем скидку
    elif df_offer['Скидка на всё'].count() == 0:
        df_offer_all = df_offer['Скидка на всё'][df_offer['Скидка на всё'].notna()]  # забираем скидку (из любой ячейки)
        df_exp['Скидка'] = df_exp['Скидка'].fillna(float(0))  # добавляем скидку
    else:
        df_offer_all = df_offer['Скидка на всё'][df_offer['Скидка на всё'].notna()]  # забираем скидку (из любой ячейки)
        df_exp['Скидка'] = df_exp['Скидка'].fillna(float(0))  # добавляем скидку
        print('Внимание! В поле "Скидка на всё" введено больше одного значения')
    return df_exp

# Тело кода
cmd = 'mode 180,30'
os.system(cmd)

# определяем абсолютный путь к файлам
data_import = os.getcwd().replace('\\', '/') + '/' + 'import.csv'
data_export = os.getcwd().replace('\\', '/') + '/' + 'Результат.csv'

# инфо
print(vesrion)
print('-' * 50)
print('Описание: программа берет список артикулов из файла "Заказ от партнера.csv" и проставляет цены с учетом процента скидки от цены 1 КПЛ.')
print('Проставьте скидки в файле "Спецификация.csv" до запуска программы.')
print('Можно проставить скидки на Актикул, Ассортиментную группу (АГ), Товарную категорию (ТК), Группу (ТГ), Весь рынок (ТР) или общую скидку на все.')
print('Сценарии скидок можно комбинировать. Например, установить скидки только на список Артикулов и Товарный рынок.')
print('Приоритет цены выстраивается от меньшей категории товара к большей.')
print('Если на весь товарный рынок Мебели применена скидка 10%, а на отдельный артикул кресла 15%, применится скидка 15%.')
print('Цены со скидокой вы увидите в файле "Результат.csv".')
print('-' * 50)
# ищем обновление
new_base = []
if os.path.exists("\\/storage-msk-tu.komus.net/stor/PM/!АНАЛИТИКА и отчетность/Super_klass"):  # Прогоняем поиск обновлений базы
    for root, dirs, files in os.walk(
            "\\/storage-msk-tu.komus.net/stor/PM/!АНАЛИТИКА и отчетность/Super_klass"):  # такой вариант кода позволяет обойти весь каталог
        for file in files:
            if file.endswith(".xlsb") and ('super_klass' in file.lower()):
                new_base.append(os.path.join(root) + '/' + os.path.join(file))
    if len(new_base) > 0:
        new_base = max(new_base, key=os.path.getctime)  # получаем последнюю версию
    print('Последняя вервия базы: ', new_base.split('/')[-1].replace('Super_klass.xlsb', ''))
else:
    print('Обновление невозможно. Нет доступа к сетевому диску: \\\storage-msk-tu.komus.net')

# загружаем базу
print('Для установки новой версии введите "1".')
print('Чтобы получить товарный классификатор, введите "2".')
mode = input('Чтобы продожлить работу с текущей базой, нажмите Enter: ')
df = updade_base(mode)

# Заказ от партнера
if os.path.exists('Заказ от партнера.csv'):  # проверяем наличие базы данных
    df_art = pd.read_csv('Заказ от партнера.csv', sep=';', encoding='windows-1251', dtype='unicode', nrows=2000)
else:
    print('-' * 50)
    print('Файл "Заказ от партнера.csv" не найден.')
    df_art = pd.DataFrame({'Артикул': ['13500', '120', '906433', '273572', '515916', '396249']})
    df_art.to_csv('Заказ от партнера.csv', sep=';', encoding='windows-1251', index=False, mode='w')
    print('Создали для вас новый шаблон, добавили несколько артикулов для примера. Заполните файл и перезапустите программу.')
    print(input())
df_art = cleaner_art(df_art)

# Спецификация
if os.path.exists('Спецификация.csv'):  # проверяем наличие базы данных
    df_offer = pd.read_csv('Спецификация.csv', sep=';', encoding='windows-1251', dtype='unicode', nrows=2000)
else:
    print('-' * 50)
    print('Файл "Спецификация.csv" не найден.')
    df_offer = pd.DataFrame({'Артикул': ['13500'],
                             'Скидка': ['0,11'],
                             'АГ': ['Калькуляторы настольные'],
                             'Скидка по АГ': ['0,12'],
                             'ТГ': ['Скобы для степлеров №10'],
                             'Скидка по ТГ': ['0,13'],
                             'ТК': ['Кофе, какао'],
                             'Скидка по ТК': ['0,14'],
                             'Название ТР (ТС)': ['ТР Гигиеническая продукция'],
                             'Скидка по ТР': ['0,15'],
                             'Скидка на всё': ['0,16']})
    df_offer.to_csv('Спецификация.csv', sep=';', encoding='windows-1251', index=False, mode='w')
    print('Создали для вас новый шаблон, добавили несколько строк для примера. Заполните файл и перезапустите программу.')
    print(input())
df_offer = cleaner_offer(df_offer)

# собираем таблицу для экспорта
time_start = time.time()
df_exp = merge_main_data(df, df_art, df_offer, mode)

# собираем скидку
df_exp = total_discount(df_exp, df_offer)

df_exp['Экономия в руб.'] = df_exp['Цена 1'] * df_exp['Скидка']
df_exp['Цена со скидкой'] = df_exp['Цена 1'] - df_exp['Экономия в руб.']
df_exp['КТН'] = df_exp['Цена со скидкой'] / df_exp['Себестоимость, руб']
df_exp['ВП'] = df_exp['Цена со скидкой'] - df_exp['Себестоимость, руб']
df_exp['Низкий КТН!'] = df_exp['КТН'] < 1
if sum(df_exp['Низкий КТН!']) == 0:
    del df_exp['Низкий КТН!']
else:
    print('Внимание! Обнаружены строки с КТН ниже 1.00. Проверьте столбец "Низкий КТН" в файле "Результат.csv"')

# сохраняем результат
df_exp = cleaner_to_excel(df_exp)
del df_exp['Себестоимость, руб']

while True:  # проверка, если файл открыт
    try:
        df_exp.to_csv(data_export, sep=';', encoding='windows-1251', index=False, mode='w')
        break
    except IOError:
        input('Для сохранения данных необходимо закрыть файл "Результат.csv"')

time_end = time.time()
print('-' * 50)
print('Время выполнения: ' + str(int(time_end - time_start)), 'сек.')
print('Данные сохранены в файл:', data_export)
print(input())
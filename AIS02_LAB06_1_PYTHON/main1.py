import xlsxwriter  # Это модуль Python для записи файлов в формате Excel 2007+ XLSX

try:
    array = ['Илья', 'Мади', 'Артем', 'Искандер']
    array_place = ['1', '2', '3', '4']
    array_prize = ['5000000', '2500000', '1250000', '0']
    my_file = 'AIS.xlsx'  # Имя файла
    book = xlsxwriter.Workbook(my_file)  # Создание файла
    sheet = book.add_worksheet()  # Добавление в него книги
    sheet.set_column('A:A', 35)  # Установка ширины колонки
    sheet.set_column('B:B', 20)
    sheet.set_column('C:C', 15)
    bold = book.add_format({'bold': True})  # Формат жирного текста
    sheet.write('A1', 'Результаты участников конкурса', bold)  # Выдача текста в ячейку
    sheet.write('B1', 'Позиции участников', bold)  # Выдача текста в ячейку
    sheet.write('C1', 'Выигрыш', bold)  # Выдача текста в ячейку
    itr = 1
    for i in array:
        sheet.write(itr, 0, i)  # Выдача значения в ячейку 3 строка 1 столбец [2,0]
        sheet.write(itr, 1, array_place[itr - 1] + ' место')  # Выдача значения в ячейку 3 строка 1 столбец [2,0]
        sheet.write(itr, 2, array_prize[itr - 1] + ' тенге')  # Выдача значения в ячейку 3 строка 1 столбец [2,0]
        itr += 1

    sheet.insert_image('D1', 'places.png', {'x_scale': 0.50, 'y_scale': 0.50})  # Вставка в ячейку картинки
    book.close()  # Закрытие файла
except Exception as a:  # Обработка ошибок
    print("Error!")
    print(a)

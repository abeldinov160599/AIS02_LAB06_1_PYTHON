from docx import Document
from docx.shared import Inches
import xlsxwriter
employees = {'Алексей':'8 702 342 12 64', 'Иван': '8 344 342 64 87', 'Ольга':' 8 708 435 32 54' ,
              'Сергей': '8 745 321 75 99',' Илья': '8 999 323 43 12' , 'Искандер':'8 721 721 32 77'}
for x in employees:
 document = Document()
 document.add_heading('Школа программирования "ItPro"', 0)
 p = document.add_paragraph('Сотрудник - ' + x + '\n')
 p.add_run('Сотовый телефон: ' + employees[x]).bold = True
 document.add_heading('Наша школа научит вас программировать, '
                      'создавать различные сайты и приложения с помощью '
                      'C++, JAVA, Python, HTML\CSS, JavaScript, PHP и C#.', level=1)
 document.add_paragraph('')
 document.add_picture('12345.png', width=Inches(3.5))
 document.add_page_break()
 document.save(x + '.docx')
try:
 my_file = 'demo.xlsx' # Имя файла
 book = xlsxwriter.Workbook(my_file) # Создание файла
 sheet = book.add_worksheet() # Добавление в него книги
 sheet.set_column('A:A', 20) # Установка ширины колонки
 bold = book.add_format({'bold': True}) # Формат жирного текста
 sheet.write('B1', 'Список сотрудников')
 i = 0
 for x in employees:
  i+=1
  if (i <=6):
   sheet.write(i,0, i)
   sheet.write(i, 1 , x)
   sheet.write(i, 2, employees[x])
  # Выдача текста в ячейку
 sheet.insert_image('E1', '12345.png',{'x_scale': 0.7, 'y_scale': 0.7}) # Вставка в ячейку картинки
 book.close() # Закрытие файла
except Exception as a: # Обработка ошибок
 print("Error!")
 print(a)
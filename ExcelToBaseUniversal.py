# ExcelToBaseUniversal.py
# Программа экспортирует данные из файла my_excel.xls в базу MySQL (имя базы my_base, имя таблицы my_table)

# Необходимы файлы:

# config.py - настройки соединения с базой
# excel.xls - сам файл Excel 

# Первая строчка файлов - названия столбцов
# Певый столбец - нумерация (id в базе)

# Обнуления ключей в базе:
# ALTER TABLE table AUTO_INCREMENT=0


excel_file = 'my_excel.xls'
table = 'my_table'



try:
    import pymysql.cursors
except ImportError:
    print ("Необходимо установить pymysql!")

try: 
	import config
except ImportError:
	print ("Проверьте наличие файла config.py (!)\r\n")

import xlrd
import os

my_os = os.name
# очистка экрана

if my_os == 'nt':
	os.system('cls')  		# on windows
else:	 
	os.system('clear')		# on linux / os x


print ("OS : ", my_os)
print ("Server :", config.server)
print ("Base :", config.base)
print ("User :", config.user)

mass = []  			# Массив - список
separator = '*'		# разделитель, разделяем им строки, помещаемые в массив - список
i = 0				# счётчик	

# Открываем XLS
rb = xlrd.open_workbook(excel_file,formatting_info=True)
# Обращаемся к листу
sheet = rb.sheet_by_index(0)
# Кол-во строк на листе
row_number = sheet.nrows

print ("Строк на листе: ", row_number, "\n\r")


for rownum in range(sheet.nrows):
	row = sheet.row_values(rownum)
	#string = row[0] + separator + row[1] + separator + row[2]	# разделяем строки, помещаемые в массив - список (fio, book, link)
	

	string = row
	print (string)
	mass.append (string)
	

title = mass[0]
print ("Title = ", title)


# подключаемся к базе данных (не забываем указать кодировку, а то в базу запишутся иероглифы)
db = pymysql.connect(host=config.server, user=config.user, passwd=config.passw, db=config.base, charset='utf8')
# формируем курсор, с помощью которого можно исполнять SQL-запросы
cursor = db.cursor()

query = "CREATE TABLE IF NOT EXISTS " + table
query += " (" + title[0] + " int(11), " + title[1] +" varchar(50), " + title [2] +" int(11), " + title [3] + " varchar (3))"
print (query)

# исполняем SQL-запрос
cursor.execute(query)

# SQL запрос
i = 0
for element in mass:
	print (element)
	i += 1
	if (i != 1):
		id = int (element [0])
		id = abs (id)
		query = "INSERT INTO " + table + " (" +  title[0] +", " +  title[1] + ", " + title[2] + ", "+  title[3] + ") VALUES ("+ str(id)+", '" + str(element [1]) +"', '"  + str(element [2]) +"', '" + str(element [3]) + "')"
		print ("Query :", query)
		# исполняем SQL-запрос
		cursor.execute(query)

	
# применяем изменения к базе данных
db.commit()

# Разрываем подключение и отключаемся от сервера
db.close()
cursor.close()
print ("\n\rВсё нормуль, данные загружены!")
import re
import requests
from bs4 import BeautifulSoup
import cx_Oracle
import win32com.client
import os

os.environ["NLS_LANG"] = ".AL32UTF8"  # кодировка для чтения русских символов

ex_fil = r'C:\Users\Excel_file.xlsm' # путь к файлу

row = 2 # номер строки в excel файле (1 строка - заголовок)
slovar = {} # словарь для парсера
date_st = '2019-12-21' # дата начала периода
date_end = '2020-01-20' # дата конец периода
txt = 'project = EIPSR  and status = Closed and (resolution = Incomplete or resolution = Fixed) and updatedDate > ' + date_st + ' and updatedDate < ' + date_end # запрос в jira
url = 'http://jira...jspa' # url к jira
jira_username = 'username'
jira_password = 'password'

SERVER = '10.10.100.10'
PORT = '5432'
SID = 'SID'
UID = 'User_Name'
PWD = 'User_Pass'
conn = cx_Oracle.connect(UID + '/' + PWD + '@' + SERVER + ':' + PORT + '/' + SID) # подключение к БД
cursor = conn.cursor() # создаем курсор

# передаем наш SQL запрос
cursor.execute('''SELECT      
                            'Мой Отдел' as Группирующий_узел, 
                                j.pkey  as Код_инцидента,
                              j.created as Дата_создания,
                                     '' as Дата_закрытия, -- Отсутствует в БД
                                  'Нет' as Нарушено_SLA,
                                     '' as Обосновав_Нарушения_срока_SLA, -- Отсутствует в БД
                            (CASE j.priority 
                              WHEN '1' THEN '1 – Наивысший'
                              WHEN '2' THEN '2 – Высокий'
                              WHEN '3' THEN '3 – Средний'
                              WHEN '4' THEN '4 – Низкий'
                            END
                                       )as Приоритет,
                            (CASE j.priority 
                              WHEN '1' THEN '1 – Наивысший'
                              WHEN '2' THEN '2 – Высокий'
                              WHEN '3' THEN '3 – Средний'
                              WHEN '4' THEN '4 – Низкий'
                            END
                                       )as Фактический_приоритет,
                            (CASE j.issuetype 
                              WHEN '7' THEN 'Консультация'
                              WHEN '12' THEN 'Инцидент'
                              WHEN '13' THEN 'Проблема'
                            END
                                       )as Тип,
                               'Закрыт' as Статус_заявки,
                                'ИР...' as Услуга, -- Отсутствует в БД
                              j.summary as Аннотация, 
  TO_CHAR(SUBSTR(j.description,1,4000)) as Описание, 
                                     '' as Решение,
               'Эксплуатация ЕИП. РТИ.' as Группа_поддержки,
                                     '' as Код_закрытия,
                             j.assignee as Кем_закрыто,  
                             j.assignee as Назначенный,  
                             j.reporter as Кем_открыто,  
                             j.reporter as Получатель  
FROM some_table1 j 
INNER JOIN some_table2 i 
ON i.id = j.issuestatus           
WHERE 1 = 1
        AND j.issuestatus = 6 -- Закрыт
        AND (j.resolution = 1 or j.resolution = 4) -- Резолюция Incomplete или Fixed
        AND j.pkey LIKE 'EIPSR%'
        AND (j.updated BETWEEN TO_DATE( TO_CHAR(DATE ' ''' + str(date_st) + ''' ', 'dd.mm.yyyy') || ' 00:00:00','dd.mm.yyyy hh24:mi:ss') 
        AND TO_DATE( TO_CHAR(DATE ' ''' + str(date_end) + ''' ', 'dd.mm.yyyy') || ' 00:00:00','dd.mm.yyyy hh24:mi:ss'))
ORDER BY j.pkey DESC ''')

# Переходим к этапу переноса полученных данных в excel файл

Excel = win32com.client.Dispatch("Excel.Application")  # создаем Application
wb = Excel.Workbooks.Open(ex_fil) # открываем excel файл
sheet = wb.Sheets('ЕИП') # обращаемся к нашему листу


for i in cursor.fetchall(): # перебираем и записываем последовательность полученных данных из БД
    row += 1 # номер строки в excel
    for j in range(len(i)):
        try:
            sheet.Cells(row, j + 1).value = i[j]  # записываем в excel данные из SQL
        except TypeError:
            sheet.Cells(row, j + 1).value = i[j].read()  # записываем в excel данные из SQL в формате CLOB

wb.Save()  # сохраняем рабочую книгу
wb.Close()  # закрываем ее

cursor.close()  # закрываем курсор
conn.close()  # закрываем соединение с БД

# Переходим к этапу получения недостающих данных через парсинг 

s = requests.Session()  # создаем сессию
r = s.post(url, data={'os_username': jira_username, 'os_password': jira_password})  # авторизация
r = s.post(url, data={'jqlQuery': txt, 'runQuery': 'true'})  # отправляем пост на нужные тикеты
r = s.get(url).text  # получаем код старницы

soup = BeautifulSoup(r, 'lxml')  # обрабатываем код страницы 
tic = re.findall(r'EIPSR.\d+', soup.text)  # найти все тикеты
per_x = len(tic)  # переменная х для дальнейшей подстановки в 

for zz in tic:  # перебор по 1 тикету
    r = s.get(
        'http://jira/' + zz + '?page=com.atlassian.jira...').text  # получаем код старницы
    soup = BeautifulSoup(r, 'lxml')  # обрабатываем код страницы с тикетом в читаемый вид 
    string = soup.text  # приравниваем весть текст в переменную
    result = []  # временный список в который будет подгружаться хронология статусов
    rownt = -1  # счетчик для словаря чтобы отделять одну резалюцию от другой
    result1 = re.findall(
        r'\d{2}.\d{2}.\d{4}.\d{2}:\d{2}:\d{2}|\bОткрыт|\bIncomplete|\bReopened|\bResolved|\bIn Progress|\bFixed|\bClosed',
        string)  # поиск даты и статуса
    pos = result1.index('Resolved') + 2  # находим позицию ненужного индекса
    for j in range(pos, len(result1)):  # перебираем с уже нужной позиции все статусы и даты
        if not (result1[j - 1][0].isdigit() and result1[j][0].isdigit()):  # условие если у нас не две подряд идущие даты
            if result1[j - 1][0].isdigit():  # если данный эдемент дата, то создаем новый список для отделения
                result.append([])  # добавляем новый список
                rownt += 1  # счетчик чтобы заносить данные в только что созданный список
            result[rownt].append(result1[j - 1])  # добавляем значения в текущий или только что созданный тикет
            if j + 1 == len(result1):  # проверяем последний элемент или нет, что бы добавить последнее значение
                result[rownt].append(result1[j])  # добавляем последнее значение

    ps, res, sps, spsc, mayk = 'Открыт', 'Resolved', [], [], 1  # переменные: статус тикета/ резолюция/ список для статусов/ список для закрытия/ маяк для определения закрыт или резолюция
    for stat in result:
        if ps in stat:  # если начало статуса совпадает (начинаем с Открыт)
            if res in stat:  # если есть
                sps.append(stat[0])  # добавляем дату
                sps.append(stat[-1])  # добавляем конечный статус
                ps = stat[-1]  # меняем статус (к примеру был открыт меняем на текущий) для дальнейшей проверки
            else:
                if 'In Progress' in stat:  # проверяем есть ли инпрогресс, что бы его не заносить в список
                    ps = stat[-1]  # меняем статус на инпрогресс
                if 'Closed' in stat:  # проверяем на закрытие
                    mayk += 1
                    if mayk % 2 == 0:  # если 0, то статус Closed не был или был но потом сменился на Resolved/ если 1 то сейчас статус Closed и выполняется другая ветка условий
                        spsc.append(stat[0])  # добавляем в отдельный список дату закрытия
                        spsc.append('Closed')  # добавляем в отдельный список статус закрытия
                        if 'Incomplete' in stat:  # проверяем какой из статусов
                            sps.append(stat[0])  # добавляем дату
                            sps.append('Incomplete')  # добавляем если прошел проверку
                        if 'Fixed' in stat:  # проверяем какой из статусов
                            sps.append(stat[0])  # добавляем дату
                            sps.append('Fixed')  # добавляем если прошел проверку
                    else:
                        sps.append(stat[0])  # добавляем дату
                        sps.append(stat[-1])  # добавляем конечный статус
                        ps = stat[-1]  # меняем статус (к примеру был открыт меняем на текущий) для дальнейшей проверки
        else:  # если стутус не совпал
            if res in stat:  # если есть Resolved
                if 'Closed' in stat:  # проверяем на закрытие
                    mayk += 1
                    if mayk % 2 == 0:  # если 0, то статус Closed не был или был но потом сменился на Resolved/ если 1 то сейчас статус Closed и выполняется другая ветка условий
                        spsc.append(stat[0])  # добавляем в отдельный список дату закрытия
                        spsc.append('Closed')  # добавляем в отдельный список статус закрытия
                        if 'Incomplete' in stat:  # проверяем какой из статусов
                            sps.append(stat[0])  # добавляем дату
                            sps.append('Incomplete')  # добавляем если прошел проверку
                        if 'Fixed' in stat:  # проверяем какой из статусов
                            sps.append(stat[0])  # добавляем дату
                            sps.append('Fixed')  # добавляем если прошел проверку
                    else:
                        sps.append(stat[0])  # добавляем дату
                        sps.append(stat[-1])  # добавляем конечный статус
                        ps = stat[-1]  # Меняем статус (к примеру был открыт меняем на текущий) для дальнейшей проверки

    slovar[zz] = [sps, spsc, []]  # добавляем в словарь созданный список нашей хронологии

# Находим название системы

for zz in tic:  # перебор по 1 тикету
    r = s.get(
        'http://jira/' + zz + '?page=com.atlassian.jira...').text  # получаем код старницы
    soup = BeautifulSoup(r, 'lxml')  # обрабатываем код страницы с тикетом в читаемый вид lxml
    string = soup.text  # приравниваем весть текст в переменную
    result = []  # временный список в который будет подгружаться хронология статусов
    rownt = -1  # счетчик для словаря чтобы отделять одну резалюцию от другой
    for line in string.strip().splitlines():
        if re.search(r'(\AИР ([\w]+)([ \w\()-]*))|(\AИнфраструктура$)', line):  # поиск компонента
            slovar[zz][2] = [line]  # добавляем в словарь название платформ
            break

# Переходим к этапу переноса недостающих данных в excel файл

wb = Excel.Workbooks.Open(ex_fil) # открываем excel файл
sheet = wb.Sheets('ЕИП') # обращаемся к нашему листу1
sheet2 = wb.Sheets('Расчет') # обращаемся к нашему листу2

# Записываем последовательность

kor_z = 8 # столбец 8 в excel
for x_kor in range(3, per_x+3): # позиции для тикетов, начинаем с 3й ячейки и до конца (сколько у нас нашолось задачь).
    for key in slovar.keys(): # перебираем наши ключи для сравнения с ключами в excel
        if sheet.Cells(x_kor, 2).value == key: # если ключи равны - заполняем данные
            sheet.Cells(x_kor, 11).value = slovar[key][2] # лист ЕИП - название ИР
            sheet.Cells(x_kor, 4).value = slovar[key][1][0] # лист ЕИП - дата закрытия
            sheet2.Cells(x_kor, 2).value = key # лист Расчет - номер задачи
            sheet2.Cells(x_kor, 3).value = sheet.Cells(x_kor, 3).value # лист Расчет - дата создания
            sheet2.Cells(x_kor, 4).value = slovar[key][1][0]  # лист Расчет - дата закрытия
            for y_kor in slovar[key][0]: # перебираем хронологию статусов и время
                sheet2.Cells(x_kor, kor_z).value = y_kor
                kor_z +=1
            kor_z = 8

wb.Save()  # сохраняем рабочую книгу
wb.Close()  # закрываем ее
Excel.Quit()  # закрываем COM объект

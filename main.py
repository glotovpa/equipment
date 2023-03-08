import os
import openpyxl


def import_comment(comment_f):
    xlsx = openpyxl.load_workbook('комментарии по оборудованию.xlsx')
    # выбираем из файла лист
    sheet = xlsx['Лист1']
    # проходим по строчкам с листа
    not_first = False
    for line in sheet:
        if not_first:
            fio = str(line[0].value)
            nomenklatura = str(line[1].value)
            sn = str(line[2].value)
            comment_equipment = str(line[3].value)
            comment_f.append((fio, nomenklatura, sn, comment_equipment))
        else:
            not_first = True


def import_virt_sklad(data_f, comment_f):
    os.chdir('in')
    for filename in os.listdir("."):
        # читаем файл
        xlsx = openpyxl.load_workbook(filename)
        # выбираем из файла лист
        sheet = xlsx['Worksheet']
        # проходим по строчкам с листа
        not_first = False
        for line in sheet:
            if not_first:
                service = str(line[2].value)
                fio = str(line[4].value)
                nomen = str(line[7].value)
                numbers = str(line[8].value)
                quality = str(line[9].value)
                summa = str(line[10].value)
                date = str(line[11].value)
                sn = str(line[13].value)
                status = str(line[14].value)
                comment_current = ''
                # поиск комментария по серийному номеру или связке фио и номенклатура
                for comment_line in comment_f:
                    fio_comment = comment_line[0].lower()
                    nomen_comment = comment_line[1].lower()
                    sn_comment = comment_line[2].lower()
                    comment_from = comment_line[3]
                    if (sn.lower() == sn_comment) and not(sn == 'None'):
                        comment_current = comment_from
                    elif (fio.lower() == fio_comment) and (nomen.lower() == nomen_comment) and sn == 'None':
                        comment_current = comment_from
                data_f.append((service, fio, nomen, numbers, quality, summa, date, sn, status, comment_current))
            else:
                not_first = True


def statistics(comment_f, data_f):
    record_comment = len(comment_f)
    print('Количество записей с комментариями: ', record_comment)
    record_data = len(data_f)
    print('Количество записей c оборудованием: ', record_data)
    number_with_sn = 0
    number_with_no_sn = 0
    number_all = 0
    number_instrument = 0
    number_sr = 0
    summa_all = 0
    summa_sr = 0
    summa_insrument = 0
    for line_data in data_f:
        number = int(line_data[3])
        number_all += number
        if line_data[7] == 'None':
            number_with_no_sn += number
        else:
            number_with_sn += number
    print('Количество общее: ', number_all)
    print('Количество с серийным номером: ', number_with_sn)
    print('Количество без серийного номера: ', number_with_no_sn)

# comet - список с комментариями по оборудованию
# ФИО сотрудника, Номенклатура, Серийный номер, Комментарий
comment = []
# data - список с информацией по всем сотрудникам
data = []
# загрузка файла "комментарии по оборудованию.xlsx"
import_comment(comment)
print(comment[0])
print(comment[1])

# загрузка информации из xlsx файлов из папки In
import_virt_sklad(data, comment)
print(data[0])
print(data[1])

statistics(comment, data)

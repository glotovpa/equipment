import os
import openpyxl


def import_comment(comment):
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
            comment.append((fio, nomenklatura, sn, comment_equipment))
        else:
            not_first = True


def import_virt_sklad(data, comment):
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
                nomenklatura = str(line[7].value)
                numbers = str(line[8].value)
                quality = str(line[9].value)
                summa = str(line[10].value)
                date = str(line[11].value)
                sn = str(line[13].value)
                status = str(line[14].value)
                comment_current = ''
                # поиск комментария по серийному номеру или связке фио и номенклатура
                for comment_line in comment:
                    fio_comment = comment_line[0].lower()
                    nomenklatura_comment = comment_line[1].lower()
                    sn_comment = comment_line[2].lower()
                    comment_from = comment_line[3]
                    if (sn.lower() == sn_comment) and not(sn == 'None'):
                        comment_current = comment_from
                    elif (fio.lower() == fio_comment) and (nomenklatura.lower() == nomenklatura_comment) and sn == 'None':
                        comment_current = comment_from
                data.append((service, fio, nomenklatura, numbers, quality, summa, date, sn, status, comment_current))
            else:
                not_first = True

# commnet - список с комментариями по оборудованию
# ФИО сотрудника, Номенклатура, Серийный номер, Комментарий
comment = []
# data - список с информацией по всем сотрудникам
data = []
# загрузка файла "комментарии по оборудованию.xlsx" из корневой папки
import_comment(comment)
print(comment[0])
print(comment[1])
print(len(comment))

# загрузка инфомрации из xlsx файлов из папки In
import_virt_sklad(data, comment)

print(data[0])
print(data[1])
print(len((data)))

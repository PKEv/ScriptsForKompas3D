import os
import re
import subprocess
import pythoncom
from win32com.client import Dispatch, gencache
from tkinter import Tk
# from tkinter.filedialog import askopenfilenames
from tkinter import filedialog

##------------------------------------------------------------------------------
##
# Программа вывода отчета по атрибутам из фалов извещений об изменений
##
##------------------------------------------------------------------------------


# Подключение к API7 программы Компас 3D
def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const


# Функция проверки, запущена-ли программа КОМПАС 3D
def is_running():
    proc_list = \
    subprocess.Popen('tasklist /NH /FI "IMAGENAME eq KOMPAS*"', shell=False, stdout=subprocess.PIPE).communicate()[0]
    return True if proc_list else False


# Посчитаем количество листов каждого из формата
def amount_sheet(doc7):
    sheets = {"A0": 0, "A1": 0, "A2": 0, "A3": 0, "A4": 0, "A5": 0}
    for sheet in range(doc7.LayoutSheets.Count):
        format = doc7.LayoutSheets.Item(sheet).Format  # sheet - номер листа, отсчёт начинается от 0
        sheets["A" + str(format.Format)] += 1 * format.FormatMultiplicity
    return sheets


# Прочитаем основную надпись чертежа
def stamp(doc7):
    for sheet in range(doc7.LayoutSheets.Count):
        doc7.LayoutSheets.Item(sheet).Update()
        style_filename = os.path.basename(doc7.LayoutSheets.Item(sheet).LayoutLibraryFileName)
        style_number = int(doc7.LayoutSheets.Item(sheet).LayoutStyleNumber)

        # Форматка для извещений
        if style_number == 47:
            stamp = doc7.LayoutSheets.Item(sheet).Stamp

            return {"Scale": re.findall(r"\d+:\d+", stamp.Text(6).Str)[0],
                    "FirstUsage": stamp.Text(25).Str,           # Первичное применение

                    "Checked": stamp.Text(111).Str,             # Проверил
                    "TChecked": stamp.Text(112).Str,            # Т.Контролер
                    "NChecked": stamp.Text(114).Str,            # Н.Контролер
                    "Approved": stamp.Text(115).Str,            # Утвердил

                    "Number": stamp.Text(2).Str,                # Номер извещения
                    "ChangedDoc": stamp.Text(203).Str,          # номер изменяемого документа
                    "DateCreate": stamp.Text(1).Str,            # дата выпуска
                    "TimeChange": stamp.Text(206).Str,          # срок изменений
                    "PageCount": stamp.Text(8).Str,             # количество листов

                    "CauseOfChange": stamp.Text(210).Str,       # причина изменений
                    "NumberCauseOfChange": stamp.Text(152).Str, # номер причины изменений

                    "SamplesAvailable": stamp.Text(212).Str,    # Указание о заделе
                    "Introduction": stamp.Text(212).Str,        # Указание о внедрении
                    "Applicability": stamp.Text(214).Str,       # Применяемость
                    "Subscribers": stamp.Text(215).Str,         # Рассылка
                    "Attachment": stamp.Text(215).Str,          # Приложение
                    "ChangeNumber": stamp.Text(17).Str,         # Номер изменения

                    "Designer": stamp.Text(110).Str}            # Составил
    return {}


# Подсчет технических требований, в том случае, если включена автоматическая нумерация
def count_demand(doc7, module7):
    IDrawingDocument = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IDrawingDocument'], pythoncom.IID_IDispatch)
    drawing_doc = module7.IDrawingDocument(IDrawingDocument)
    text_demand = drawing_doc.TechnicalDemand.Text

    count = 0  # Количество пунктов технических требований
    for i in range(text_demand.Count):  # Прохоим по каждой строчке технических требований
        if text_demand.TextLines[i].Numbering == 1:  # и проверяем, есть ли у строки нумерация
            count += 1

    # Если нет нумерации, но есть текст
    if not count and text_demand.TextLines[0]:
        count += 1

    return count

def specWork(doc7):
    IDrawingDocument = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IDrawingDocument'], pythoncom.IID_IDispatch)


def parse_design_documents(paths):
    is_run = is_running()  # True, если программа Компас уже запущена

    module7, api7, const7 = get_kompas_api7()  # Подключаемся к программе
    app7 = api7.Application  # Получаем основной интерфейс программы
    app7.Visible = True  # Показываем окно пользователю (если скрыто)
    app7.HideMessage = const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы

    table = []  # Создаём таблицу парметров
    for path in paths:
        print("Чтение файла: " + path + "\n")
        doc7 = app7.Documents.Open(PathName=path,
                                   Visible=False,
                                   ReadOnly=True)  # Откроем файл в видимом режиме без права его изменять

        row = amount_sheet(doc7)  	# Посчитаем кол-во листов каждого формат
        row.update(stamp(doc7))  	# Читаем основную надпись
        row.update({
            "Filename": doc7.Name,  # Имя файла
        })
        table.append(row)  # Добавляем строку параметров в таблицу

        doc7.Close(const7.kdDoNotSaveChanges)  # Закроем файл без изменения

    if not is_run: app7.Quit()  # Закрываем программу при необходимости
    return table


def getKeyFromDict(myDict, myKey):
    return myDict[myKey] if (myKey) in myDict else ""


def print_to_excel(result):
    excel = Dispatch("Excel.Application")  # Подключаемся к программе Excel
    excel.Visible = True  # Делаем окно видимым
    wb = excel.Workbooks.Add()  # Добавляем новую книгу
    sheet = wb.ActiveSheet  # Получаем ссылку на активный лист

    # Создаём заголовок таблицы
    sheet.Range("A1:U1").value = ["Имя файла",
                                    "Разработчик",
                                    "Проверил",
                                    "Т.Контр.",
                                    "Н.Контр.",
                                    "Утвердил",
                                    "Номер изменяемого документа",
                                    "Номер извещения",
                                    "Дата выпуска",
                                    "Срок изменений",
                                    "Количество листов",
                                    "Причина изменений",
                                    "Номер причины изменений",
                                    "Указание о заделе",
                                    "Указание о внедрении",
                                    "Применяемость",
                                    "Разослать",
                                    "Приложение",
                                    "Номер изменения",
                                    "Кол-во размеров",
                                    "Кол-во пунктов ТТ",
                                    "А0", "А1", "А2", "А3", "А4", "Масштаб"]

    # Заполняем таблицу
    for i, row in enumerate(result):
        sheet.Cells(i + 2, 1).value = row['Filename']
        sheet.Cells(i + 2, 2).value = getKeyFromDict(row, 'Designer')
        sheet.Cells(i + 2, 3).value = getKeyFromDict(row, 'Checked')
        sheet.Cells(i + 2, 4).value = getKeyFromDict(row, 'TChecked')
        sheet.Cells(i + 2, 5).value = getKeyFromDict(row, 'NChecked')
        sheet.Cells(i + 2, 6).value = getKeyFromDict(row, 'Approved')
        sheet.Cells(i + 2, 7).value = getKeyFromDict(row, 'ChangedDoc')
        sheet.Cells(i + 2, 8).value = getKeyFromDict(row, 'Number')
        sheet.Cells(i + 2, 9).value = getKeyFromDict(row, 'DateCreate')
        sheet.Cells(i + 2, 10).value = getKeyFromDict(row, 'TimeChange')
        sheet.Cells(i + 2, 11).value = getKeyFromDict(row, 'PageCount')
        sheet.Cells(i + 2, 12).value = getKeyFromDict(row, 'CauseOfChange')
        sheet.Cells(i + 2, 13).value = getKeyFromDict(row, 'NumberCauseOfChange')
        sheet.Cells(i + 2, 14).value = getKeyFromDict(row, 'SamplesAvailable')
        sheet.Cells(i + 2, 15).value = getKeyFromDict(row, 'Introduction')
        sheet.Cells(i + 2, 16).value = getKeyFromDict(row, 'Applicability')
        sheet.Cells(i + 2, 17).value = getKeyFromDict(row, 'Subscribers')
        sheet.Cells(i + 2, 18).value = getKeyFromDict(row, 'Attachment')
        sheet.Cells(i + 2, 19).value = getKeyFromDict(row, 'ChangeNumber')
        sheet.Cells(i + 2, 22).value = row['A0']
        sheet.Cells(i + 2, 23).value = row['A1']
        sheet.Cells(i + 2, 24).value = row['A2']
        sheet.Cells(i + 2, 25).value = row['A3']
        sheet.Cells(i + 2, 26).value = row['A4']
        sheet.Cells(i + 2, 27).value = "".join(('="', row['Scale'], '"')) if ('Scale') in row else ""

def getFilesFromDir(dirName, listNames):

    names = os.listdir(dirName)
    for name in names:
        fullname = os.path.join(dirName, name).replace("\\", "/") # получаем полное имя
        ext = os.path.splitext(fullname)[1][1:]
        if os.path.isfile(fullname) and ext == "cdw" :
            listNames.append(fullname)
        elif os.path.isdir(fullname):
            listNames = getFilesFromDir(fullname, listNames)
    return listNames


if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Скрываем основное окно и сразу окно выбора файлов

    dirName = filedialog.askdirectory()
    print("Каталог поиска файлов " + dirName + "\n")
    listNames = []
    filenames = getFilesFromDir(dirName, listNames)

    # Исключаем файлы в каталогах old
    filenames = [filename for filename in filenames if filename.find('/old/') == -1]

    if len(filenames) != 0:
        print_to_excel(parse_design_documents(filenames))
    else:
        print("Файлы не выбраны - завершение программы")

    root.destroy()  # Уничтожаем основное окно
    root.mainloop()

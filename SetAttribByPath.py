import os
import re
import subprocess
import pythoncom
from win32com.client import Dispatch, gencache
from tkinter import Tk
# from tkinter.filedialog import askopenfilenames
from tkinter import filedialog

##-----------------------------------------------------------------------------
##
##           Программа пакетной установки атрибутов в файлах
##
##-----------------------------------------------------------------------------

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

def specWork(doc7):
    IDrawingDocument = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IDrawingDocument'], pythoncom.IID_IDispatch)


def parse_design_documents(paths):
    is_run = is_running()  # True, если программа Компас уже запущена

    module7, api7, const7 = get_kompas_api7()  # Подключаемся к программе

    app7 = api7.Application  # Получаем основной интерфейс программы

    #app5 = api5.Application
    app7.Visible = True  # Показываем окно пользователю (если скрыто)
    #app7.HideMessage = const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы

    for path in paths:
        print("Чтение файла: " + path + "\n")
        doc7 = app7.Documents.Open(PathName=path,
                                   Visible=True,
                                   ReadOnly=False)  # Откроем файл в видимом режиме без права его изменять

        stamp = doc7.LayoutSheets.Item(0).Stamp

        stamp.Text(110).Str = "Кустов"
        stamp.Text(111).Str = "Покровков"
        stamp.Text(115).Str = "Павленко"
        stamp.Update()

        app7.Visible = True  # Показываем окно пользователю (если скрыто)

        doc7.Close(const7.kdSaveChanges)  # Закроем файл с сохранением

    if not is_run: app7.Quit()  # Закрываем программу при необходимости
    return


# получение файлов из директории
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
        parse_design_documents(filenames)
    else:
        print("Файлы не выбраны - завершение программы")

    root.destroy()  # Уничтожаем основное окно
    root.mainloop()

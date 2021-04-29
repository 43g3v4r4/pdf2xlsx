import re
import os
import time
import openpyxl
import pdfplumber
import tkinter.ttk as ttk
from tkinter import *
from tkinter import messagebox


def get_text(path): # Получаем данные из PDF
    all_data = []
    all_dogovor = []
    dogovor = 1
    number_kadastr = ''

    with pdfplumber.open(path) as pdf:
        pages = pdf.pages

        for page in pdf.pages:
            text = page.extract_text()

            triger = 0

            for line in text.split('\n'):

                if triger == 1:
                    triger = 0
                    all_dogovor[-1] = all_dogovor[-1] + ' ' + line

                    all_data.append(all_dogovor)  # Добавляем в результирующий список - список с данными о договоре
                    dogovor += 1


                kadastr_nomer = re.search(r'Кадастровый номер:', line)
                rekviziti = re.search(r'реквизиты договора:', line)
                data_gos_reg = re.search(r'дата государственной регистрации:', line)
                nomer_gos_reg = re.search(r'номер государственной регистрации:', line)
                object_dole_stroit = re.search(r'объект долевого строительства:', line)


                if number_kadastr == '':
                    if kadastr_nomer: # Если найдено 'Кадастровый номер:' в строке, то записывает в переменную number_kadastr значение
                        number_kadastr = line.split('номер:')[1].split(',')[0].strip()


                if rekviziti:  # Если найдено 'реквизиты договора:' в строке
                    all_dogovor = [number_kadastr]  # Очищаем список с данными о договоре
                    all_dogovor.append(line.split('реквизиты договора:')[1].split(',')[0].strip())

                if data_gos_reg:  # Если найдено 'дата государственной регистрации:' в строке
                    all_dogovor.append(line.split('дата государственной регистрации:')[1].strip())

                if nomer_gos_reg:  # Если найдено 'номер государственной регистрации:' в строке
                    all_dogovor.append(line.split('номер государственной регистрации:')[1].strip())

                if object_dole_stroit:  # Если найдено 'объект долевого строительства:' в строке
                    all_dogovor.append(line.split('объект долевого строительства:')[1].strip())

                    if 'данные отсутствуют' in all_dogovor[-1]:
                        triger = 0
                        all_data.append(all_dogovor)  # Добавляем в результирующий список - список с данными о договоре
                        dogovor += 1
                    else:
                        if not 'кв.м' in all_dogovor[-1]:
                            triger = 1
                        else:
                            triger = 0
                            all_data.append(all_dogovor)  # Добавляем в результирующий список - список с данными о договоре
                            dogovor += 1



    return all_data


def main():
    try:
        button_start['state'] = 'disabled'
        pdf_tab.update()

        paths = []

        for root, dirs, files in os.walk(".", topdown=False):
            for name in files:
                if '.pdf' in name:
                    paths.append(os.path.join(root, name))

        print(paths)

        ot = 0
        percent = 100 / len(paths)

        result_data = []

        for path in paths:
            result_data.extend(get_text(path))
            print(result_data)

            progressbar['value'] += percent
            pdf_tab.update()

        # создаем новый excel-файл
        wb = openpyxl.Workbook()

        # добавляем новый лист
        wb.create_sheet(title='Первый лист', index=0)

        # получаем лист, с которым будем работать
        sheet = wb['Первый лист']

        sheet.column_dimensions['A'].width = 30
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 40
        sheet.column_dimensions['D'].width = 50
        sheet.column_dimensions['E'].width = 70

        sheet.append(['Кадастровый номер', 'Дата заключения договора', 'Дата государственной регистрации', 'Номер государственной регистрации', 'Объект долевого строительства'])

        for line in result_data:
            sheet.append(line)

        file_name = time.strftime("%d-%m-%Y_%H-%M-%S") + '.xlsx'
        wb.save(file_name)

        messagebox.showinfo("Info", message="Данные в Excel!")
        button_start['state'] = 'normal'
        progressbar['value'] = 0  # обнуляем progressbar

    except:
        messagebox.showinfo("Info", message="Поместите в папку PDF файлы")
        button_start['state'] = 'normal'


root = Tk()
x = (root.winfo_screenwidth() - root.winfo_reqwidth()) / 2
y = (root.winfo_screenheight() - root.winfo_reqheight()) / 2
root.geometry('270x130+{}+{}'.format(int(x)-50, int(y)-50))
root.title('pdf2xlsx v.1.0')

icon="iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAA2tpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuNi1jMTExIDc5LjE1ODMyNSwgMjAxNS8wOS8xMC0wMToxMDoyMCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0ieG1wLmRpZDo1NTQ1REREQjIwMjA2ODExOEMxNEIzQ0U0NzRCNzJEQSIgeG1wTU06RG9jdW1lbnRJRD0ieG1wLmRpZDpDN0YzRUNDRUQ3NUMxMUVBQUM5MDlBQjlFRDA4QkIwQiIgeG1wTU06SW5zdGFuY2VJRD0ieG1wLmlpZDpDN0YzRUNDREQ3NUMxMUVBQUM5MDlBQjlFRDA4QkIwQiIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgQ0MgKE1hY2ludG9zaCkiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmlpZDpkMGI0ODM3Mi1iODc5LTRhMzEtYjIwZi1mMzJjOTZhOTZlZWYiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6NTU0NUREREIyMDIwNjgxMThDMTRCM0NFNDc0QjcyREEiLz4gPC9yZGY6RGVz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"
root.call('wm', 'iconphoto', root._w, PhotoImage(data=icon))

nb = ttk.Notebook(root)
nb.pack(fill='both', expand='yes')

pdf_tab = Frame(root)

nb.add(pdf_tab, text='Главная')

button_start = Button(pdf_tab, text='Старт', width=20, height=2, command=main)
button_start.place(x=70, y=15)

log_label = Label(pdf_tab, text='Прогресс: ').place(x=5, y=60)

progressbar = ttk.Progressbar(pdf_tab, orient=HORIZONTAL, length=100, mode='determinate')
progressbar.place(x=5, y=80, width=255)

root.resizable(0,0) # запрет на изменение окна

root.mainloop()
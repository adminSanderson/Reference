import datetime
from pathlib import Path

import PySimpleGUI as sg
from docxtpl import DocxTemplate

document_path = Path(__file__).parent / "test-reference.docx"
doc = DocxTemplate(document_path)

today = datetime.datetime.today()
today_in_one_week = today + datetime.timedelta(days=7)

layout = [
    [sg.Text("Ф.И.О пациента (в Д.п.):"), sg.Input(key="CLIENT", do_not_clear=False)],
    [sg.Text("Имя файла:"), sg.Input(key="VENDOR", do_not_clear=False)],
    [sg.Text("Дата начала заболевания:"), sg.Input(key="START", do_not_clear=False)],
    [sg.Text("Выздоровление:"), sg.Input(key="FINICH", do_not_clear=False)],
    [sg.Text("Перенесла/Перенёс:"), sg.Input(key="LINE", do_not_clear=False)],
    [sg.Text("Дата начала посещения учебного заведения:"), sg.Input(key="LINE1", do_not_clear=False)],
    [sg.Text("Освобождение от занятий ФК (дней):"), sg.Input(key="LINE2", do_not_clear=False)],
    [sg.Text("Освобождение от профилактический прививок (дней):"), sg.Input(key="LINE3", do_not_clear=False)],
    [sg.Text("В контакте с больными (был(а)/не был(а)):"), sg.Input(key="LINE4", do_not_clear=False)],
    [sg.Text("Дата создание данной справки:"), sg.Input(key="LINE5", do_not_clear=False)],
    [sg.Text("Ф.И.О врача:"), sg.Input(key="PROVIDER", do_not_clear=False)],
    [sg.Button("Создать"), sg.Exit()],
]
#Дата начала посещения учебного заведения
window = sg.Window("Справка", layout, element_justification="right")

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Выход":
        break
    if event == "Создать":
        values["TODAY"] = today.strftime("%Y-%m-%d")
        values["TODAY_IN_ONE_WEEK"] = today_in_one_week.strftime("%Y-%m-%d")

        #Сохранение документа и отправка сообщение пользователю
        doc.render(values)
        output_path = Path(__file__).parent / f"{values['VENDOR']}-reference.docx"
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")

window.close()
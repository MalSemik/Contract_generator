from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
from num2words import num2words
import PySimpleGUI as sg


def fill_template(name, hiring, address1, address2, remuneration, gender, num, clause1, clause2, clause3):
    template = "ABB_template.docx"
    document = MailMerge(template)
    print(document.get_merge_fields())
    # {'Date', 'Address1', 'hiring', 'Klauzula3', 'Name', 'Address2', 'Klauzula1', 'remuneration', 'salutation', 'Klauzula2'}
    remuneration = int(remuneration)
    remuneration_word = num2words(remuneration, lang='it') + "/00"
    if gender == 'M':
        salutation = "Sig."
    else:
        salutation = "Sig.ra"
    if clause1:
        clause1 = "To jest Klauzula 1\n \n"
    else:
        clause1 = ""

    if clause2:
        clause2 = "To jest Klauzula 2\n \n"
    else:
        clause2 = ""
    if clause3:
        clause3 = f"To jest Klauzula 3 z numerkiem {num}\n \n"
    else:
        clause3 = ""

    document.merge(
        Date='{:%d.%m.%Y}'.format(date.today()),
        salutation=salutation,
        Address1=address1,
        Address2=address2,
        hiring=hiring,
        Name=name,
        remuneration=str(remuneration),
        remuneration_word=remuneration_word,
        Klauzula1=clause1,
        Klauzula2=clause2,
        Klauzula3=clause3,

    )
    document.write(f'{name}.docx')


employees = ["Jan Kowalski", "Kowal Jankowski", "Jawal Kojanski"]
sg.theme('Black')  # Add a touch of color
# All the stuff inside your window.
layout = [[sg.Text('Create a hiring letter')],
          [sg.Text('Select name and surname'), sg.Combo(employees, key='name')],
          [sg.Text('Select the gender'), sg.Radio('M', 'gender', key='male', default=True),
           sg.Radio('F', 'gender', key='female')],
          [sg.Text('Enter the hiring date: '), sg.Input(key='hiring')],
          [sg.Text('Enter the remuneration: '), sg.Input(key='remuneration')],
          [sg.Text('Enter street: '), sg.Input(key='address1')],
          [sg.Text('Enter city and city code: '), sg.Input(key='address2')],
          [sg.Text('Clause1: '), sg.Checkbox('Clause1', key='clause1')],
          [sg.Text('Clause2: '), sg.Checkbox('Clause2', key='clause2')],
          [sg.Text('Clause3: '), sg.Checkbox('Clause3', key='clause3'), sg.Text('Number: '), sg.Input(key='num')],

          [sg.Button('Ok'), sg.Button('Cancel')]]

# Create the Window
window = sg.Window('Window Title', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    try:
        event, values = window.read()
        print(event, values)
        name = values['name']
        hiring = values['hiring']
        address1 = values['address1']
        address2 = values['address2']
        remuneration = values['remuneration']
        if values['male']:
            gender = 'M'
        else:
            gender = 'F'
        num = values['num']
        clause1 = values['clause1']
        clause2 = values['clause2']
        clause3 = values['clause3']
        if event == sg.WIN_CLOSED or event == 'Cancel':  # if user closes window or clicks cancel
            break
        else:
            fill_template(name, hiring, address1, address2, remuneration, gender, num, clause1, clause2, clause3)
            sg.Popup('Lettera generata')
            break
    except Exception as e:
        sg.Popup(f'Something went wrong\n {e}')
window.close()

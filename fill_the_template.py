from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
from num2words import num2words
import PySimpleGUI as sg
import pandas as pd



def fill_template(name, hiring, address1, address2, remuneration, gender):
    template = "ABB_template.docx"
    document = MailMerge(template)
    print(document.get_merge_fields())
    remuneration = int(remuneration)
    remuneration_word = num2words(remuneration, lang='it') + "/00"
    if gender == 'M':
        salutation = "Sig."
    else:
        salutation = "Sig.ra"


    document.merge(
        Date='{:%d.%m.%Y}'.format(date.today()),
        salutation=salutation,
        Address1=address1,
        Address2=address2,
        hiring=hiring,
        Name=name,
        remuneration=str(remuneration),
        remuneration_word=remuneration_word,

    )
    document.write(f'{name}.docx')


df = pd.read_excel(r'tracker.xlsx')
employees = []
select_not_completed = df.loc[df['Completed'] == 'N']
for item in select_not_completed['Name']:
    employees.append(item)
print(employees)

sg.theme('Black')  # Add a touch of color


# All the stuff inside your window.


def create_select_employee_window():
    layout = [[sg.Text('Select an employee')],
              [sg.Text('Select name and surname'), sg.Combo(employees, key='name')],
              [sg.Text('Select the gender'), sg.Radio('M', 'gender', key='male', default=True),
               sg.Radio('F', 'gender', key='female')],
              [sg.Text('Select the type of template letter:')],
              [sg.Radio('Indeterminato', 'letter_type', key='indet', default=True)],
              [sg.Radio('Determinato', 'letter_type', key='det')],
              [sg.Radio('Determinato Dirigenti', 'letter_type', key='det_dir')],
              [sg.Radio('Indeterminato Dirigenti', 'letter_type', key='indet_dir')],
              [sg.Radio('Sostituzione maternità', 'letter_type', key='sost')],
              [sg.Button('Ok'), sg.Button('Cancel')]]
    # Create the Window
    window = sg.Window('Select an employee', layout)
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        try:
            event, values = window.read()
            print(event, values)
            # get the name
            name = values['name']

            # get the template file
            if values['indet']:
                template = 'ABB_iT_LETTERA INDETERMINATO_impiegato.docx'
            elif values['det']:
                template = 'ABB_IT_TEMPLATE LETTERA _DETERMINATO.docx'
            elif values['det_dir']:
                template = 'ABB_IT_ TEMPLATE DIRIGENTE INDETERMINATO.docx'
            elif values['indet_dir']:
                template = 'ABB_IT_ TEMPLATE DIRIGENTE INDETERMINATO.docx'
            elif values['sost']:
                template = 'ABB_IT_Template Contratto maternita\' malattia.docx'

            # get the gender
            if values['male']:
                gender = 'M'
            else:
                gender = 'F'

            if event == sg.WIN_CLOSED or event == 'Cancel':  # if user closes window or clicks cancel
                break
            else:
                window.close()
                return name, gender, template

        except Exception as e:
            sg.Popup(f'Something went wrong\n {e}')
    window.close()


def create_fill_letter_window(name,gender, template):
    layout = [[sg.Text('Create a hiring letter')],
              [sg.Text(f'Selected employee: {name}')],
              [sg.Text('Enter the hiring date: '), sg.Input(key='hiring')],
              [sg.Text('Enter the remuneration: '), sg.Input(key='remuneration')],
              [sg.Text('Enter street: '), sg.Input(key='address1')],
              [sg.Text('Enter city and city code: '), sg.Input(key='address2')],
              [sg.Text('Enter probation period (months)'), sg.Input(key='probation')],
              [sg.Checkbox('Addendum field service technician', key='service technician')],
              [sg.Checkbox('Addendum Intellectual Property', key='Intellectual_Property')],
              [sg.Checkbox('Addendum PES', key='PES')],
              [sg.Checkbox('Addendum  MBO', key='Add_MBO')],
              [sg.Checkbox('Addendum UNA TANTUM ', key='Add_UT')],
              [sg.Checkbox('Addendum UNA TANTUM Policy NEO', key='Add_UT_NEO')],
              [sg.Checkbox('Addendum Servizi relocation ', key='reloc')],
              [sg.Checkbox('Addendum passaggio categoria', key='pass_cat')],
              [sg.Checkbox('Addendum Carta Blu', key='C_Blu')],
              [sg.Checkbox('Addendum Turni Avvicendati', key='T_Avv')],
              [sg.Button('Ok'), sg.Button('Cancel')]]
    # Create the Window
    window = sg.Window('Window Title', layout)
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        try:
            event, values = window.read()
            print(event, values)
            hiring = values['hiring']
            address1 = values['address1']
            address2 = values['address2']
            remuneration = values['remuneration']

            if event == sg.WIN_CLOSED or event == 'Cancel':  # if user closes window or clicks cancel
                break
            else:
                fill_template(name, hiring, address1, address2, remuneration, gender)
                sg.Popup('Lettera generata')
                break
        except Exception as e:
            sg.Popup(f'Something went wrong\n {e}')

    window.close()


name, gender, template = create_select_employee_window()

select_ee_data = df.loc[df['Name'] == name]
df['Hiring_date'] = pd.to_datetime(df['Hiring_date']).dt.date
print(df['Hiring_date'])

create_fill_letter_window(name, gender, template)

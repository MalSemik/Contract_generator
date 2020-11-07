from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
from num2words import num2words

template = "ABB_template.docx"

document = MailMerge(template)
print(document.get_merge_fields())
# {'Date', 'Address1', 'hiring', 'Klauzula3', 'Name', 'Address2', 'Klauzula1', 'remuneration', 'salutation', 'Klauzula2'}
num = '5'
name = "Rafcio"
document.merge(
    Date='{:%d.%m.%Y}'.format(date.today()),
    salutation="Pan",
    Address1="ul. Kolorowa",
    Address2="Kanarkowo",
    hiring="12.11.2020",
    Name = name,
    #remuneration_word
    Klauzula1="To jest Klauzula 1\n \n",
    Klauzula2="To jest Klauzula 2\n \n",
    Klauzula3=f"To jest Klauzula 3 ze wstawionym numerkien{num}\n \n",

)
document.write(f'{name}.docx')
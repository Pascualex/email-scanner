import re
import os
import ctypes
import fitz
from io import StringIO
import xlsxwriter
import codecs


txt_name = 'usuarios'
txt_final_name = None
excel_name = 'tabla_usuarios'
excel_final_name = None

users = []
fields = [
    {
        'pdf_name': 'DNI',
        'txt_name': 'dni',
        'excel_name': 'DNI',
        'column_width': 20,
        'line_split': False
    },
    {
        'pdf_name': 'Nombre',
        'txt_name': 'nombre',
        'excel_name': 'Nombre',
        'column_width': 20,
        'line_split': False
    },
    {
        'pdf_name': 'Apellidos',
        'txt_name': 'apellidos',
        'excel_name': 'Apellidos',
        'column_width': 30,
        'line_split': False
    },
    {
        'pdf_name': 'Correo electrónico',
        'txt_name': 'correo',
        'excel_name': 'Correo electrónico',
        'column_width': 60,
        'line_split': False
    },
    {
        'pdf_name': 'Teléfono',
        'txt_name': 'telefono',
        'excel_name': 'Teléfono',
        'column_width': 20,
        'line_split': False
    },
    {
        'pdf_name': 'Entidad/Organización/Ayuntamiento',
        'txt_name': 'entidad',
        'excel_name': 'Entidad',
        'column_width': 50,
        'line_split': False
    },
    {
        'pdf_name': 'Puesto de trabajo',
        'txt_name': 'puesto',
        'excel_name': 'Puesto de trabajo',
        'column_width': 80,
        'line_split': False
    },
    {
        'pdf_name': 'Fecha',
        'txt_name': 'fecha',
        'excel_name': 'Fecha',
        'column_width': 40,
        'line_split': True
    }
]


def read_pdf():
    doc = fitz.open('correos.pdf')
    reset_fields()

    for i in range(doc.pageCount):
        p = doc.loadPage(i)
        text = p.getText()

        multi_line_field = None
        for line in text.splitlines():
            if multi_line_field is not None:
                for field in fields:
                    if multi_line_field == field['pdf_name']:
                        field['content'] = line
                        break
                multi_line_field = None
            else:
                parts = re.compile('[ ]*:[ ]*').split(line)
                if len(parts) == 2:
                    multi_line_field = process_field(parts[0], parts[1])

    store_fields()

    return True


def process_field(field_name, content):
    global fields
    if field_name == 'De':
        store_fields()
        reset_fields()
    else:
        for field in fields:
            if field_name == field['pdf_name']:
                if field['line_split']:
                    return field['pdf_name']
                else:
                    field['content'] = content
                    return


def reset_fields():
    global fields
    for field in fields:
        field['content'] = ''


def store_fields():
    empty = True
    for field in fields:
        if field['content'] != '':
            empty = False
            break

    if empty:
        return

    user = []
    for field in fields:
        user.append(field['content'])
    users.append(user)


def write_txt():
    global txt_final_name

    if os.path.isfile('./' + txt_name + '.txt') is True:
        file_id = 1
        while os.path.isfile('./' + txt_name + '_' + str(file_id) + '.txt') is True:
            file_id += 1
        txt_final_name = txt_name + '_' + str(file_id) + '.txt'
    else:
        txt_final_name = txt_name + '.txt'

    file = codecs.open(txt_final_name, 'w', 'utf-8')

    for i, field in enumerate(fields):
        if i < (len(fields) - 1):
            file.write(field['txt_name'] + '|')
        else:
            file.write(field['txt_name'] + '\n')

    for user in users:
        for i, field in enumerate(user):
            if i < (len(user) - 1):
                file.write(field + '|')
            else:
                file.write(field + '\n')

    file.close()


def write_excel():
    global excel_final_name

    if os.path.isfile(excel_name + '.xlsx') is True:
        file_id = 1
        while os.path.isfile(excel_name + '_' + str(file_id) + '.xlsx') is True:
            file_id += 1
        excel_final_name = excel_name + '_' + str(file_id) + '.xlsx'
    else:
        excel_final_name = excel_name + '.xlsx'
    
    workbook = xlsxwriter.Workbook(excel_final_name)
    worksheet = workbook.add_worksheet(excel_name)

    worksheet.add_table(0, 0, len(users), len(fields) - 1, {'name': excel_name})

    for i, field in enumerate(fields):
        worksheet.write(0, i, field['excel_name'])
        worksheet.set_column(i, i, field['column_width'])

    for i, user in enumerate(users):
        for j, field in enumerate(user):
            worksheet.write(i + 1, j, field)

    workbook.close()


if __name__ == '__main__':
    if read_pdf():
        write_txt()
        write_excel()
        title = 'Éxito'
        message = 'Se han procesado correctamente ' + str(len(users)) + ' usuarios. Los resultados se han almacenado en ' + txt_final_name + ' y ' + excel_final_name + '.'
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x40000)
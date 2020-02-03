import re
import os
import ctypes
import fitz
from io import StringIO
import xlsxwriter


txt_name = 'usuarios'
txt_final_name = None
excel_name = 'tabla_usuarios'
excel_final_name = None

users = []

name = ''
lastname = ''
entity = ''
email = ''
occupation = ''


def read_pdf():
    try:
        doc = fitz.open('correos.pdf')
        for i in range(doc.pageCount):
            p = doc.loadPage(i)
            text = p.getText()

            for line in text.splitlines():
                parts = re.compile('[ ]*:[ ]*').split(line)
                if len(parts) == 2:
                    process_field(parts[0], parts[1])

        store_fields()

        return True
    except:        
        title = 'Error'
        message = 'No se ha encontrado el fichero correos.pdf. Por favor, colóquelo en la misma carpeta que este ejecutable.'
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x40000)

        return False


def process_field(field, content):
    global name, lastname, entity, email, occupation
    if field == 'De':
        store_fields()
        reset_fields()
    elif field == 'Nombre':
        name = content
    elif field == 'Apellidos':
        lastname = content
    elif field == 'Entidad/Organización/Ayuntamiento':
        entity = content
    elif field == 'Correo electrónico':
        email = content
    elif field == 'Puesto de trabajo':
        occupation = content


def reset_fields():
    global name, lastname, entity, email, occupation
    name = ''
    lastname = ''
    entity = ''
    email = ''
    occupation = ''


def store_fields():
    if name == '' and lastname == '' and entity == '' and email == '' and occupation == '':
        return
    users.append({
        'name': name, 
        'lastname': lastname,
        'entity': entity, 
        'email': email,
        'occupation': occupation
    })


def write_txt():
    global txt_final_name

    if os.path.isfile('./' + txt_name + '.txt') is True:
        file_id = 1
        while os.path.isfile('./' + txt_name + '_' + str(file_id) + '.txt') is True:
            file_id += 1
        txt_final_name = txt_name + '_' + str(file_id) + '.txt'
    else:
        txt_final_name = txt_name + '.txt'
        
    file = open(txt_final_name, 'w')
    file.write('nombre|apellidos|entidad|correo|puesto\n')

    for user in users:
        file.write(user['name'] + '|' + user['lastname'] + '|' + user['entity'] + '|' + user['email'] + '|' + user['occupation'] + '\n')

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

    worksheet.add_table(0, 0, len(users), 4, {'name': excel_name})

    worksheet.write(0, 0, 'Nombre')
    worksheet.set_column(0, 0, 20)
    worksheet.write(0, 1, 'Apellidos')
    worksheet.set_column(1, 1, 30)
    worksheet.write(0, 2, 'Entidad')
    worksheet.set_column(2, 2, 60)
    worksheet.write(0, 3, 'Correo Electrónico')
    worksheet.set_column(3, 3, 60)
    worksheet.write(0, 4, 'Ocupación')
    worksheet.set_column(4, 4, 80)

    position = 1
    for user in users:
        worksheet.write(position, 0, user['name'])
        worksheet.write(position, 1, user['lastname'])
        worksheet.write(position, 2, user['entity'])
        worksheet.write(position, 3, user['email'])
        worksheet.write(position, 4, user['occupation'])
        
        position += 1

    workbook.close()


if __name__ == '__main__':
    if read_pdf():
        write_txt()
        write_excel()
        title = 'Éxito'
        message = 'Se han procesado correctamente ' + str(len(users)) + ' usuarios. Los resultados se han almacenado en ' + txt_final_name + ' y ' + excel_final_name + '.'
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x40000)
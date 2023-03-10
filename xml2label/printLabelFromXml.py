# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import pprint
import xml.etree.ElementTree as ET
import xmlrpc.client

import xlwings as xw
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML

from generateCsv import generate_panels_csv, generate_pieces_csv, generate_retals_csv


@xw.sub
def main():

    used_integers = []
    unique_pattern_list = []
    list_unique_used_board_data = []
    list_unique_used_part_data = []
    download_stack = {}
    list_unique_used_board_data_for_csv = []
    list_unique_used_part_data_for_csv = []
    list_excel_dict = []
    list_unique_used_retal_data_for_csv = []
    code = espesor = date = veta = ""

    # Odoo connection
    url_list = []
    url = 'http://mail.insca.com:8199'
    db = 'preproduccion'
    username = 'onduex'
    password = 'tomas2021'
    common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
    common.version()
    uid = common.authenticate(db, username, password, {})
    models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models.execute_kw(db, uid, password, 'product.template', 'check_access_rights', ['read'],
                      {'raise_exception': False})
    product_tmpl_ids = models.execute_kw(db, uid, password, 'product.template', 'search_read',
                                         [[['default_code', 'like', 'RT.']]])
    for rec in product_tmpl_ids:
        used_integers.append(rec['default_code'][3:])
    print(used_integers)

    # Usar pp.pprint()
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=0)

    wb = xw.Book.caller()
    sheet = wb.sheets["SQL ODOO"]
    xml_name = str(sheet["A1"].value)
    user_excel = str(sheet["C2"].value)
    list_name = str(sheet["A2"].value)[:-5]
    orden_corte = str(sheet["C1"].value)

    for i in range(6, wb.sheets["SQL ODOO"].
                   range('F' + str(wb.sheets["SQL ODOO"].cells.last_cell.row)).end('up').row + 1):
        excel_dict = ({
            'fila': i,
            'colF': sheet["F" + str(i)].value,
            'colJ': sheet["J" + str(i)].value,
            'colO': sheet["O" + str(i)].value,
            'colP': sheet["P" + str(i)].value,
            'colQ': sheet["Q" + str(i)].value,
            'colR': sheet["R" + str(i)].value,
        })
        list_excel_dict.append(excel_dict)
    # pp.pprint(list_excel_dict)

    # Fecha creación XML
    path = str('O:/XmlJob/' + xml_name + '.xml')
    tree = ET.parse(path)
    root = tree.getroot()
    for child in root:
        if child.tag == 'Solution':
            print('Número de patrones: ', child.attrib['NPatterns'])
        if child.tag == 'Tx':
            date = child.attrib['Date'][3:]
        if child.tag == 'Material':
            code = child.attrib['Code']
            espesor = child.attrib['Thickness']
            veta = child.attrib['Grain']

    # Cantidad de los tableros enteros usados
    for rec in root.iter('BrdInfo'):
        if rec.get('QUsed') != '0':
            unique_pattern = ({
                'id': rec.attrib['BrdId'],
                'QUsed': rec.attrib['QUsed'],
            })
            unique_pattern_list.append(unique_pattern)

    # Datos de los tableros enteros usados + Cantidad
    for board in root.findall('Board'):
        for rec in unique_pattern_list:
            if rec['id'] == board.get('id'):
                # Añadir a diccionarios
                unique_used_board_data = ({
                    'id': board.get('id'),
                    'L': str(board.get('L')).replace('.', ',')[:-3],
                    'W': str(board.get('W')).replace('.', ',')[:-3],
                    'BrdCode': board.get('BrdCode'),
                    'QUsed': rec['QUsed'],
                    'Qty': board.get('Qty')
                })
                unique_used_board_data_for_csv = ({
                    'ID': board.get('id'),
                    'LARGO': str(board.get('L')[:-3]),
                    'ANCHO': str(board.get('W')[:-3]),
                    'CANT': rec['QUsed'],
                    'MATERIAL': code,
                    'ESPESOR': str(board.get('Thickness')[:-3]),
                    'CATEGORIA': 'MPRIMA' + ' / ' + 'MADERA ' + code,
                    'CODIGO': board.get('BrdCode'),
                    'OC': orden_corte
                })
                list_unique_used_board_data.append(unique_used_board_data)
                list_unique_used_board_data_for_csv.append(unique_used_board_data_for_csv)

    # Cantidad de piezas cortadas
    for piece in root.iter('Piece'):
        # Añadir a diccionario
        unique_used_part_data = ({
            'id': piece.get('N'),
            'L': str(piece.get('L')).replace('.', ','),
            'W': str(piece.get('W')).replace('.', ','),
            'Qty': piece.get('Q')
        })
        for part in root.findall('Part'):
            if piece.get('N') == part.get('Code'):
                # Buscar en excel columna F el código
                key = 'colF'
                val = part.get('Desc2')
                y = next(filter(lambda x: x.get(key) == val, list_excel_dict), None)
                if y is not None:
                    # print(d)
                    unique_used_part_data.update({
                        'Op': part.get('Desc1'),
                        'BrdCode': part.get('Desc2'),
                        'Ruta': y.get('colJ'),
                        'Semana': y.get('colO'),
                    })
        list_unique_used_part_data.append(unique_used_part_data)

    # Cantidad de piezas cortadas para CSV
    for part in root.findall('Part'):
        # Añadir a diccionario
        unique_used_part_data_for_csv = ({
            'ID': part.get('id'),
            'LARGO': str(part.get('L')[:-3]),
            'ANCHO': str(part.get('W')[:-3]),
            'CODE': part.get('Code'),
            'OT': part.get('Desc1'),
            'CODCONF': part.get('Desc2'),
            'MATERIAL': part.get('Material'),
            'ESPESOR': espesor[:-3],
            'CATEGORIA': 'PSEMIELABORADO' + ' / ' + 'MADERA ' + code,
            'OC': orden_corte,
        })
        for piece in root.iter('Piece'):
            if part.get('Code') == piece.get('N'):
                unique_used_part_data_for_csv.update({
                    'CANT': piece.get('Q'),
                })
        # Buscar en excel columna R el código
        key = 'colF'
        val = part.get('Desc2')
        y = next(filter(lambda x: x.get(key) == val, list_excel_dict), None)
        if y is not None:
            unique_used_part_data_for_csv.update({
                'TIMECOEF': y.get('colR'),
            })
        list_unique_used_part_data_for_csv.append(unique_used_part_data_for_csv)

    # Cantidad de retales para CSV
    for retal in root.iter('Drop'):
        # Añadir a diccionario
        unique_used_retal_data_for_csv = ({
            'CODIGO': 'test',
            'LARGO': str(retal.get('L')[:-3]),
            'ANCHO': str(retal.get('W')[:-3]),
            'CANT': retal.get('Q'),
            'MATERIAL': code,
            'ESPESOR': espesor[:-3],
            'CATEGORIA': 'MPRIMA' + ' / ' + 'MADERA ' + code,
            'OC': orden_corte,
            'DESCRIPCION': code + ' ' +
            str(retal.get('L')[:-3]).zfill(4) + 'X' +
            str(retal.get('W')[:-3]).zfill(4) + 'X' +
            espesor[:-3].zfill(2) + 'MM',
            'VETA': veta
        })
        list_unique_used_retal_data_for_csv.append(unique_used_retal_data_for_csv)

    # Descarga de pilas
    for pattern in root.iter('Pattern'):
        download_stack.update({
            pattern.get('id'): [],
        })
    for piece in root.iter('Piece'):
        for sid in piece.iter('Sid'):
            download_stack[sid.attrib['id']].append(piece.get('N'))
    list_download_stack = [{k: v} for k, v in download_stack.items()]

    # Definición de plantilla y variables
    environment = Environment(loader=FileSystemLoader('C:/PycharmProjects/apps-insca/xml2label/templates/'))
    template = environment.get_template('informe.html')
    template_vars = {"image_path": "file:///C:/PycharmProjects/apps-insca/xml2label/static/images/LogoNegro.png",
                     "title": orden_corte,
                     "date": date,
                     "code": code,
                     "program": root.get('name'),
                     "list_name": list_name,
                     "user_excel": user_excel,
                     "boards": list_unique_used_board_data,
                     "parts": list_unique_used_part_data,
                     "list_download_stacks": list_download_stack,
                     }
    html_out = template.render(template_vars)
    filename = ('O:/PdfJob/' + orden_corte.replace('/', '-') + '.pdf')
    HTML(string=html_out).write_pdf(filename)
    # os.startfile(filename)

    # Generar CSVs
    generate_panels_csv(list_unique_used_board_data_for_csv, orden_corte)
    generate_pieces_csv(list_unique_used_part_data_for_csv, orden_corte)
    generate_retals_csv(list_unique_used_retal_data_for_csv, orden_corte)


if __name__ == "__main__":
    xw.Book("Plantilla Odoo - Optiplanning.xlsm").set_mock_caller()
    main()

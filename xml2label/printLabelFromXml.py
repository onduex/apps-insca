# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import os
import pprint
import xlwings as xw
from xlwings import Range, constants
import xml.etree.ElementTree as ET
from weasyprint import HTML
from jinja2 import Environment, FileSystemLoader
from generateCsv import GeneratePanelsCsv, GeneratePiecesCsv
from datetime import date
from datetime import datetime

@xw.sub
def main():

    UniquePattern = {}
    UniquePatternList = []
    UniqueUsedBoardData = {}
    ListUniqueUsedBoardData = []
    UniqueUsedPartData = {}
    ListUniqueUsedPartData = []
    DownloadStack = {}
    ListDownloadStack = []
    ListUniqueUsedBoardDataForCsv = []
    ListUniqueUsedPartDataForCsv = []
    ListExcelDict = []
    d = {}

    # Usar pp.pprint()
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=0)

    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    XmlName = str(sheet["A1"].value)
    UserExcel = str(sheet["C2"].value)
    ListName = str(sheet["A2"].value)[:-5]

    for i in range(6, wb.sheets[0].range('F' + str(wb.sheets[0].cells.last_cell.row)).end('up').row + 1):
        ExcelDict = ({
            'fila': i,
            'colF': sheet["F" + str(i)].value,
            'colJ': sheet["J" + str(i)].value,
            'colO': sheet["O" + str(i)].value,
            'colP': sheet["P" + str(i)].value,
            'colQ': sheet["Q" + str(i)].value,
            })
        ListExcelDict.append(ExcelDict)
    # pp.pprint(ListExcelDict)

    tree = ET.parse('O:/XmlJob/' + XmlName + '.xml')
    root = tree.getroot()
    for child in root:
        if child.tag == 'Solution':
            print('Número de patrones: ', child.attrib['NPatterns'])
        if child.tag == 'Tx':
            date = child.attrib['Date'][3:]
        if child.tag == 'Material':
            code = child.attrib['Code']
            espesor = child.attrib['Thickness']

    # Cantidad de los tableros enteros usados
    for brdinfo in child.findall('BrdInfo'):
        if brdinfo.get('QUsed') != '0':
            UniquePattern = ({
                'id': brdinfo.attrib['BrdId'],
                'QUsed': brdinfo.attrib['QUsed'],
                })
            UniquePatternList.append(UniquePattern)
    
    # Datos de los tableros enteros usados + Cantidad
    now = datetime.now()
    for board in root.findall('Board'):
        for rec in UniquePatternList:
            if rec['id'] == board.get('id'):
                # Añadir a diccionarios
                UniqueUsedBoardData = ({
                    'id': board.get('id'),
                    'L': str(board.get('L')).replace('.', ',')[:-3],
                    'W': str(board.get('W')).replace('.', ',')[:-3],
                    'BrdCode': board.get('BrdCode'),
                    'QUsed': rec['QUsed'],
                    'Qty': board.get('Qty') 
                    })
                UniqueUsedBoardDataForCsv = ({
                    'ID': board.get('id'),
                    'LARGO': str(board.get('L')[:-3]),
                    'ANCHO': str(board.get('W')[:-3]),
                    'CANT': rec['QUsed'],
                    'MATERIAL': code,
                    'ESPESOR': str(board.get('Thickness')[:-3]),
                    'CATEGORIA': 'MPRIMA' + ' / ' + 'MADERA ' + code,
                    'CODIGO': board.get('BrdCode'),
                    'OC': 'OC/' + now.strftime('%y%m%d/%H%M%S') 
                    })
                ListUniqueUsedBoardData.append(UniqueUsedBoardData)
                ListUniqueUsedBoardDataForCsv.append(UniqueUsedBoardDataForCsv)

    # Cantidad de piezas cortadas
    
    for piece in root.iter('Piece'):
        # Añadir a diccionario
        UniqueUsedPartData = ({
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
                d = next(filter(lambda d: d.get(key) == val, ListExcelDict), None)
                if d != None:
                    # print(d)
                    UniqueUsedPartData.update({
                        'Op': part.get('Desc1'),
                        'BrdCode': part.get('Desc2'),
                        'Ruta': d.get('colJ'),
                        'Semana': d.get('colO'),
                        })
        ListUniqueUsedPartData.append(UniqueUsedPartData)

    # Cantidad de piezas cortadas para CSV
    for part in root.findall('Part'):
        # Añadir a diccionario
        UniqueUsedPartDataForCsv = ({
        'ID': part.get('id'),
        'LARGO': str(part.get('L')[:-3]),
        'ANCHO': str(part.get('W')[:-3]),
        'CANT': part.get('qMin'),
        'CODE': part.get('Code'),
        'OT': part.get('Desc1'),
        'CODCONF': part.get('Desc2'),
        'MATERIAL': part.get('Material'),
        'ESPESOR': espesor[:-3],
        'CATEGORIA': 'PSEMIELABORADO' + ' / ' + 'MADERA ' + code,
        'OC': 'OC/' + now.strftime('%y%m%d/%H%M%S')
        })
        ListUniqueUsedPartDataForCsv.append(UniqueUsedPartDataForCsv)
    

    # Descarga de pilas
    for pattern in child.findall('Pattern'):
        DownloadStack.update({
            pattern.get('id'): [],
            })
    for piece in child.findall('Piece'):
        for sid in piece.iter('Sid'):
            DownloadStack[sid.attrib['id']].append(piece.get('N'))
    ListDownloadStack = [{k:v} for k, v in DownloadStack.items()]

    # Definición de plantilla y variables
    environment = Environment(loader=FileSystemLoader('C:/vs-projects/apps-insca/xml2label/templates/'))
    template = environment.get_template('informe.html')    
    template_vars = {"image_path": "file:///C:/vs-projects/apps-insca/xml2label/static/images/LogoNegro.png",
                     "title": 'OC/' + now.strftime('%y%m%d/%H%M%S'),
                     "date": date,
                     "code": code,
                     "program": root.get('name'),
                     "listname": ListName,
                     "userexcel": UserExcel,
                     "boards": ListUniqueUsedBoardData,
                     "parts": ListUniqueUsedPartData,
                     "listdownloadstacks": ListDownloadStack,
                     }
    html_out = template.render(template_vars)
    filename = ('O:/PdfJob/' + 'OC-' + now.strftime('%y%m%d') + '-' + now.strftime('%H%M%S') + '.pdf')
    HTML(string=html_out).write_pdf(filename)
    os.startfile(filename)

    # Generar CSVs
    GeneratePanelsCsv(ListUniqueUsedBoardDataForCsv, now)
    GeneratePiecesCsv(ListUniqueUsedPartDataForCsv, now)
    
if __name__ == "__main__":
    xw.Book("Plantilla Odoo - Optiplanning.xlsm").set_mock_caller()
    main()

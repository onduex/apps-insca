# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import os
import pprint
import xlwings as xw
import xml.etree.ElementTree as ET

from weasyprint import HTML
from jinja2 import Environment, FileSystemLoader


@xw.sub
def main():
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=2)
    
    UniquePattern = {}
    UniquePatternList = []
    UniqueUsedBoardData = {}
    ListUniqueUsedBoardData = []

    UniqueUsedPartData = {}
    ListUniqueUsedPartData = []

    tree = ET.parse('C:/vs-projects/apps-insca/xml2label/04756565.xml')
    root = tree.getroot()
    for child in root:
        if child.tag == 'Solution':
            print('Número de patrones: ', child.attrib['NPatterns'])
        if child.tag == 'Tx':
            date = child.attrib['Date']
        if child.tag == 'Material':
            code = child.attrib['Code']

    # Cantidad de los tableros enteros usados
    for brdinfo in child.findall('BrdInfo'):
        if brdinfo.get('QUsed') != '0':
            UniquePattern = ({
                'id': brdinfo.attrib['BrdId'],
                'QUsed': brdinfo.attrib['QUsed'],
                })
            UniquePatternList.append(UniquePattern)
    
    # Datos de los tableros enteros usados + Cantidad
    for board in root.findall('Board'):
        for rec in UniquePatternList:
            if rec['id'] == board.get('id'):
                # Añadir a diccionario
                UniqueUsedBoardData = ({
                    'id': board.get('id'),
                    'L': str(board.get('L')).replace('.', ','),
                    'W': str(board.get('W')).replace('.', ','),
                    'BrdCode': board.get('BrdCode'),
                    'QUsed': rec['QUsed'],
                    'Qty': board.get('Qty') 
                    })
                ListUniqueUsedBoardData.append(UniqueUsedBoardData)

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
                UniqueUsedPartData.update({
                    'Op': part.get('Desc1'),
                    'BrdCode': part.get('Desc2'),
                    })
        ListUniqueUsedPartData.append(UniqueUsedPartData)
    pp.pprint(ListUniqueUsedPartData)


    # Cantidad de piezas cortadas por patrón
    # for piece in root.iter('Piece'):
    #     for sid in piece.iter('Sid'):
    #         # print(piece.attrib, sid.attrib)
    #         for pattern in child.findall('Pattern'):
    #             if pattern.get('id') == sid.get('id'):
    #                 print('De la pieza', piece.get('N'), 'de', piece.get('L'), 'X',
    #                 piece.get('W'),'hay que fabricar', piece.get('Q'), 'con el patrón',
    #                 pattern.get('id'), 'y el tablero', pattern.get('BrdNo'))
    #                 # Añadir a diccionario
    #                 UniqueUsedPartData = ({
    #                 'id': piece.get('N'),
    #                 'L': str(piece.get('L')).replace('.', ','),
    #                 'W': str(piece.get('W')).replace('.', ','),
    #                 'Qty': piece.get('Q') 
    #                 })
    #             ListUniqueUsedPartData.append(UniqueUsedPartData)
    # print(ListUniqueUsedPartData)


    # Definición de plantilla y variables
    environment = Environment(loader=FileSystemLoader('C:/vs-projects/apps-insca/xml2label/templates/'))
    template = environment.get_template('informe.html')    
    template_vars = {"title" : root.get('name') ,
                     "date": date,
                     "code": code,  
                     "boards": ListUniqueUsedBoardData,
                     "parts": ListUniqueUsedPartData,
                     }
    html_out = template.render(template_vars)
    HTML(string=html_out).write_pdf('C:/vs-projects/apps-insca/xml2label//templates/report.pdf')
    os.startfile('C:/vs-projects/apps-insca/xml2label/templates/report.pdf')

    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet["A1"].value = 'Número de patrones: ', child.attrib['NPatterns']
    sheet["A2"].value = 'Patrones diferentes: ', str(UniquePattern)

if __name__ == "__main__":
    xw.Book("test.xlsm").set_mock_caller()
    main()

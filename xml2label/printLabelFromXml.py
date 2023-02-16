# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import os
import pprint
import xlwings as xw
import xml.etree.ElementTree as ET

from weasyprint import HTML
from jinja2 import Environment, FileSystemLoader

UniquePattern = {}
UniquePatternList = []
UniqueUsedBoardData = {}
ListUniqueUsedBoardData = []
UniqueUsedPartData = {}
ListUniqueUsedPartData = []
DownloadStack = {}
ListDownloadStack = []
ListPiecePerPattern = []

@xw.sub
def main():

    # Usar pp.print()
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=0)

    tree = ET.parse('C:/vs-projects/apps-insca/xml2label/03964113.xml')
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

    # Descarga de pilas
    for pattern in child.findall('Pattern'):
        DownloadStack = ({
            'Stack': pattern.get('id'),
            })
        for piece in child.findall('Piece'):
            for sid in piece.iter('Sid'):
                if pattern.get('id') == sid.attrib['id']:
                    ListPiecePerPattern.append(piece.get('N'))
                    DownloadStack.update({
                        'Pieces': ListPiecePerPattern,
                        })
        pp.pprint(DownloadStack)
        ListPiecePerPattern.clear()

    # Definición de plantilla y variables
    environment = Environment(loader=FileSystemLoader('C:/vs-projects/apps-insca/xml2label/templates/'))
    template = environment.get_template('informe.html')    
    template_vars = {"title" : root.get('name') ,
                     "date": date,
                     "code": code,  
                     "boards": ListUniqueUsedBoardData,
                     "parts": ListUniqueUsedPartData,
                     # "downloadstack": DownloadStack,
                     }
    html_out = template.render(template_vars)
    HTML(string=html_out).write_pdf('C:/vs-projects/apps-insca/xml2label//templates/report.pdf')
    os.startfile('C:/vs-projects/apps-insca/xml2label/templates/report.pdf')

    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet["O1"].value = 'Número de patrones: ', child.attrib['NPatterns']
    sheet["O2"].value = 'Patrones diferentes: ', str(UniquePattern)

if __name__ == "__main__":
    xw.Book("Plantilla Odoo - Optiplanning.xlsm").set_mock_caller()
    main()

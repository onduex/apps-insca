# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import xlwings as xw
import xml.etree.ElementTree as ET
import pprint

@xw.sub
def main():
    UniquePattern = {}
    UniquePatternList = []
    UniqueUsedBoardData = {}
    ListUniqueUsedBoardData = []

    tree = ET.parse('D:/vs-projects/apps-insca/xml2label/03964113.xml')
    root = tree.getroot()
    for child in root:
        # print(child.tag)
        if child.tag == 'Solution':
            print('Número de patrones: ', child.attrib['NPatterns'])

    # Cantidad de los tableros entero usados 
    for brdinfo in child.findall('BrdInfo'):
        if brdinfo.get('QUsed') != '0':
            UniquePattern = ({
                'id': brdinfo.attrib['BrdId'],
                'QUsed': brdinfo.attrib['QUsed'],
                })
            UniquePatternList.append(UniquePattern)
    # print(UniquePatternList)
    
    # Datos de los tableros enteros usados + Cantidad
    for board in root.findall('Board'):
        for rec in UniquePatternList:
            if rec['id'] == board.get('id'):
                UniqueUsedBoardData = ({
                    'id': board.get('id'),
                    'L': board.get('L'),
                    'W': board.get('W'),
                    'BrdCode': board.get('BrdCode'),
                    'QUsed': rec['QUsed'], 
                    })
                ListUniqueUsedBoardData.append(UniqueUsedBoardData)
    print('Tableros enteros: ', len(ListUniqueUsedBoardData))
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=4)
    pp.pprint(ListUniqueUsedBoardData)

    # Piezas en cada patrón Id
    """ for rec in root.iter('Sid'):
        print(rec.attrib) """

    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet["A1"].value = 'Número de patrones: ', child.attrib['NPatterns']
    sheet["A2"].value = 'Patrones diferentes: ', str(UniquePattern)

if __name__ == "__main__":
    xw.Book("test.xlsm").set_mock_caller()
    main()

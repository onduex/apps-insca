# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import xlwings as xw
import xml.etree.ElementTree as ET
from datetime import datetime


@xw.sub
def main():
    UniquePattern = []
    UniqueIdPattern = []

    tree = ET.parse('D:/vs-projects/apps-insca/xml2label/03964113.xml')
    root = tree.getroot()
    for child in root:
        # print(child.tag)
        if child.tag == 'Solution':
            print('Número de patrones: ', child.attrib['NPatterns'])
    for rec in root.iter('Pattern'):
        # print(rec.attrib['BrdNo'])
        if rec.attrib['BrdNo'] not in UniquePattern:
            UniquePattern.append(rec.attrib['BrdNo'])
            UniqueIdPattern.append(rec.attrib['id'])
    print('Patrones diferentes: ', UniquePattern)
    print('Patrones Id: ', UniqueIdPattern)

    # Piezas en cada patrón Id
    for rec in root.iter('Sid'):
        print(rec.attrib)

    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet["A1"].value = 'Número de patrones: ', child.attrib['NPatterns']
    sheet["A2"].value = 'Patrones diferentes: ', str(UniquePattern)

if __name__ == "__main__":
    xw.Book("test.xlsm").set_mock_caller()
    main()

# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import os
import pprint
import xml.etree.ElementTree as eT
import xlwings as xw


@xw.sub
def main():
    # Usar pp.pprint()
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=0)

    wb = xw.Book.caller()
    sheet = wb.sheets["DATA"]

    directory = str('C:/PycharmProjects/apps-insca/xml2excel-1/XmlJob')
    file_names = os.listdir(directory)
    list_unique_used_part_data = []

    for rec in file_names:
        sum_sup = 0
        seg_m2 = 0
        if rec.endswith('.xml'):
            path = str('C:/PycharmProjects/apps-insca/xml2excel-1/XmlJob/' + rec)
            print(path)
            tree = eT.parse(path)
            root = tree.getroot()
            job_time = root.find('Param').find('AddParam').get('JobTime')
            thickness = root.find('Material').get('Thickness')
            material_code = root.find('Material').get('Code')
            material_desc = root.find('Material').get('Desc')

            # Cantidad de piezas cortadas
            for part in root.iter('Part'):
                # Añadir a diccionario
                unique_used_part_data = ({
                    'id': part.get('id'),
                    'L': float(part.get('L')),
                    'W': float(part.get('W')),
                    'qMin': part.get('qMin'),
                    'Code': part.get('Code'),
                    'Thickness': thickness,
                    'Material Code': material_code,
                    'Material Desc': material_desc,
                })
                list_unique_used_part_data.append(unique_used_part_data)

            for j in range(wb.sheets[0].range('E' + str(wb.sheets[0].cells.last_cell.row)).end('up').row - 1,
                           len(list_unique_used_part_data)):
                if list_unique_used_part_data[j]['Code']:
                    sum_sup += float(list_unique_used_part_data[j]['Code'])

        # Escribir en excel
        for i in range(wb.sheets[0].range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row - 1,
                       len(list_unique_used_part_data)):
            # print(i)
            sheet["A" + str(i + 2)].value = list_unique_used_part_data[i]['id']
            sheet["B" + str(i + 2)].value = list_unique_used_part_data[i]['L']
            sheet["C" + str(i + 2)].value = list_unique_used_part_data[i]['W']
            sheet["D" + str(i + 2)].value = list_unique_used_part_data[i]['qMin']
            sheet["E" + str(i + 2)].value = list_unique_used_part_data[i]['Code']
            sheet["F" + str(i + 2)].value = list_unique_used_part_data[i]['Thickness']
            sheet["G" + str(i + 2)].value = list_unique_used_part_data[i]['Material Code']
            sheet["H" + str(i + 2)].value = list_unique_used_part_data[i]['Material Desc']


if __name__ == "__main__":
    xw.Book("Dimensiones-Piezas.xlsx").set_mock_caller()
    main()

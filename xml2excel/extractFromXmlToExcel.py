# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import pprint
import xml.etree.ElementTree as eT
import xlwings as xw


@xw.sub
def main():
    list_unique_used_part_data = []

    # Usar pp.pprint()
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=0)

    wb = xw.Book.caller()
    sheet = wb.sheets["DATA"]
    xml_name = str("00530293")
    path = str('O:/XmlToRead/' + xml_name + '.xml')
    tree = eT.parse(path)
    root = tree.getroot()
    job_time = root.find('Param').find('AddParam').get('JobTime')
    sum_sup = 0.0
    seg_m2 = 0.0

    # Cantidad de piezas cortadas
    for part in root.iter('Part'):
        # Añadir a diccionario
        sup = (float(part.get('L')) / 1000 * float(part.get('W')) / 1000) * float(part.get('qMin'))
        unique_used_part_data = ({
            'id': part.get('id'),
            'L': float(part.get('L')),
            'W': float(part.get('W')),
            'qMin': part.get('qMin'),
            'Code': part.get('Code'),
            'MatNo': part.get('MatNo'),
            'Desc1': part.get('Desc1'),
            'Desc2': part.get('Desc2'),
            'Material': part.get('Material'),
            'Sup': float(sup),
            'JobTime': job_time,
        })
        list_unique_used_part_data.append(unique_used_part_data)
    pp.pprint(len(list_unique_used_part_data))

    for i in range(len(list_unique_used_part_data)):
        sum_sup += list_unique_used_part_data[i]['Sup']
        seg_m2 = float(list_unique_used_part_data[i]['JobTime']) / float(sum_sup)

    # Escribir en excel
    for i in range(len(list_unique_used_part_data)):
        if i == 0:
            sheet["A" + str(i + 2)].value = list_unique_used_part_data[i]['id']
            sheet["B" + str(i + 2)].value = list_unique_used_part_data[i]['L']
            sheet["C" + str(i + 2)].value = list_unique_used_part_data[i]['W']
            sheet["D" + str(i + 2)].value = list_unique_used_part_data[i]['qMin']
            sheet["E" + str(i + 2)].value = list_unique_used_part_data[i]['Code']
            sheet["F" + str(i + 2)].value = list_unique_used_part_data[i]['MatNo']
            sheet["G" + str(i + 2)].value = list_unique_used_part_data[i]['Desc1']
            sheet["H" + str(i + 2)].value = list_unique_used_part_data[i]['Desc2']
            sheet["I" + str(i + 2)].value = list_unique_used_part_data[i]['Material']
            sheet["J" + str(i + 2)].value = list_unique_used_part_data[i]['Sup']
            sheet["K" + str(i + 2)].value = list_unique_used_part_data[i]['JobTime']
            sheet["L" + str(i + 2)].value = seg_m2
        else:
            sheet["A" + str(i + 2)].value = list_unique_used_part_data[i]['id']
            sheet["B" + str(i + 2)].value = list_unique_used_part_data[i]['L']
            sheet["C" + str(i + 2)].value = list_unique_used_part_data[i]['W']
            sheet["D" + str(i + 2)].value = list_unique_used_part_data[i]['qMin']
            sheet["E" + str(i + 2)].value = list_unique_used_part_data[i]['Code']
            sheet["F" + str(i + 2)].value = list_unique_used_part_data[i]['MatNo']
            sheet["G" + str(i + 2)].value = list_unique_used_part_data[i]['Desc1']
            sheet["H" + str(i + 2)].value = list_unique_used_part_data[i]['Desc2']
            sheet["I" + str(i + 2)].value = list_unique_used_part_data[i]['Material']
            sheet["J" + str(i + 2)].value = list_unique_used_part_data[i]['Sup']


if __name__ == "__main__":
    xw.Book("Tiempos seccionadora.xlsx").set_mock_caller()
    main()

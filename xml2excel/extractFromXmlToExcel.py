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

    directory = str('O:/XmlToRead')
    file_names = os.listdir(directory)
    list_unique_used_part_data = []

    for rec in file_names:
        sum_sup = 0
        seg_m2 = 0
        if rec.endswith('.xml'):
            path = str('O:/XmlToRead/' + rec)
            print(path)
            tree = eT.parse(path)
            root = tree.getroot()
            job_time = root.find('Param').find('AddParam').get('JobTime')

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
                    'JobTime': '',
                    'seg_m2': '',
                })
                list_unique_used_part_data.append(unique_used_part_data)

            for j in range(wb.sheets[0].range('J' + str(wb.sheets[0].cells.last_cell.row)).end('up').row - 1,
                           len(list_unique_used_part_data)):
                if list_unique_used_part_data[j]['Sup']:
                    sum_sup += float(list_unique_used_part_data[j]['Sup'])

            if job_time != '0':
                seg_m2 = sum_sup / float(job_time)

            list_unique_used_part_data.append({
                    'id': '--',
                    'L': '--',
                    'W': '--',
                    'qMin': '--',
                    'Code': '--',
                    'MatNo': '--',
                    'Desc1': '--',
                    'Desc2': '--',
                    'Material': '--',
                    'Sup': sum_sup,
                    'JobTime': job_time,
                    'seg_m2': seg_m2,
                })

            list_unique_used_part_data.append({
                'id': '',
                'L': '',
                'W': '',
                'qMin': '',
                'Code': '',
                'MatNo': '',
                'Desc1': '',
                'Desc2': '',
                'Material': '',
                'Sup': '',
                'JobTime': '',
                'seg_m2': '',
            })

            # pp.pprint(len(list_unique_used_part_data))

        # Escribir en excel
        for i in range(wb.sheets[0].range('A' + str(wb.sheets[0].cells.last_cell.row)).end('up').row - 1,
                       len(list_unique_used_part_data)):
            print(i)
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
            sheet["L" + str(i + 2)].value = list_unique_used_part_data[i]['seg_m2']


if __name__ == "__main__":
    xw.Book("Tiempos seccionadora.xlsx").set_mock_caller()
    main()

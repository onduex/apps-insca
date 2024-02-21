# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import os
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

    # Cantidad de piezas cortadas
    for part in root.iter('Part'):
        # Añadir a diccionario
        unique_used_part_data = ({
            'id': part.get('id'),
            'L': str(part.get('L')).replace('.', ','),
            'W': str(part.get('W')).replace('.', ','),
            'qMin': part.get('qMin'),
            'Code': part.get('Code'),
            'MatNo': part.get('MatNo'),
            'Desc1': part.get('Desc1'),
            'Desc2': part.get('Desc2'),
            'Material': part.get('Material'),
        })
        list_unique_used_part_data.append(unique_used_part_data)
    pp.pprint(list_unique_used_part_data)


if __name__ == "__main__":
    xw.Book("Tiempos seccionadora.xlsx").set_mock_caller()
    main()

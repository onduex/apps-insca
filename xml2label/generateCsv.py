# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import csv
import pprint


def generate_panels_csv(list_unique_used_board_data_for_csv, orden_corte):
    # Usar pp.pprint()
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=0)

    # field names 
    fields = ['ID', 'LARGO', 'ANCHO', 'CANT', 'MATERIAL', 'ESPESOR', 'CATEGORIA', 'CODIGO', 'OC']

    # name of csv file 
    filename = 'O:/CSV/paneles/' + 'Paneles-' + orden_corte.replace('/', '-') + '.csv'

    # writing to csv file 
    with open(filename, 'w', newline='') as csvfile:
        # creating a csv dict writer object 
        writer = csv.DictWriter(csvfile, delimiter=';', dialect='excel', fieldnames=fields)
        writer.writeheader()
        writer.writerows(list_unique_used_board_data_for_csv)


def generate_pieces_csv(list_unique_used_part_data_for_csv, orden_corte):
    # field names 
    fields = ['ID', 'LARGO', 'ANCHO', 'CANT', 'CODE', 'OT', 'CODCONF',
              'MATERIAL', 'ESPESOR', 'CATEGORIA', 'OC', 'TIMECOEF']
    # name of csv file 
    filename = 'O:/CSV/piezas/' + 'Piezas-' + orden_corte.replace('/', '-') + '.csv'
    # writing to csv file 
    with open(filename, 'w', newline='') as csvfile:
        # creating a csv dict writer object 
        writer = csv.DictWriter(csvfile, delimiter=';', dialect='excel', fieldnames=fields)
        writer.writeheader()
        writer.writerows(list_unique_used_part_data_for_csv)


def generate_retals_csv(list_unique_used_retal_data_for_csv, orden_corte):
    # field names
    fields = ['CODIGO', 'LARGO', 'ANCHO', 'CANT', 'MATERIAL', 'ESPESOR', 'CATEGORIA', 'OC', 'DESCRIPCION', 'VETA']
    # name of csv file
    filename = 'O:/CSV/retales/' + 'Retales-' + orden_corte.replace('/', '-') + '.csv'
    # writing to csv file
    with open(filename, 'w', newline='') as csvfile:
        # creating a csv dict writer object
        writer = csv.DictWriter(csvfile, delimiter=';', dialect='excel', fieldnames=fields)
        writer.writeheader()
        writer.writerows(list_unique_used_retal_data_for_csv)

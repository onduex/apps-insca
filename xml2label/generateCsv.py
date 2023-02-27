# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import csv
import pprint

def GeneratePanelsCsv(ListUniqueUsedBoardDataForCsv, T_stamp2):

    # Usar pp.pprint()
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=0)

    # field names 
    fields = ['ID', 'LARGO', 'ANCHO', 'CANT', 'MATERIAL', 'ESPESOR', 'CATEGORIA', 'CODIGO', 'OC']

    # name of csv file 
    filename = 'O:/CSV/paneles/' + 'Paneles-' + T_stamp2 + '.csv'

    # writing to csv file 
    with open(filename, 'w', newline='') as csvfile:
        # creating a csv dict writer object 
        writer = csv.DictWriter(csvfile, delimiter =';', dialect='excel', fieldnames=fields)  
        writer.writeheader() 
        writer.writerows(ListUniqueUsedBoardDataForCsv)


def GeneratePiecesCsv(ListUniqueUsedPartDataForCsv, T_stamp2):

    # Usar pp.pprint()
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=0)

    # field names 
    fields = ['ID', 'LARGO', 'ANCHO', 'CANT', 'CODE', 'OT', 'CODCONF', 'MATERIAL', 'ESPESOR', 'CATEGORIA', 'OC']

    # name of csv file 
    filename = 'O:/CSV/piezas/' + 'Piezas-' + T_stamp2 + '.csv'

    # writing to csv file 
    with open(filename, 'w', newline='') as csvfile:
        # creating a csv dict writer object 
        writer = csv.DictWriter(csvfile, delimiter =';', dialect='excel', fieldnames=fields)  
        writer.writeheader() 
        writer.writerows(ListUniqueUsedPartDataForCsv) 


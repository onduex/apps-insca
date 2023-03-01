# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import csv
import pprint

def GeneratePanelsCsv(ListUniqueUsedBoardDataForCsv, OrdenCorte):

    # Usar pp.pprint()
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=0)

    # field names 
    fields = ['ID', 'LARGO', 'ANCHO', 'CANT', 'MATERIAL', 'ESPESOR', 'CATEGORIA', 'CODIGO', 'OC']

    # name of csv file 
    filename = 'O:/CSV/paneles/' + 'Paneles-' + OrdenCorte.replace('/', '-') + '.csv'

    # writing to csv file 
    with open(filename, 'w', newline='') as csvfile:
        # creating a csv dict writer object 
        writer = csv.DictWriter(csvfile, delimiter =';', dialect='excel', fieldnames=fields)  
        writer.writeheader() 
        writer.writerows(ListUniqueUsedBoardDataForCsv)


def GeneratePiecesCsv(ListUniqueUsedPartDataForCsv, OrdenCorte):

    # Usar pp.pprint()
    pp = pprint.PrettyPrinter(sort_dicts=False, indent=0)

    # field names 
    fields = ['ID', 'LARGO', 'ANCHO', 'CANT', 'CODE', 'OT', 'CODCONF', 'MATERIAL', 'ESPESOR', 'CATEGORIA', 'OC']

    # name of csv file 
    filename = 'O:/CSV/piezas/' + 'Piezas-' + OrdenCorte.replace('/', '-') + '.csv'

    # writing to csv file 
    with open(filename, 'w', newline='') as csvfile:
        # creating a csv dict writer object 
        writer = csv.DictWriter(csvfile, delimiter =';', dialect='excel', fieldnames=fields)  
        writer.writeheader() 
        writer.writerows(ListUniqueUsedPartDataForCsv) 


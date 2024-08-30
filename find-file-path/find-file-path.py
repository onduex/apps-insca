# -*- coding: utf-8 -*-
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl).

import os
import pprint
import xml.etree.ElementTree as eT
import xlwings as xw


@xw.sub
def main():
    def list_all_files(path):
        lista_de_archivos = []
        for defpath, surnames, filenames in os.walk(path):
            for filename in filenames:
                lista_de_archivos.append(os.path.join(defpath, filename))
        return lista_de_archivos

    ruta = "R:/dtecnic/PLANOS/0_PNG"
    archivos = list_all_files(ruta)
    print(archivos)


if __name__ == "__main__":
    xw.Book("find-file-path.xlsx").set_mock_caller()
    main()

# my_sub()
#     RunPython("import find-file-path; find-file-path.main()")
# End


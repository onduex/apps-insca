# Script para leer los archivos xml de la seccionadora y colocar los campos en una excel

INSTALACIÓN:

1.- Instalar Python* en C:\Python*

2.- cmd ejecutar como administrador

    cd C:\Python*\Scripts

    pip3 install xlwings pywin32 lxml

    pip3 --proxy HTTP://pperis:Amima56seE@172.31.30.254:3128 install xlwings pywin32 lxml

    xlwings addin install

3.- Guardar extractFromXmlToExcel.py en la misma carpeta que el archivo excel xlsx

4.- En VBA, verificar que en Herramientas/Referencias está marcado xlwings y que está la macro que llama al py

    Interprete --> C:\Python311\python.exe
    PYTHONPATH --> C:\PycharmProjects\apps-insca\xml2excel

5.- Reiniciar Excel
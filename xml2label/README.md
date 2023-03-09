# Script para imprimir en pdf la información del xml que extrae la seccionadora

INSTALACIÓN:

1.- Instalar Python* en C:\Python*

2.- cmd ejecutar como administrador

    cd C:\Python*\Scripts ó cd C:\Program Files\Python*\Scripts

    pip3 install xlwings pywin32 lxml jinja2 weasyprint requests

    xlwings addin install

Importante orden instalación siguientes programas

    C:\\PycharmProjects\apps-insca\docs\msys2-x86_64-20230127.exe

    C:\\PycharmProjects\apps-insca\docs\gtk3-runtime-3.24.31-2022-01-04-ts-win64.exe

3.- Guardar printLabelFromXml.py en la misma carpeta que el archivo excel xlsm (con macros)

4.- En VBA, verificar que en Herramientas/Referencias está marcado xlwings y que está la macro que llama al py

5.- Reiniciar Excel

NOTAS:

1.- El archivo python (py) y la excel a modificar deben estar en el mismo directorio

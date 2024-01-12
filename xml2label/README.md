# Script para imprimir en pdf la información del xml que extrae la seccionadora

INSTALACIÓN:

1.- Instalar Python* en C:\Python*

2.- cmd ejecutar como administrador

    cd C:\Python*\Scripts ó cd C:\Program Files\Python*\Scripts

    pip3 install xlwings pywin32 lxml jinja2 weasyprint requests

    pip3 --proxy HTTP://pperis:Amima56seE@172.31.30.254:3128 install xlwings pywin32 lxml jinja2 weasyprint requests

    xlwings addin install

    C:\PycharmProjects\apps-insca\soft\msys2-x86_64-20230127.exe

    C:\PycharmProjects\apps-insca\soft\gtk3-runtime-3.24.31-2022-01-04-ts-win64.exe

3.- Guardar printLabelFromXml.py en la misma carpeta que el archivo excel xlsm (con macros)

4.- En VBA, verificar que en Herramientas/Referencias está marcado xlwings y que está la macro que llama al py
    Interprete --> C:\Python311\python.exe
    PYTHONPATH --> C:\PycharmProjects\apps-insca\xml2label
    Install Fonts
    Reiniciar

5.- Reiniciar Excel

NOTAS:

1.- El archivo python (py) y la excel a modificar deben estar en el mismo directorio

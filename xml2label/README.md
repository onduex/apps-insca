# Script para imprimir en pdf la información del xml que extrae la seccionadora

INSTALACIÓN:

1.- Instalar Python* en C:\Python*

2.- cmd ejecutar como administrador

    cd C:\Python*\Scripts cd C:\Program Files\Python*\Scripts

    pip3 install xlwings pywin32

    xlwings addin install

3.- Guardar printLabelFromXml.py en la misma carpeta que el archivo excel xlsm (con macros)

4.- En VBA, verificar que en Herramientas/Referencias está marcado xlwings y que está la macro que llama al py

5.- Reiniciar Excel

NOTAS:

1.- El archivo python (py) y la excel a modificar deben estar en el mismo directorio

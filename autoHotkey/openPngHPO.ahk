#Requires AutoHotkey v2.0
#WinActivateForce

openPng() ; Función para abrir el fichero PNG
{
    A_Clipboard := ""
    Send "^c"
    ClipWait
    rutaFija := "R:\DTECNIC\PLANOS\0_PNG\"
	; http://mercury/0_PNG/A10/A10.201/A10.201029.AAB000_R0A.png
	httpLength := StrLen(A_Clipboard)
    fileName := SubStr(A_Clipboard, 22, httpLength)
    OutputDebug "fileName: " fileName

    rutaFinal := rutaFija fileName
	
	OutputDebug "rutaFinal: " rutaFinal

    try
    {
        Run rutaFinal ; Ejecutar la ruta final
        WinWait "FastStone Image Viewer" ; Esperar a que aparezca la ventana "FastStone Image Viewer
        WinActivate "FastStone Image Viewer" ; Activar la ventana de FastStone Image Viewer
        Sleep 200 ; Esperar 0.2 segundos
        WinActivate "HPO" ; Activar la ventana de HPO
    }
    catch as e  ; Si no existe el fichero PNG
    {
        MsgBox "El fichero de imagen PNG no existe"
        WinActivate "HPO" ; Activar la ventana de HPO
        Exit
    }
}


^~LButton::
{
    if WinActive("HPO")
        {
            openPng()
        }
return
}

Down::
{
    if WinActive("HPO") and WinExist("FastStone Image Viewer")
        {
            Send "{Down}"
            openPng()
        }
    else
        {
            Send "{Down}"
        }
return
}

Up::
{
    if WinActive("HPO") and WinExist("FastStone Image Viewer")
        {
            Send "{Up}"
            openPng()
        }
    else
        {
            Send "{Up}"
        }
return
}


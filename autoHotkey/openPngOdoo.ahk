#Requires AutoHotkey v2.0
#WinActivateForce

^~LButton::
{
    A_Clipboard := ""
    Click "Right" ; Click derecho
    {
        Send "Down"
        sleep 100
        Send "Down"
        sleep 100
        Send "Down"
        sleep 100
        Send "Down"
        sleep 100
        Send "Down"
        sleep 100
        Send "Down"
    }
    OutputDebug A_Clipboard
    try
        {
            Run A_Clipboard ; Ejecutar la ruta final
            WinWait "FastStone Image Viewer" ; Esperar a que aparezca la ventana "FastStone Image Viewer
            WinActivate "FastStone Image Viewer" ; Activar la ventana de FastStone Image Viewer
        }
        catch as e  ; Si no existe el fichero PNG
        {
            MsgBox "El fichero de imagen PNG o la carpeta no existen"
            Exit
        }
    return
}

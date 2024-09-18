#Requires AutoHotkey v2.0
#WinActivateForce

openPng() ; Función para abrir el fichero PNG
{
    A_Clipboard := ""
    Send "^c"
    ClipWait
    vText := Trim(A_Clipboard, "`t") ; Eliminar los TABs del principio del texto
    wordAarray := StrSplit(vText, A_Tab, ".")
    itemNumber := wordAarray[1]
    itemRevision := wordAarray[2]
    rutaFija := "R:\DTECNIC\PLANOS\0_PNG\"
    revLength := StrLen(itemRevision)
    itemNumberLength := StrLen(itemNumber)
    dirTres := SubStr(itemNumber, 1, 3)
    dirSiete := SubStr(itemNumber, 1, 7)
    ; OutputDebug "ItemNumber: " itemNumber

    if (SubStr(itemNumber, -3) = "ipt" or SubStr(itemNumber, -3) = "IPT" or
    SubStr(itemNumber, -3) = "iam" or SubStr(itemNumber, -3) = "IAM" or
    SubStr(itemNumber, -3) = "idw" or SubStr(itemNumber, -3) = "IDW" or
    Substr(itemNumber, -3) = "ipn" or Substr(itemNumber, -3) = "IPN")
        {
            itemNumber := SubStr(itemNumber, 1, itemNumberLength - 4)
        }

    if (revLength = 1)
        {
            rutaFinal := rutaFija dirTres "\" dirSiete "\" itemNumber "_R0" itemRevision ".png"
        }
    else
        {
            rutaFinal := rutaFija dirTres "\" dirSiete "\" itemNumber "_R" itemRevision ".png"
        }

    try
    {
        Run rutaFinal ; Ejecutar la ruta final
        WinWait "FastStone Image Viewer" ; Esperar a que aparezca la ventana "FastStone Image Viewer
        WinActivate "FastStone Image Viewer" ; Activar la ventana de FastStone Image Viewer
        Sleep 400 ; Esperar 0.4 segundos
        WinActivate "Autodesk Vault Professional" ; Activar la ventana de Vault
    }
    catch as e  ; Si no existe el fichero PNG
    {
        MsgBox "El fichero de imagen PNG no existe"
        WinActivate "Autodesk Vault Professional" ; Activar la ventana de Vault
        Exit
    }

}

Down::
{
    if WinActive("Autodesk Vault Professional") and WinExist("FastStone Image Viewer")
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
    if WinActive("Autodesk Vault Professional") and WinExist("FastStone Image Viewer")
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

~LButton & ~q::
{
    if WinActive("Autodesk Vault Professional")
        {
            openPng()
            
        }
return
}

^ñ::
{
    if WinActive("Autodesk Vault Professional")
        {
            openPng()
        }
return
}

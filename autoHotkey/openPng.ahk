#Requires AutoHotkey v2.0
#WinActivateForce

Down::
{
    if WinActive("Autodesk Vault Professional") and WinExist("FastStone Image Viewer")
        {
            Send "{Down}"
            A_Clipboard := ""
            Send "^c"
            ClipWait
            wordAarray := StrSplit(A_Clipboard, A_Tab, ".")
            itemNumber := wordAarray[1]
            itemRevision := wordAarray[2]
            rutaFija := "R:\DTECNIC\PLANOS\0_PNG\"
            revLength := StrLen(itemRevision)
            dirTres := SubStr(itemNumber, 1, 3)
            dirSiete := SubStr(itemNumber, 1, 7)

            If (revLength = 1)
                {
                    rutaFinal := rutaFija dirTres "\" dirSiete "\" itemNumber "_R0" itemRevision ".png"
                }
            Else
                {
                    rutaFinal := rutaFija dirTres "\" dirSiete "\" itemNumber "_R" itemRevision ".png"
                }

            try
            {
                Run rutaFinal ; Ejecutar la ruta final
                WinWait "FastStone Image Viewer" ; Esperar a que aparezca la ventana "FastStone Image Viewer 7.8
                WinActivate "FastStone Image Viewer" ; Activar la ventana de FastStone Image Viewer 7.8
                Sleep 200
                WinActivate "Autodesk Vault Professional" ; Activar la ventana de Vault
            }
            catch as e  ; Si no existe el fichero PNG
            {
                MsgBox "El fichero de imagen PNG no existe"
                Sleep 200
                WinActivate "Autodesk Vault Professional" ; Activar la ventana de Vault
                Exit
            }

        }
    else
        {
            Send "{Down}"
        }
Return
}

Up::
{
    if WinActive("Autodesk Vault Professional") and WinExist("FastStone Image Viewer")
        {
            Send "{Up}"
            A_Clipboard := ""
            Send "^c"
            ClipWait
            wordAarray := StrSplit(A_Clipboard, A_Tab, ".")
            itemNumber := wordAarray[1]
            itemRevision := wordAarray[2]
            rutaFija := "R:\DTECNIC\PLANOS\0_PNG\"
            revLength := StrLen(itemRevision)
            dirTres := SubStr(itemNumber, 1, 3)
            dirSiete := SubStr(itemNumber, 1, 7)

            If (revLength = 1)
                {
                    rutaFinal := rutaFija dirTres "\" dirSiete "\" itemNumber "_R0" itemRevision ".png"
                }
            Else
                {
                    rutaFinal := rutaFija dirTres "\" dirSiete "\" itemNumber "_R" itemRevision ".png"
                }

            try
            {
                Run rutaFinal ; Ejecutar la ruta final
                WinWait "FastStone Image Viewer" ; Esperar a que aparezca la ventana "FastStone Image Viewer 7.8
                WinActivate "FastStone Image Viewer" ; Activar la ventana de FastStone Image Viewer 7.8
                Sleep 200
                WinActivate "Autodesk Vault Professional" ; Activar la ventana de Vault
            }
            catch as e  ; Si no existe el fichero PNG
            {
                MsgBox "El fichero de imagen PNG no existe"
                Sleep 200
                WinActivate "Autodesk Vault Professional" ; Activar la ventana de Vault
                Exit
            }

        }
    else
        {
            Send "{Up}"
        }
Return
}

; PID := ProcessExist("OpenConsole.exe")
; OutputDebug "ItemNumber: " itemNumber
; OutputDebug "ItemRevision: " itemRevision
; OutputDebug "RutaFija: " rutaFija
; OutputDebug "RevLength: " revLength
; OutputDebug "dirTres: " dirTres
; OutputDebug "dirSiete: " dirSiete
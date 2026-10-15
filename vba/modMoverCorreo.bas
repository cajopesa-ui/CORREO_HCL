Attribute VB_Name = "modMoverCorreo"
'==============================================================================
' MÓDULO: modMoverCorreo
'------------------------------------------------------------------------------
' Trabaja con los correos que aparecen en la hoja "Bandeja":
'
'   MoverCorreos     -> mueve cada correo a la carpeta elegida en "Mover a"
'   CopiarCorreos    -> igual, pero el correo SIGUE también en su carpeta
'   MarcarLeidos     -> marca como leídos los correos de las filas seleccionadas
'   MarcarNoLeidos   -> marca como NO leídos los de las filas seleccionadas
'
' CÓMO SE USA "MOVER":
'   1. Lee la bandeja, una carpeta o haz una búsqueda.
'   2. En la columna "Mover a" elige la carpeta de destino con el desplegable
'      (solo en las filas que quieras mover).
'   3. Pulsa "Mover a carpeta". El resultado aparece en la columna "Resultado".
'
' En Notes, "mover" es: meter el correo en la carpeta de destino y sacarlo de
' la carpeta en la que estaba (columna "Carpeta"). Un mismo correo puede estar
' en varias carpetas a la vez; "copiar" hace justo eso.
'==============================================================================
Option Explicit


'------------------------------------------------------------------------------
' MoverCorreos / CopiarCorreos  -> MACROS (botones)
'------------------------------------------------------------------------------
Public Sub MoverCorreos()
    ProcesarDestinos True
End Sub

Public Sub CopiarCorreos()
    ProcesarDestinos False
End Sub


'------------------------------------------------------------------------------
' MarcarLeidos / MarcarNoLeidos  -> MACROS (botones)
'------------------------------------------------------------------------------
Public Sub MarcarLeidos()
    MarcarSeleccion True
End Sub

Public Sub MarcarNoLeidos()
    MarcarSeleccion False
End Sub


'------------------------------------------------------------------------------
' ProcesarDestinos: mueve (mover = True) o copia (mover = False) los correos
' que tienen algo elegido en la columna "Mover a".
'------------------------------------------------------------------------------
Private Sub ProcesarDestinos(ByVal mover As Boolean)
    Dim ws As Worksheet
    Dim db As Object
    Dim ultimaFila As Long
    Dim fila As Long
    Dim pendientes As Long
    Dim correctos As Long
    Dim fallidos As Long
    Dim destino As String
    Dim mensaje As String
    Dim accion As String

    accion = IIf(mover, "Mover", "Copiar")

    On Error GoTo ManejoError

    Set ws = PrepararHojaBandeja()
    ultimaFila = ws.Cells(ws.Rows.Count, COL_B_ID).End(xlUp).Row

    ' 1) ¿Cuántas filas tienen carpeta de destino?
    For fila = 2 To ultimaFila
        If FilaConDestino(ws, fila) Then pendientes = pendientes + 1
    Next fila

    If pendientes = 0 Then
        MsgBox "No hay ningún correo con carpeta de destino." & vbCrLf & vbCrLf & _
               "En la hoja '" & HOJA_BANDEJA & "', columna 'Mover a', elige la carpeta " & _
               "con el desplegable en las filas que quieras " & LCase$(accion) & ".", _
               vbInformation, accion & " correos"
        Exit Sub
    End If

    If MsgBox("Se van a " & LCase$(accion) & " " & pendientes & " correo(s). ¿Continuar?", _
              vbQuestion + vbYesNo, accion & " correos") = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.StatusBar = "Conectando con HCL Notes..."
    Set db = AbrirBaseCorreo()

    ' 2) Fila a fila
    For fila = 2 To ultimaFila
        If FilaConDestino(ws, fila) Then
            Application.StatusBar = accion & ": fila " & fila & "..."
            destino = Trim$(CStr(ws.Cells(fila, COL_B_MOVER).Value))

            If MoverUnCorreo(db, CStr(ws.Cells(fila, COL_B_ID).Value), _
                             Trim$(CStr(ws.Cells(fila, COL_B_CARPETA).Value)), _
                             destino, mover, mensaje) Then
                correctos = correctos + 1
                ws.Cells(fila, COL_B_RESULTADO).Font.Color = RGB(0, 128, 0)
                ' Al mover, el correo ya está en su nueva carpeta
                If mover Then ws.Cells(fila, COL_B_CARPETA).Value = TextoSeguro(destino)
                ws.Cells(fila, COL_B_MOVER).ClearContents
            Else
                fallidos = fallidos + 1
                ws.Cells(fila, COL_B_RESULTADO).Font.Color = RGB(192, 0, 0)
            End If
            ws.Cells(fila, COL_B_RESULTADO).Value = mensaje
        End If
    Next fila

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Hecho." & vbCrLf & vbCrLf & _
           "Correctos: " & correctos & vbCrLf & _
           "Con error: " & fallidos & IIf(fallidos > 0, "  (mira la columna 'Resultado')", ""), _
           IIf(fallidos > 0, vbExclamation, vbInformation), accion & " correos"

Salir:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError accion & " correos", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' FilaConDestino: True si la fila tiene un correo y una carpeta en "Mover a".
'------------------------------------------------------------------------------
Private Function FilaConDestino(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    FilaConDestino = Len(Trim$(CStr(ws.Cells(fila, COL_B_ID).Value))) > 0 And _
                     Len(Trim$(CStr(ws.Cells(fila, COL_B_MOVER).Value))) > 0
End Function


'------------------------------------------------------------------------------
' MoverUnCorreo: mueve o copia UN correo. Devuelve True si ha ido bien y deja
' en "mensaje" el texto para la columna Resultado.
'------------------------------------------------------------------------------
Private Function MoverUnCorreo(ByVal db As Object, ByVal unid As String, _
                               ByVal origen As String, ByVal destino As String, _
                               ByVal mover As Boolean, ByRef mensaje As String) As Boolean
    Dim doc As Object

    MoverUnCorreo = False
    On Error GoTo Fallo

    ' Buscamos el correo por su identificador único
    On Error Resume Next
    Set doc = db.GetDocumentByUNID(unid)
    On Error GoTo Fallo

    If doc Is Nothing Then
        mensaje = "ERROR: el correo ya no está en el buzón (¿borrado?)"
        Exit Function
    End If

    If LCase$(origen) = LCase$(destino) Then
        mensaje = "Ya estaba en '" & destino & "'"
        MoverUnCorreo = True
        Exit Function
    End If

    ' 1) Lo metemos en la carpeta de destino (True = crearla si no existe)
    doc.PutInFolder NombreInterno(destino), True

    ' 2) Si es "mover", lo sacamos de la carpeta de origen
    If mover And Len(origen) > 0 Then
        doc.RemoveFromFolder NombreInterno(origen)
    End If

    mensaje = IIf(mover, "Movido a '", "Copiado a '") & destino & "' " & Format$(Now, "dd/mm/yyyy hh:mm")
    MoverUnCorreo = True
    Exit Function

Fallo:
    mensaje = "ERROR: " & Err.Description
End Function


'------------------------------------------------------------------------------
' MarcarSeleccion: marca como leídos (True) o no leídos (False) los correos de
' las filas seleccionadas en la hoja Bandeja.
'------------------------------------------------------------------------------
Private Sub MarcarSeleccion(ByVal leido As Boolean)
    Dim ws As Worksheet
    Dim db As Object
    Dim doc As Object
    Dim filas As Collection
    Dim fila As Variant
    Dim unid As String
    Dim titulo As String
    Dim marcados As Long
    Dim fallidos As Long

    titulo = IIf(leido, "Marcar como leído", "Marcar como no leído")

    On Error GoTo ManejoError

    Set ws = PrepararHojaBandeja()
    Set filas = FilasSeleccionadas(ws)

    If filas.Count = 0 Then
        MsgBox "Selecciona en la hoja '" & HOJA_BANDEJA & "' las filas de los correos " & _
               "(puedes usar Ctrl+clic o Mayús+clic) y vuelve a pulsar el botón.", _
               vbInformation, titulo
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set db = AbrirBaseCorreo()

    For Each fila In filas
        unid = Trim$(CStr(ws.Cells(fila, COL_B_ID).Value))
        If Len(unid) > 0 Then
            Set doc = Nothing
            On Error Resume Next
            Set doc = db.GetDocumentByUNID(unid)
            If Not doc Is Nothing Then
                Err.Clear
                If leido Then
                    doc.MarkRead
                Else
                    doc.MarkUnread
                End If
            End If

            If doc Is Nothing Or Err.Number <> 0 Then
                fallidos = fallidos + 1
                ws.Cells(fila, COL_B_RESULTADO).Value = "ERROR al marcar: " & _
                    IIf(doc Is Nothing, "correo no encontrado", Err.Description)
                ws.Cells(fila, COL_B_RESULTADO).Font.Color = RGB(192, 0, 0)
            Else
                marcados = marcados + 1
                ws.Cells(fila, COL_B_LEIDO).Value = IIf(leido, "Sí", "No")
            End If
            On Error GoTo ManejoError
        End If
    Next fila

    Application.ScreenUpdating = True
    MsgBox marcados & " correo(s) marcados." & _
           IIf(fallidos > 0, vbCrLf & fallidos & " con error (mira la columna 'Resultado')." & vbCrLf & _
               "Nota: marcar leído/no leído necesita HCL Notes 8 o superior.", ""), _
           IIf(fallidos > 0, vbExclamation, vbInformation), titulo

Salir:
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError titulo, Err.Number, Err.Description
    Resume Salir
End Sub

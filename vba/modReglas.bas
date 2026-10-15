Attribute VB_Name = "modReglas"
'==============================================================================
' MÓDULO: modReglas
'------------------------------------------------------------------------------
' Archiva automáticamente los correos de la Bandeja de entrada según las
' reglas escritas en la hoja "Reglas". Cada fila es una regla:
'
'   A Buscar en   -> Remitente, Para/CC, Asunto o Cuerpo (desplegable)
'   B Contiene    -> texto a buscar (no distingue mayúsculas)
'   C Mover a     -> carpeta de destino (desplegable)
'   D Activa      -> Sí / No
'   E Resultado   -> lo rellena la macro
'
' Ejemplo: Remitente | @acme.com | Clientes\Acme | Sí
'   -> todos los correos de la Bandeja de entrada cuyo remitente contenga
'      "@acme.com" se mueven a la carpeta "Clientes\Acme".
'
' Las reglas se revisan de arriba abajo: si un correo cumple varias, manda
' la PRIMERA. Antes de mover nada, se muestra un resumen y se pide permiso.
'
' Macro principal: AplicarReglas
'==============================================================================
Option Explicit


'------------------------------------------------------------------------------
' AplicarReglas  -> MACRO (botón "Aplicar reglas a la Bandeja de entrada")
'------------------------------------------------------------------------------
Public Sub AplicarReglas()
    Dim ws As Worksheet
    Dim db As Object
    Dim vista As Object
    Dim doc As Object
    Dim ultimaFila As Long
    Dim fila As Long
    Dim n As Long             ' Nº de reglas activas
    Dim r As Long
    Dim i As Long
    Dim campos() As String
    Dim textos() As String
    Dim destinos() As String
    Dim filasRegla() As Long
    Dim cuenta() As Long      ' Cuántos correos cumple cada regla
    Dim errores() As Long
    Dim docs As New Collection       ' Correos a mover...
    Dim reglaDeDoc As New Collection ' ...y qué regla cumple cada uno
    Dim revisados As Long
    Dim resumen As String
    Dim totalMovidos As Long

    On Error GoTo ManejoError

    Set ws = PrepararHojaReglas()
    ultimaFila = ws.Cells(ws.Rows.Count, COL_R_CONTIENE).End(xlUp).Row

    ' --- 1) Cargamos las reglas activas ----------------------------------------
    If ultimaFila >= 2 Then
        ReDim campos(1 To ultimaFila)
        ReDim textos(1 To ultimaFila)
        ReDim destinos(1 To ultimaFila)
        ReDim filasRegla(1 To ultimaFila)
    End If

    For fila = 2 To ultimaFila
        If ReglaActiva(ws, fila) Then
            n = n + 1
            campos(n) = Trim$(CStr(ws.Cells(fila, COL_R_CAMPO).Value))
            textos(n) = LCase$(Trim$(CStr(ws.Cells(fila, COL_R_CONTIENE).Value)))
            destinos(n) = Trim$(CStr(ws.Cells(fila, COL_R_DESTINO).Value))
            filasRegla(n) = fila
        End If
    Next fila

    If n = 0 Then
        MsgBox "No hay reglas activas." & vbCrLf & vbCrLf & _
               "En la hoja '" & HOJA_REGLAS & "' rellena en cada fila: dónde buscar, " & _
               "el texto, la carpeta de destino y 'Sí' en la columna Activa.", _
               vbInformation, "Aplicar reglas"
        Exit Sub
    End If

    ReDim cuenta(1 To n)
    ReDim errores(1 To n)

    ' --- 2) Revisamos la Bandeja de entrada (sin mover nada todavía) -----------
    Application.ScreenUpdating = False
    Application.StatusBar = "Conectando con HCL Notes..."
    Set db = AbrirBaseCorreo()

    Set vista = db.GetView("($Inbox)")
    If vista Is Nothing Then
        Err.Raise ERR_BASE_CORREO, "AplicarReglas", "No se encontró la Bandeja de entrada."
    End If
    vista.AutoUpdate = False

    Set doc = vista.GetFirstDocument()
    Do While Not doc Is Nothing
        revisados = revisados + 1
        If revisados Mod 20 = 0 Then
            Application.StatusBar = "Revisando la Bandeja de entrada: " & revisados & " correos..."
        End If

        For r = 1 To n
            If CumpleRegla(doc, campos(r), textos(r)) Then
                docs.Add doc
                reglaDeDoc.Add r
                cuenta(r) = cuenta(r) + 1
                Exit For            ' Manda la primera regla que se cumple
            End If
        Next r

        Set doc = vista.GetNextDocument(doc)
    Loop

    Application.StatusBar = False
    Application.ScreenUpdating = True

    If docs.Count = 0 Then
        For r = 1 To n
            EscribirResultado ws, filasRegla(r), "0 correos " & Format$(Now, "dd/mm/yyyy hh:mm"), False
        Next r
        MsgBox "Se han revisado " & revisados & " correo(s) de la Bandeja de entrada y " & _
               "ninguno cumple las reglas.", vbInformation, "Aplicar reglas"
        Exit Sub
    End If

    ' --- 3) Resumen y confirmación ---------------------------------------------
    For r = 1 To n
        If cuenta(r) > 0 Then
            resumen = resumen & "  Fila " & filasRegla(r) & ": " & cuenta(r) & _
                      " correo(s)  ->  " & destinos(r) & vbCrLf
        End If
    Next r

    If MsgBox("Se han revisado " & revisados & " correo(s). Se moverán " & docs.Count & ":" & _
              vbCrLf & vbCrLf & resumen & vbCrLf & "¿Continuar?", _
              vbQuestion + vbYesNo, "Aplicar reglas") = vbNo Then Exit Sub

    ' --- 4) Movemos -----------------------------------------------------------
    Application.ScreenUpdating = False
    For i = 1 To docs.Count
        Application.StatusBar = "Moviendo correo " & i & " de " & docs.Count & "..."
        r = reglaDeDoc(i)
        If MoverDesdeBandeja(docs(i), destinos(r)) Then
            totalMovidos = totalMovidos + 1
        Else
            errores(r) = errores(r) + 1
        End If
    Next i

    For r = 1 To n
        EscribirResultado ws, filasRegla(r), _
            (cuenta(r) - errores(r)) & " movido(s)" & _
            IIf(errores(r) > 0, ", " & errores(r) & " con error", "") & _
            " " & Format$(Now, "dd/mm/yyyy hh:mm"), errores(r) > 0
    Next r

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox totalMovidos & " correo(s) movido(s)." & _
           IIf(totalMovidos < docs.Count, vbCrLf & (docs.Count - totalMovidos) & " no se pudieron mover.", "") & _
           vbCrLf & vbCrLf & "Pulsa 'Actualizar carpetas' para ver los nuevos totales.", _
           IIf(totalMovidos < docs.Count, vbExclamation, vbInformation), "Aplicar reglas"

Salir:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Aplicar reglas", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' ReglaActiva: True si la fila tiene texto, destino y no está marcada "No".
'------------------------------------------------------------------------------
Private Function ReglaActiva(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    ReglaActiva = Len(Trim$(CStr(ws.Cells(fila, COL_R_CONTIENE).Value))) > 0 And _
                  Len(Trim$(CStr(ws.Cells(fila, COL_R_DESTINO).Value))) > 0 And _
                  UCase$(Left$(Trim$(CStr(ws.Cells(fila, COL_R_ACTIVA).Value)), 1)) <> "N"
End Function


'------------------------------------------------------------------------------
' CumpleRegla: True si el campo indicado del correo contiene el texto.
' (texto ya en minúsculas)
'------------------------------------------------------------------------------
Private Function CumpleRegla(ByVal doc As Object, ByVal campo As String, ByVal texto As String) As Boolean
    Dim contenido As String

    Select Case LCase$(campo)
        Case "para/cc"
            contenido = TextoCampo(doc, "SendTo") & "; " & TextoCampo(doc, "CopyTo")
        Case "asunto"
            contenido = TextoCampo(doc, "Subject")
        Case "cuerpo"
            contenido = TextoCuerpo(doc)
        Case Else    ' "Remitente" o vacío
            contenido = TextoCampo(doc, "From")
    End Select

    CumpleRegla = InStr(1, LCase$(contenido), texto) > 0
End Function


'------------------------------------------------------------------------------
' MoverDesdeBandeja: mete el correo en la carpeta y lo saca de la Bandeja.
'------------------------------------------------------------------------------
Private Function MoverDesdeBandeja(ByVal doc As Object, ByVal destinoVisible As String) As Boolean
    On Error GoTo Fallo
    doc.PutInFolder NombreInterno(destinoVisible), True
    doc.RemoveFromFolder "($Inbox)"
    MoverDesdeBandeja = True
    Exit Function
Fallo:
    MoverDesdeBandeja = False
End Function


'------------------------------------------------------------------------------
' EscribirResultado: escribe en la columna Resultado (rojo si hubo errores).
'------------------------------------------------------------------------------
Private Sub EscribirResultado(ByVal ws As Worksheet, ByVal fila As Long, _
                              ByVal texto As String, ByVal conError As Boolean)
    ws.Cells(fila, COL_R_RESULTADO).Value = texto
    ws.Cells(fila, COL_R_RESULTADO).Font.Color = IIf(conError, RGB(192, 0, 0), RGB(0, 128, 0))
End Sub

Attribute VB_Name = "modEnviarCorreo"
'==============================================================================
' MÓDULO: modEnviarCorreo
'------------------------------------------------------------------------------
' Envía correos con HCL Notes a partir de la hoja "Enviar":
'   A Para | B CC | C Asunto | D Cuerpo | E Adjunto (ruta) | F Estado
'
' Cada fila es un correo. Solo se envían las filas que:
'   - tienen algo en "Para", y
'   - tienen la columna "Estado" VACÍA (así no se envía dos veces lo mismo).
' Tras enviar, la macro escribe en "Estado" la fecha de envío o el error.
' Para reenviar una fila, borra su celda de "Estado".
'
' Macro principal: EnviarCorreos  (asígnala a un botón)
'==============================================================================
Option Explicit


'------------------------------------------------------------------------------
' EnviarCorreos  -> MACRO PRINCIPAL
'------------------------------------------------------------------------------
Public Sub EnviarCorreos()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim fila As Long
    Dim pendientes As Long
    Dim enviados As Long
    Dim fallidos As Long

    On Error GoTo ManejoError

    Set ws = PrepararHojaEnviar()

    ' Última fila con datos en la columna "Para"
    ultimaFila = ws.Cells(ws.Rows.Count, COL_E_PARA).End(xlUp).Row

    ' 1) Contamos cuántos correos hay pendientes de enviar
    For fila = 2 To ultimaFila
        If FilaPendiente(ws, fila) Then pendientes = pendientes + 1
    Next fila

    If pendientes = 0 Then
        MsgBox "No hay correos pendientes de enviar." & vbCrLf & vbCrLf & _
               "Rellena la hoja '" & HOJA_ENVIAR & "' (al menos 'Para' y 'Asunto') " & _
               "y deja vacía la columna 'Estado'.", vbInformation, "Enviar correo"
        Exit Sub
    End If

    ' 2) Pedimos confirmación: enviar correos no se puede deshacer
    If MsgBox("Se van a enviar " & pendientes & " correo(s) con HCL Notes." & vbCrLf & _
              "¿Quieres continuar?", vbQuestion + vbYesNo, "Enviar correo") = vbNo Then
        Exit Sub
    End If

    ' 3) Conectamos antes del bucle: si falla la conexión, paramos aquí
    '    en lugar de marcar todas las filas con error.
    AbrirBaseCorreo

    ' 4) Enviamos fila a fila
    For fila = 2 To ultimaFila
        If FilaPendiente(ws, fila) Then
            Application.StatusBar = "Enviando correo de la fila " & fila & "..."
            If EnviarFila(ws, fila) Then
                enviados = enviados + 1
            Else
                fallidos = fallidos + 1
            End If
        End If
    Next fila

    Application.StatusBar = False
    MsgBox "Proceso terminado." & vbCrLf & vbCrLf & _
           "Enviados: " & enviados & vbCrLf & _
           "Con error: " & fallidos & IIf(fallidos > 0, "  (mira la columna 'Estado')", ""), _
           IIf(fallidos > 0, vbExclamation, vbInformation), "Enviar correo"

Salir:
    Application.StatusBar = False
    Exit Sub

ManejoError:
    MostrarError "Enviar correo", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' FilaPendiente: True si la fila tiene destinatario y aún no se ha enviado.
'------------------------------------------------------------------------------
Private Function FilaPendiente(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    FilaPendiente = Len(Trim$(CStr(ws.Cells(fila, COL_E_PARA).Value))) > 0 And _
                    Len(Trim$(CStr(ws.Cells(fila, COL_E_ESTADO).Value))) = 0
End Function


'------------------------------------------------------------------------------
' EnviarFila: envía el correo de UNA fila y escribe el resultado en "Estado".
' Devuelve True si se envió bien. Los errores de una fila no detienen las demás.
'------------------------------------------------------------------------------
Private Function EnviarFila(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    Dim celdaEstado As Range
    Set celdaEstado = ws.Cells(fila, COL_E_ESTADO)

    On Error GoTo FalloEnvio

    EnviarUnCorreo _
        para:=CStr(ws.Cells(fila, COL_E_PARA).Value), _
        cc:=CStr(ws.Cells(fila, COL_E_CC).Value), _
        asunto:=CStr(ws.Cells(fila, COL_E_ASUNTO).Value), _
        cuerpo:=CStr(ws.Cells(fila, COL_E_CUERPO).Value), _
        adjuntos:=CStr(ws.Cells(fila, COL_E_ADJUNTO).Value)

    celdaEstado.Value = "Enviado " & Format$(Now, "dd/mm/yyyy hh:nn")
    celdaEstado.Font.Color = RGB(0, 128, 0)     ' Verde
    EnviarFila = True
    Exit Function

FalloEnvio:
    celdaEstado.Value = TextoSeguro("ERROR: " & Err.Description)
    celdaEstado.Font.Color = RGB(192, 0, 0)     ' Rojo
    EnviarFila = False
End Function


'------------------------------------------------------------------------------
' EnviarUnCorreo
' Envía un correo con HCL Notes. Puedes usarla también desde tus propias macros:
'
'   EnviarUnCorreo "ana@empresa.com", "", "Hola", "Texto del mensaje", "C:\informe.pdf"
'
'   para      -> destinatarios, separados por ;   (obligatorio)
'   cc        -> copia, separados por ;           (puede ser "")
'   asunto    -> asunto del correo
'   cuerpo    -> texto. Los saltos de línea de la celda (Alt+Intro) se respetan.
'   adjuntos  -> ruta(s) completa(s) de archivo(s), separadas por ; (puede ser "")
'
' Si algo va mal, lanza un error (que recoge quien la llama).
'------------------------------------------------------------------------------
Public Sub EnviarUnCorreo(ByVal para As String, ByVal cc As String, _
                          ByVal asunto As String, ByVal cuerpo As String, _
                          Optional ByVal adjuntos As String = "")
    Dim db As Object
    Dim doc As Object
    Dim campoCuerpo As Object
    Dim lineas() As String
    Dim rutas As Variant
    Dim i As Long
    Dim fso As Object

    '--- Comprobaciones previas -----------------------------------------------
    If Len(Trim$(para)) = 0 Then
        Err.Raise ERR_DATOS, "EnviarUnCorreo", "Falta el destinatario ('Para')."
    End If

    ' Comprobamos que los adjuntos existen ANTES de crear el correo
    rutas = ListaSeparada(adjuntos, False)   ' En rutas NO usamos la coma como separador
    If UBound(rutas) >= 0 Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        For i = 0 To UBound(rutas)
            If Not fso.FileExists(rutas(i)) Then
                Err.Raise ERR_DATOS, "EnviarUnCorreo", _
                          "No se encuentra el archivo adjunto: " & rutas(i)
            End If
        Next i
    End If

    '--- Crear el correo --------------------------------------------------------
    Set db = AbrirBaseCorreo()
    Set doc = db.CreateDocument()

    ' "Memo" es el formulario estándar de correo en Notes
    doc.ReplaceItemValue "Form", "Memo"
    doc.ReplaceItemValue "SendTo", ListaSeparada(para)
    If Len(Trim$(cc)) > 0 Then
        doc.ReplaceItemValue "CopyTo", ListaSeparada(cc)
    End If
    doc.ReplaceItemValue "Subject", asunto

    ' El cuerpo es un campo de "texto enriquecido" (RichText), que admite
    ' saltos de línea y adjuntos.
    Set campoCuerpo = doc.CreateRichTextItem("Body")

    ' Escribimos el texto línea a línea para respetar los saltos de línea
    cuerpo = Replace(cuerpo, vbCrLf, vbLf)
    lineas = Split(cuerpo, vbLf)
    For i = 0 To UBound(lineas)
        If i > 0 Then campoCuerpo.AddNewLine 1
        campoCuerpo.AppendText lineas(i)
    Next i

    ' Adjuntamos los archivos al final del cuerpo
    If UBound(rutas) >= 0 Then
        campoCuerpo.AddNewLine 2
        For i = 0 To UBound(rutas)
            campoCuerpo.EmbedObject NOTES_EMBED_ATTACHMENT, "", rutas(i)
        Next i
    End If

    '--- Enviar -----------------------------------------------------------------
    ' SaveMessageOnSend = True guarda una copia en tu carpeta "Enviados"
    doc.SaveMessageOnSend = True
    doc.Send False
End Sub


'------------------------------------------------------------------------------
' ListaSeparada: convierte "a@x.com; b@y.com" en un array ("a@x.com","b@y.com").
' Acepta ; como separador (y también , si aceptarComa = True) e ignora los
' huecos vacíos. Si el texto está vacío devuelve un array vacío (UBound = -1).
'------------------------------------------------------------------------------
Private Function ListaSeparada(ByVal texto As String, _
                               Optional ByVal aceptarComa As Boolean = True) As Variant
    Dim partes() As String
    Dim resultado() As String
    Dim i As Long
    Dim n As Long
    Dim valor As String

    If aceptarComa Then texto = Replace(texto, ",", ";")
    partes = Split(texto, ";")

    n = 0
    For i = 0 To UBound(partes)
        valor = Trim$(partes(i))
        If Len(valor) > 0 Then
            ReDim Preserve resultado(0 To n)
            resultado(n) = valor
            n = n + 1
        End If
    Next i

    If n = 0 Then
        ListaSeparada = Array()
    Else
        ListaSeparada = resultado
    End If
End Function

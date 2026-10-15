Attribute VB_Name = "modLeerCorreo"
'==============================================================================
' MÓDULO: modLeerCorreo
'------------------------------------------------------------------------------
' Lee los correos de la Bandeja de entrada de HCL Notes y los copia en la hoja
' "Bandeja" con estas columnas:
'   A Fecha | B Remitente | C Asunto | D Extracto | E Nº adjuntos | F ID interno
'
' Macro principal: LeerBandejaEntrada  (asígnala a un botón)
'
' Un poco de vocabulario de Notes:
'   - Base de datos (NotesDatabase): tu buzón, un archivo .nsf.
'   - Vista (NotesView): una "lista" de documentos. La Bandeja de entrada es
'     una carpeta especial que se llama "($Inbox)".
'   - Documento (NotesDocument): cada correo. Tiene "campos" (items) como
'     Subject (asunto), From (remitente), Body (cuerpo)...
'==============================================================================
Option Explicit


'------------------------------------------------------------------------------
' LeerBandejaEntrada  -> MACRO PRINCIPAL
'------------------------------------------------------------------------------
Public Sub LeerBandejaEntrada()
    Dim sesion As Object      ' Sesión de Notes
    Dim db As Object          ' Base de correo
    Dim vista As Object       ' Carpeta Bandeja de entrada
    Dim doc As Object         ' Correo actual
    Dim ws As Worksheet       ' Hoja de Excel donde escribimos
    Dim fila As Long          ' Fila de Excel en la que escribimos
    Dim contador As Long      ' Cuántos correos llevamos

    ' Si ocurre cualquier error, saltamos a la etiqueta "ManejoError".
    On Error GoTo ManejoError

    ' Preparamos la hoja (la crea si no existe y pone los encabezados).
    Set ws = PrepararHojaBandeja()

    ' Desactivar el refresco de pantalla hace que la macro vaya más rápido.
    Application.ScreenUpdating = False
    Application.StatusBar = "Conectando con HCL Notes..."

    ' 1) Conectar y abrir el buzón
    Set sesion = ObtenerSesion()
    Set db = AbrirBaseCorreo()

    ' 2) Abrir la carpeta "Bandeja de entrada"
    Set vista = db.GetView("($Inbox)")
    If vista Is Nothing Then
        Err.Raise ERR_BASE_CORREO, "LeerBandejaEntrada", _
                  "No se encontró la carpeta Bandeja de entrada en tu buzón."
    End If

    ' 3) Borrar los datos anteriores (dejamos la fila 1 de encabezados)
    ws.Range(ws.Cells(2, COL_B_FECHA), ws.Cells(ws.Rows.Count, COL_B_ID)).ClearContents

    ' 4) Recorrer los correos, del más reciente (el último) al más antiguo
    fila = 2
    contador = 0
    Set doc = vista.GetLastDocument()

    Do While Not doc Is Nothing
        If contador >= MAX_CORREOS Then Exit Do

        contador = contador + 1
        Application.StatusBar = "Leyendo correo " & contador & "..."

        ws.Cells(fila, COL_B_FECHA).Value = FechaDelCorreo(doc)
        ws.Cells(fila, COL_B_REMITENTE).Value = TextoSeguro(NombreLegible(sesion, ValorCampo(doc, "From")))
        ws.Cells(fila, COL_B_ASUNTO).Value = TextoSeguro(CStr(ValorCampo(doc, "Subject")))
        ws.Cells(fila, COL_B_EXTRACTO).Value = TextoSeguro(ExtractoCuerpo(doc))
        ws.Cells(fila, COL_B_ADJUNTOS).Value = ContarAdjuntos(sesion, doc)
        ' El UniversalID identifica el correo de forma única. Lo usa la
        ' macro de guardar adjuntos para volver a encontrar el correo.
        ws.Cells(fila, COL_B_ID).Value = doc.UniversalID

        fila = fila + 1
        ' Pasamos al correo anterior de la carpeta
        Set doc = vista.GetPrevDocument(doc)
    Loop

    ' 5) Dar formato bonito
    ws.Columns(COL_B_FECHA).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Columns(COL_B_FECHA).AutoFit
    ws.Columns(COL_B_ADJUNTOS).AutoFit

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Se han leído " & contador & " correo(s) de la Bandeja de entrada.", _
           vbInformation, "Leer correo"

Salir:
    ' Pase lo que pase, dejamos Excel como estaba.
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Leer la bandeja de entrada", Err.Number, Err.Description
    Resume Salir
End Sub


'==============================================================================
' FUNCIONES DE AYUDA (algunas son Public porque las usa también modAdjuntos)
'==============================================================================

'------------------------------------------------------------------------------
' ValorCampo: devuelve el primer valor de un campo del documento, o "" si no
' existe. En Notes, GetItemValue siempre devuelve una LISTA (array), aunque el
' campo solo tenga un valor; por eso cogemos el primer elemento.
'------------------------------------------------------------------------------
Public Function ValorCampo(ByVal doc As Object, ByVal nombreCampo As String) As Variant
    Dim valores As Variant

    ValorCampo = ""
    On Error GoTo NoHayValor

    If doc.HasItem(nombreCampo) Then
        valores = doc.GetItemValue(nombreCampo)
        If IsArray(valores) Then
            If UBound(valores) >= LBound(valores) Then
                ValorCampo = valores(LBound(valores))
            End If
        Else
            ValorCampo = valores
        End If
    End If
    Exit Function

NoHayValor:
    ValorCampo = ""
End Function


'------------------------------------------------------------------------------
' FechaDelCorreo: fecha de entrega o de envío del correo.
'------------------------------------------------------------------------------
Private Function FechaDelCorreo(ByVal doc As Object) As Variant
    Dim valor As Variant

    ' Probamos varios campos, por orden de preferencia.
    valor = ValorCampo(doc, "DeliveredDate")   ' Cuándo llegó a tu buzón
    If Not IsDate(valor) Then valor = ValorCampo(doc, "PostedDate")  ' Cuándo se envió
    If Not IsDate(valor) Then
        On Error Resume Next
        valor = doc.Created                     ' Cuándo se creó el documento
        On Error GoTo 0
    End If

    If IsDate(valor) Then
        FechaDelCorreo = CDate(valor)
    Else
        FechaDelCorreo = ""
    End If
End Function


'------------------------------------------------------------------------------
' NombreLegible: Notes guarda los nombres así: "CN=Juan Pérez/O=Empresa".
' Esta función lo convierte en "Juan Pérez".
'------------------------------------------------------------------------------
Private Function NombreLegible(ByVal sesion As Object, ByVal nombre As Variant) As String
    Dim objNombre As Object
    Dim resultado As String

    resultado = CStr(nombre)
    If Len(resultado) = 0 Then
        NombreLegible = ""
        Exit Function
    End If

    On Error Resume Next
    Set objNombre = sesion.CreateName(resultado)
    If Not objNombre Is Nothing Then
        If Len(objNombre.Common) > 0 Then resultado = objNombre.Common
    End If
    On Error GoTo 0

    NombreLegible = resultado
End Function


'------------------------------------------------------------------------------
' ExtractoCuerpo: primeros caracteres del cuerpo, en una sola línea.
'------------------------------------------------------------------------------
Private Function ExtractoCuerpo(ByVal doc As Object) As String
    Dim campo As Object
    Dim texto As String

    On Error Resume Next
    Set campo = doc.GetFirstItem("Body")
    If Not campo Is Nothing Then
        ' .Text devuelve el contenido como texto plano (sin formato)
        texto = campo.Text
    End If
    On Error GoTo 0

    ' Quitamos saltos de línea y tabuladores para que quepa en una celda
    texto = Replace(texto, vbCrLf, " ")
    texto = Replace(texto, vbCr, " ")
    texto = Replace(texto, vbLf, " ")
    texto = Replace(texto, vbTab, " ")
    ' Quitamos espacios repetidos
    Do While InStr(texto, "  ") > 0
        texto = Replace(texto, "  ", " ")
    Loop
    texto = Trim$(texto)

    If Len(texto) > LONGITUD_EXTRACTO Then
        texto = Left$(texto, LONGITUD_EXTRACTO) & "..."
    End If

    ExtractoCuerpo = texto
End Function


'------------------------------------------------------------------------------
' NombresAdjuntos: devuelve una lista (array) con los nombres de los archivos
' adjuntos del correo. Usa la fórmula de Notes @AttachmentNames.
' Si no hay adjuntos devuelve un array vacío (o con un único texto vacío).
'------------------------------------------------------------------------------
Public Function NombresAdjuntos(ByVal sesion As Object, ByVal doc As Object) As Variant
    Dim resultado As Variant

    On Error Resume Next
    resultado = sesion.Evaluate("@AttachmentNames", doc)
    On Error GoTo 0

    If IsArray(resultado) Then
        NombresAdjuntos = resultado
    Else
        NombresAdjuntos = Array()   ' Lista vacía
    End If
End Function


'------------------------------------------------------------------------------
' ContarAdjuntos: cuántos archivos adjuntos tiene el correo.
'------------------------------------------------------------------------------
Private Function ContarAdjuntos(ByVal sesion As Object, ByVal doc As Object) As Long
    Dim nombres As Variant
    Dim i As Long
    Dim total As Long

    nombres = NombresAdjuntos(sesion, doc)
    For i = LBound(nombres) To UBound(nombres)
        If Len(CStr(nombres(i))) > 0 Then total = total + 1
    Next i

    ContarAdjuntos = total
End Function


'------------------------------------------------------------------------------
' TextoSeguro: si un texto empieza por = + - o @, Excel creería que es una
' fórmula. Le ponemos un apóstrofo delante para que se guarde como texto.
'------------------------------------------------------------------------------
Public Function TextoSeguro(ByVal texto As String) As String
    If Len(texto) > 0 Then
        Select Case Left$(texto, 1)
            Case "=", "+", "-", "@"
                texto = "'" & texto
        End Select
    End If
    TextoSeguro = texto
End Function

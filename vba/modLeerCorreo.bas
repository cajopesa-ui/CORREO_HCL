Attribute VB_Name = "modLeerCorreo"
'==============================================================================
' MÓDULO: modLeerCorreo
'------------------------------------------------------------------------------
' Lee correos de HCL Notes y los copia en la hoja "Bandeja" con estas columnas:
'   A Fecha | B Remitente | C Asunto | D Extracto | E Nº adjuntos | F ID interno
'   G Carpeta | H Leído | I Mover a | J Resultado
'
' Macro principal: LeerBandejaEntrada  (asígnala a un botón)
'
' La función LeerCorreosEnHoja es el "motor" que usan también la lectura de
' carpetas y la búsqueda (módulo modBuscar): lee una o varias carpetas,
' aplica filtros (texto, remitente, fechas) y escribe el resultado.
'
' Un poco de vocabulario de Notes:
'   - Base de datos (NotesDatabase): tu buzón, un archivo .nsf.
'   - Vista (NotesView): una "lista" de documentos. Las carpetas son vistas
'     especiales. La Bandeja de entrada se llama por dentro "($Inbox)".
'   - Documento (NotesDocument): cada correo. Tiene "campos" (items) como
'     Subject (asunto), From (remitente), Body (cuerpo)...
'==============================================================================
Option Explicit


'------------------------------------------------------------------------------
' LeerBandejaEntrada  -> MACRO PRINCIPAL
'------------------------------------------------------------------------------
Public Sub LeerBandejaEntrada()
    Dim db As Object
    Dim carpetas As New Collection
    Dim contador As Long
    Dim limiteAlcanzado As Boolean

    ' Si ocurre cualquier error, saltamos a la etiqueta "ManejoError".
    On Error GoTo ManejoError

    ' Desactivar el refresco de pantalla hace que la macro vaya más rápido.
    Application.ScreenUpdating = False
    Application.StatusBar = "Conectando con HCL Notes..."

    ' Comprobamos que la Bandeja de entrada existe en el buzón
    Set db = AbrirBaseCorreo()
    If db.GetView("($Inbox)") Is Nothing Then
        Err.Raise ERR_BASE_CORREO, "LeerBandejaEntrada", _
                  "No se encontró la carpeta Bandeja de entrada en tu buzón."
    End If

    ' Leemos solo la Bandeja de entrada, sin filtros
    carpetas.Add "($Inbox)"
    contador = LeerCorreosEnHoja(carpetas, MAX_CORREOS, "", "", 0, 0, limiteAlcanzado)

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


'------------------------------------------------------------------------------
' LeerCorreosEnHoja  -> EL "MOTOR" DE LECTURA
' Lee los correos de las carpetas indicadas y los escribe en la hoja Bandeja
' (borrando antes lo que hubiera). Devuelve cuántos correos ha escrito.
'
'   carpetas  -> lista con los nombres INTERNOS de las carpetas
'                (ej. "($Inbox)", "Clientes\Acme")
'   maximo    -> número máximo de correos a escribir
'   texto     -> solo correos cuyo asunto o cuerpo contengan este texto ("" = todos)
'   remitente -> solo correos cuyo remitente contenga este texto ("" = todos)
'   desde     -> solo correos a partir de esta fecha (0 = sin límite)
'   hasta     -> solo correos hasta esta fecha, incluida (0 = sin límite)
'   limiteAlcanzado -> (de salida) True si se paró por llegar al máximo
'
' Si un correo está en varias de esas carpetas, solo se escribe una vez.
'------------------------------------------------------------------------------
Public Function LeerCorreosEnHoja(ByVal carpetas As Collection, ByVal maximo As Long, _
                                  ByVal texto As String, ByVal remitente As String, _
                                  ByVal desde As Date, ByVal hasta As Date, _
                                  ByRef limiteAlcanzado As Boolean) As Long
    Dim sesion As Object
    Dim db As Object
    Dim vista As Object
    Dim doc As Object
    Dim ws As Worksheet
    Dim vistos As Object          ' Diccionario con los correos ya revisados
    Dim carpeta As Variant
    Dim unid As String
    Dim fila As Long
    Dim contador As Long
    Dim revisados As Long

    Set ws = PrepararHojaBandeja()
    Set sesion = ObtenerSesion()
    Set db = AbrirBaseCorreo()

    ' La columna "Mover a" necesita la lista de carpetas. Si todavía no se ha
    ' creado (primera vez), la creamos ahora.
    AsegurarListaCarpetas db

    LimpiarHojaBandeja ws

    ' Un "Dictionary" es como una lista con buscador: nos dice muy rápido si
    ' un correo ya ha salido (por si está en dos carpetas a la vez).
    Set vistos = CreateObject("Scripting.Dictionary")

    ' Comparamos siempre en minúsculas para que "Factura" = "factura"
    texto = LCase$(Trim$(texto))
    remitente = LCase$(Trim$(remitente))

    fila = 2
    limiteAlcanzado = False

    For Each carpeta In carpetas
        If limiteAlcanzado Then Exit For

        ' Si la carpeta no existe, la saltamos sin dar error
        Set vista = Nothing
        On Error Resume Next
        Set vista = db.GetView(CStr(carpeta))
        vista.AutoUpdate = False      ' Más rápido y estable mientras leemos
        On Error GoTo 0

        If Not vista Is Nothing Then
            ' Del más reciente (el último) al más antiguo
            Set doc = vista.GetLastDocument()
            Do While Not doc Is Nothing
                revisados = revisados + 1
                If revisados Mod 20 = 0 Then
                    Application.StatusBar = "Carpeta '" & NombreVisible(CStr(carpeta)) & _
                        "': revisados " & revisados & ", encontrados " & contador & "..."
                End If

                unid = doc.UniversalID
                If Not vistos.Exists(unid) Then
                    vistos.Add unid, True

                    If CumpleFiltro(doc, texto, remitente, desde, hasta) Then
                        If contador >= maximo Then
                            limiteAlcanzado = True
                            Exit Do
                        End If
                        EscribirCorreo ws, fila, sesion, doc, NombreVisible(CStr(carpeta))
                        fila = fila + 1
                        contador = contador + 1
                    End If
                End If

                ' Pasamos al correo anterior de la carpeta
                Set doc = vista.GetPrevDocument(doc)
            Loop
        End If
    Next carpeta

    FormatearHojaBandeja ws, fila - 1
    LeerCorreosEnHoja = contador
End Function


'------------------------------------------------------------------------------
' EscribirCorreo: escribe los datos de un correo en una fila de la hoja.
'------------------------------------------------------------------------------
Public Sub EscribirCorreo(ByVal ws As Worksheet, ByVal fila As Long, ByVal sesion As Object, _
                          ByVal doc As Object, ByVal carpetaVisible As String)
    ws.Cells(fila, COL_B_FECHA).Value = FechaDelCorreo(doc)
    ws.Cells(fila, COL_B_REMITENTE).Value = TextoSeguro(NombreLegible(sesion, ValorCampo(doc, "From")))
    ws.Cells(fila, COL_B_ASUNTO).Value = TextoSeguro(CStr(ValorCampo(doc, "Subject")))
    ws.Cells(fila, COL_B_EXTRACTO).Value = TextoSeguro(ExtractoCuerpo(doc))
    ws.Cells(fila, COL_B_ADJUNTOS).Value = ContarAdjuntos(sesion, doc)
    ' El UniversalID identifica el correo de forma única. Lo usan las macros
    ' de guardar adjuntos, mover y marcar para volver a encontrar el correo.
    ws.Cells(fila, COL_B_ID).Value = doc.UniversalID
    ws.Cells(fila, COL_B_CARPETA).Value = TextoSeguro(carpetaVisible)
    ws.Cells(fila, COL_B_LEIDO).Value = TextoLeido(doc)
End Sub


'------------------------------------------------------------------------------
' LimpiarHojaBandeja: borra los correos anteriores (deja los encabezados).
'------------------------------------------------------------------------------
Public Sub LimpiarHojaBandeja(ByVal ws As Worksheet)
    Dim rango As Range

    Set rango = ws.Range(ws.Cells(2, COL_B_FECHA), ws.Cells(ws.Rows.Count, COL_B_RESULTADO))
    rango.ClearContents
    rango.Font.ColorIndex = xlColorIndexAutomatic
    ' La columna ID se ve en gris (es interna)
    ws.Columns(COL_B_ID).Font.Color = RGB(150, 150, 150)

    ' Quitamos los desplegables viejos de "Mover a"
    ws.Range(ws.Cells(2, COL_B_MOVER), ws.Cells(ws.Rows.Count, COL_B_MOVER)).Validation.Delete
End Sub


'------------------------------------------------------------------------------
' FormatearHojaBandeja: formato de fechas, orden (más recientes arriba) y
' desplegable de "Mover a" en las filas con correos.
'------------------------------------------------------------------------------
Public Sub FormatearHojaBandeja(ByVal ws As Worksheet, ByVal ultimaFila As Long)
    ws.Columns(COL_B_FECHA).NumberFormat = "dd/mm/yyyy hh:mm"

    If ultimaFila >= 3 Then
        ' Ordenamos por fecha, de la más reciente a la más antigua
        ws.Range(ws.Cells(1, COL_B_FECHA), ws.Cells(ultimaFila, COL_B_RESULTADO)).Sort _
            Key1:=ws.Cells(1, COL_B_FECHA), Order1:=xlDescending, Header:=xlYes
    End If

    If ultimaFila >= 2 Then
        AplicarDesplegableCarpetas ws.Range(ws.Cells(2, COL_B_MOVER), ws.Cells(ultimaFila, COL_B_MOVER))
    End If
End Sub


'==============================================================================
' FUNCIONES DE AYUDA (algunas son Public porque las usan otros módulos)
'==============================================================================

'------------------------------------------------------------------------------
' CumpleFiltro: True si el correo cumple TODOS los filtros indicados.
' (texto y remitente deben venir ya en minúsculas)
'------------------------------------------------------------------------------
Private Function CumpleFiltro(ByVal doc As Object, ByVal texto As String, _
                              ByVal remitente As String, _
                              ByVal desde As Date, ByVal hasta As Date) As Boolean
    Dim fecha As Variant

    CumpleFiltro = False

    ' 1) Fechas
    If desde <> 0 Or hasta <> 0 Then
        fecha = FechaDelCorreo(doc)
        If Not IsDate(fecha) Then Exit Function
        If desde <> 0 And CDate(fecha) < desde Then Exit Function
        ' "hasta" incluye todo ese día: comparamos con el día siguiente
        If hasta <> 0 And CDate(fecha) >= hasta + 1 Then Exit Function
    End If

    ' 2) Remitente (vale el nombre o la dirección de correo)
    If Len(remitente) > 0 Then
        If InStr(1, LCase$(TextoCampo(doc, "From")), remitente) = 0 Then Exit Function
    End If

    ' 3) Texto en el asunto o en el cuerpo (el cuerpo solo si hace falta,
    '    porque leerlo es más lento)
    If Len(texto) > 0 Then
        If InStr(1, LCase$(TextoCampo(doc, "Subject")), texto) = 0 Then
            If InStr(1, LCase$(TextoCuerpo(doc)), texto) = 0 Then Exit Function
        End If
    End If

    CumpleFiltro = True
End Function


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
' TextoCampo: TODOS los valores de un campo unidos con "; " (útil para campos
' con varios valores, como los destinatarios). "" si no existe.
'------------------------------------------------------------------------------
Public Function TextoCampo(ByVal doc As Object, ByVal nombreCampo As String) As String
    Dim valores As Variant
    Dim i As Long
    Dim resultado As String

    On Error GoTo NoHayValor

    If doc.HasItem(nombreCampo) Then
        valores = doc.GetItemValue(nombreCampo)
        If IsArray(valores) Then
            For i = LBound(valores) To UBound(valores)
                If i > LBound(valores) Then resultado = resultado & "; "
                resultado = resultado & CStr(valores(i))
            Next i
        Else
            resultado = CStr(valores)
        End If
    End If

    TextoCampo = resultado
    Exit Function

NoHayValor:
    TextoCampo = resultado
End Function


'------------------------------------------------------------------------------
' FechaDelCorreo: fecha de entrega o de envío del correo ("" si no tiene).
'------------------------------------------------------------------------------
Public Function FechaDelCorreo(ByVal doc As Object) As Variant
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
' TextoCuerpo: todo el cuerpo del correo como texto plano ("" si no tiene).
'------------------------------------------------------------------------------
Public Function TextoCuerpo(ByVal doc As Object) As String
    Dim campo As Object

    On Error Resume Next
    Set campo = doc.GetFirstItem("Body")
    If Not campo Is Nothing Then
        ' .Text devuelve el contenido como texto plano (sin formato)
        TextoCuerpo = campo.Text
    End If
    On Error GoTo 0
End Function


'------------------------------------------------------------------------------
' ExtractoCuerpo: primeros caracteres del cuerpo, en una sola línea.
'------------------------------------------------------------------------------
Private Function ExtractoCuerpo(ByVal doc As Object) As String
    Dim texto As String

    texto = TextoCuerpo(doc)

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
' TextoLeido: "Sí" si ya has leído el correo, "No" si no, "" si Notes no
' puede decirlo (versiones antiguas).
'------------------------------------------------------------------------------
Public Function TextoLeido(ByVal doc As Object) As String
    Dim leido As Boolean

    On Error GoTo NoSeSabe
    leido = doc.Read()
    If leido Then
        TextoLeido = "Sí"
    Else
        TextoLeido = "No"
    End If
    Exit Function

NoSeSabe:
    TextoLeido = ""
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

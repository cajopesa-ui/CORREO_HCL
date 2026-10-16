Attribute VB_Name = "modClasificar"
'==============================================================================
' MÓDULO: modClasificar
'------------------------------------------------------------------------------
' Clasifica los correos POR TEMAS y los guarda en carpetas que cuelgan de una
' carpeta propia, "Clasificacion nueva" (ver CARPETA_CLASIFICACION), para que
' no se mezclen con las que ya tienes.
'
' Hay dos formas de obtener la clasificación, y las dos acaban en la hoja
' "Clasificacion", donde puedes REVISARLA antes de aplicarla:
'
'   A) ImportarClasificacion -> carga el archivo que te prepara Claude a
'      partir del .txt de contenido completo. Cada línea del archivo es:
'          ID;Carpeta;Asunto (opcional, solo para que lo leas tú)
'      ejemplo: 8A6F511905F202E4C1258E92002AC05D;RRHH;Jubilación voluntaria
'
'   B) ClasificarConReglas -> clasifica la Bandeja de entrada con las
'      palabras clave de la hoja "Palabras clave" (sin salir de tu PC).
'
' Después:
'   AplicarClasificacion -> crea las carpetas que falten y COPIA cada correo a
'      su carpeta (o lo MUEVE, si CLASIFICACION_MUEVE = True en modConfiguracion).
'      Solo se aplican las filas con "Sí" en la columna Aplicar.
'==============================================================================
Option Explicit


'------------------------------------------------------------------------------
' ClasificarConReglas  -> MACRO (botón "Clasificar con palabras clave")
'------------------------------------------------------------------------------
Public Sub ClasificarConReglas()
    Dim ws As Worksheet
    Dim wsPal As Worksheet
    Dim sesion As Object
    Dim db As Object
    Dim vista As Object
    Dim doc As Object
    Dim respuesta As Variant
    Dim maximo As Long
    Dim n As Long                  ' Nº de reglas activas
    Dim r As Long
    Dim fila As Long
    Dim ultimaFila As Long
    Dim carpetas() As String
    Dim campos() As String
    Dim palabras() As Variant      ' Cada elemento: lista de palabras de esa regla
    Dim filasRegla() As Long
    Dim carpeta As String
    Dim motivo As String
    Dim clasificados As Long
    Dim sinClasificar As Long
    Dim mimeAntes As Boolean

    On Error GoTo ManejoError

    Set wsPal = ObtenerHoja(HOJA_PALABRAS)
    If Len(CStr(wsPal.Cells(2, COL_P_CARPETA).Value)) = 0 Then PrepararHojaPalabras

    ' --- 1) Cargamos las reglas activas -----------------------------------------
    ultimaFila = wsPal.Cells(wsPal.Rows.Count, COL_P_CARPETA).End(xlUp).Row
    If ultimaFila >= 2 Then
        ReDim carpetas(1 To ultimaFila)
        ReDim campos(1 To ultimaFila)
        ReDim palabras(1 To ultimaFila)
        ReDim filasRegla(1 To ultimaFila)
    End If

    For fila = 2 To ultimaFila
        If Len(Trim$(CStr(wsPal.Cells(fila, COL_P_CARPETA).Value))) > 0 And _
           Len(Trim$(CStr(wsPal.Cells(fila, COL_P_PALABRAS).Value))) > 0 And _
           UCase$(Left$(Trim$(CStr(wsPal.Cells(fila, COL_P_ACTIVA).Value)), 1)) <> "N" Then
            n = n + 1
            carpetas(n) = LimpiarRuta(CStr(wsPal.Cells(fila, COL_P_CARPETA).Value))
            campos(n) = LCase$(Trim$(CStr(wsPal.Cells(fila, COL_P_CAMPO).Value)))
            palabras(n) = Split(Normalizar(CStr(wsPal.Cells(fila, COL_P_PALABRAS).Value)), ";")
            filasRegla(n) = fila
        End If
    Next fila

    If n = 0 Then
        MsgBox "No hay reglas activas en la hoja '" & HOJA_PALABRAS & "'.", vbInformation, "Clasificar"
        Exit Sub
    End If

    ' --- 2) ¿Cuántos correos? ----------------------------------------------------
    respuesta = Application.InputBox( _
        Prompt:="¿Cuántos correos de la Bandeja de entrada quieres clasificar?" & vbCrLf & vbCrLf & _
                "(Se empieza por los más recientes.)", _
        Title:="Clasificar con palabras clave", Default:=MAX_CORREOS, Type:=1)
    If VarType(respuesta) = vbBoolean Then Exit Sub
    If respuesta < 1 Or respuesta <> Int(respuesta) Then
        MsgBox "Escribe un número entero mayor que 0.", vbExclamation, "Clasificar"
        Exit Sub
    End If
    maximo = CLng(respuesta)

    ' --- 3) Recorremos la Bandeja de entrada -----------------------------------
    Application.ScreenUpdating = False
    Application.StatusBar = "Conectando con HCL Notes..."
    Set sesion = ObtenerSesion()
    Set db = AbrirBaseCorreo()

    Set vista = db.GetView("($Inbox)")
    If vista Is Nothing Then
        Err.Raise ERR_BASE_CORREO, "ClasificarConReglas", "No se encontró la Bandeja de entrada."
    End If
    vista.AutoUpdate = False

    ' Para poder leer el cuerpo de los correos en formato web (HTML)
    On Error Resume Next
    mimeAntes = sesion.ConvertMime
    sesion.ConvertMime = False
    On Error GoTo ManejoError

    Set ws = PrepararHojaClasificacion()
    LimpiarHojaClasificacion ws

    fila = 2
    Set doc = vista.GetLastDocument()
    Do While Not doc Is Nothing And fila - 2 < maximo
        Application.StatusBar = "Clasificando correo " & (fila - 1) & "..."

        carpeta = ""
        motivo = ""
        For r = 1 To n
            If CumplePalabras(doc, campos(r), palabras(r)) Then
                carpeta = carpetas(r)
                motivo = "Palabras clave, fila " & filasRegla(r)
                Exit For            ' Gana la primera regla que se cumple
            End If
        Next r

        EscribirFilaClasificacion ws, fila, sesion, doc, carpeta, motivo
        If Len(carpeta) > 0 Then
            clasificados = clasificados + 1
        Else
            sinClasificar = sinClasificar + 1
        End If

        fila = fila + 1
        Set doc = vista.GetPrevDocument(doc)
    Loop

    On Error Resume Next
    sesion.ConvertMime = mimeAntes
    On Error GoTo ManejoError

    Application.StatusBar = False
    Application.ScreenUpdating = True
    ws.Activate
    MsgBox "Clasificados: " & clasificados & vbCrLf & _
           "Sin clasificar: " & sinClasificar & " (marcados con Aplicar = No)" & vbCrLf & vbCrLf & _
           "Revisa la hoja '" & HOJA_CLASIFICACION & "': puedes cambiar la carpeta o poner " & _
           "Aplicar = No. Cuando esté a tu gusto, pulsa 'Aplicar clasificación'.", _
           vbInformation, "Clasificar con palabras clave"

Salir:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Clasificar con palabras clave", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' ImportarClasificacion  -> MACRO (botón "Importar clasificación de Claude")
'------------------------------------------------------------------------------
Public Sub ImportarClasificacion()
    Dim ws As Worksheet
    Dim sesion As Object
    Dim db As Object
    Dim doc As Object
    Dim ruta As Variant
    Dim lineas As Variant
    Dim linea As String
    Dim unid As String
    Dim carpeta As String
    Dim p As Long
    Dim i As Long
    Dim fila As Long
    Dim importados As Long
    Dim noEncontrados As Long
    Dim invalidas As Long

    On Error GoTo ManejoError

    ' --- 1) Elegir el archivo ------------------------------------------------------
    ruta = Application.GetOpenFilename( _
               FileFilter:="Clasificación (*.txt;*.csv), *.txt;*.csv", _
               Title:="Elige el archivo de clasificación que te ha preparado Claude")
    If VarType(ruta) = vbBoolean Then Exit Sub        ' Cancelar

    lineas = Split(LeerTextoUtf8(CStr(ruta)), vbLf)

    ' --- 2) Conectar y preparar la hoja ------------------------------------------
    Application.ScreenUpdating = False
    Application.StatusBar = "Conectando con HCL Notes..."
    Set sesion = ObtenerSesion()
    Set db = AbrirBaseCorreo()

    Set ws = PrepararHojaClasificacion()
    LimpiarHojaClasificacion ws
    fila = 2

    ' --- 3) Una línea por correo: ID;Carpeta;Asunto ------------------------------
    For i = LBound(lineas) To UBound(lineas)
        linea = Trim$(Replace(CStr(lineas(i)), vbCr, ""))

        ' Saltamos líneas vacías, comentarios (#) y la cabecera
        If Len(linea) > 0 And Left$(linea, 1) <> "#" And UCase$(Left$(linea, 3)) <> "ID;" Then
            p = InStr(linea, ";")
            If p = 0 Then
                invalidas = invalidas + 1
            Else
                unid = Trim$(Left$(linea, p - 1))
                carpeta = Mid$(linea, p + 1)
                p = InStr(carpeta, ";")
                If p > 0 Then carpeta = Left$(carpeta, p - 1)
                carpeta = LimpiarRuta(carpeta)

                If Len(unid) <> 32 Or Len(carpeta) = 0 Then
                    invalidas = invalidas + 1
                Else
                    Application.StatusBar = "Importando " & (fila - 1) & "..."
                    Set doc = Nothing
                    On Error Resume Next
                    Set doc = db.GetDocumentByUNID(unid)
                    On Error GoTo ManejoError

                    If doc Is Nothing Then
                        noEncontrados = noEncontrados + 1
                        ws.Cells(fila, COL_K_CARPETA).Value = TextoSeguro(carpeta)
                        ws.Cells(fila, COL_K_MOTIVO).Value = "Archivo de Claude"
                        ws.Cells(fila, COL_K_APLICAR).Value = "No"
                        ws.Cells(fila, COL_K_ID).Value = unid
                        ws.Cells(fila, COL_K_RESULTADO).Value = "ERROR: el correo no está en el buzón"
                        ws.Cells(fila, COL_K_RESULTADO).Font.Color = RGB(192, 0, 0)
                    Else
                        EscribirFilaClasificacion ws, fila, sesion, doc, carpeta, "Archivo de Claude"
                        importados = importados + 1
                    End If
                    fila = fila + 1
                End If
            End If
        End If
    Next i

    Application.StatusBar = False
    Application.ScreenUpdating = True
    ws.Activate
    MsgBox "Importados: " & importados & " correo(s)." & _
           IIf(noEncontrados > 0, vbCrLf & "No encontrados en el buzón: " & noEncontrados, "") & _
           IIf(invalidas > 0, vbCrLf & "Líneas no válidas (ignoradas): " & invalidas, "") & vbCrLf & vbCrLf & _
           "Revisa la hoja '" & HOJA_CLASIFICACION & "' y pulsa 'Aplicar clasificación'.", _
           vbInformation, "Importar clasificación"

Salir:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Importar clasificación", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' AplicarClasificacion  -> MACRO (botón "Aplicar clasificación")
' Crea las carpetas que falten dentro de "Clasificacion nueva" y copia (o
' mueve) cada correo marcado con Aplicar = Sí a su carpeta.
'------------------------------------------------------------------------------
Public Sub AplicarClasificacion()
    Dim ws As Worksheet
    Dim db As Object
    Dim doc As Object
    Dim destinos As Object          ' Diccionario: carpetas de destino distintas
    Dim nuevas As Long
    Dim ultimaFila As Long
    Dim fila As Long
    Dim pendientes As Long
    Dim correctos As Long
    Dim fallidos As Long
    Dim destino As String
    Dim accion As String
    Dim d As Variant

    accion = IIf(CLASIFICACION_MUEVE, "mover", "copiar")

    On Error GoTo ManejoError

    Set ws = PrepararHojaClasificacion()
    ultimaFila = ws.Cells(ws.Rows.Count, COL_K_ID).End(xlUp).Row
    Set destinos = CreateObject("Scripting.Dictionary")

    ' --- 1) ¿Qué hay que hacer? ----------------------------------------------------
    For fila = 2 To ultimaFila
        If FilaParaAplicar(ws, fila) Then
            pendientes = pendientes + 1
            destino = RutaCompleta(CStr(ws.Cells(fila, COL_K_CARPETA).Value))
            If Not destinos.Exists(LCase$(destino)) Then destinos.Add LCase$(destino), destino
        End If
    Next fila

    If pendientes = 0 Then
        MsgBox "No hay ningún correo para aplicar." & vbCrLf & vbCrLf & _
               "Primero pulsa 'Clasificar con palabras clave' o 'Importar clasificación de Claude', " & _
               "y comprueba que las filas tienen carpeta y Aplicar = Sí.", vbInformation, "Aplicar clasificación"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set db = AbrirBaseCorreo()
    For Each d In destinos.Items
        If Not ExisteCarpeta(db, CStr(d)) Then nuevas = nuevas + 1
    Next d
    Application.ScreenUpdating = True

    If MsgBox("Se van a " & UCase$(accion) & " " & pendientes & " correo(s) a " & destinos.Count & _
              " carpeta(s) dentro de '" & CARPETA_CLASIFICACION & "'." & vbCrLf & _
              "Carpetas nuevas que se crearán: " & nuevas & vbCrLf & vbCrLf & _
              IIf(CLASIFICACION_MUEVE, "Los correos SALDRÁN de la Bandeja de entrada.", _
                  "Los correos SEGUIRÁN también en la Bandeja de entrada (es una copia).") & vbCrLf & vbCrLf & _
              "¿Continuar?", vbQuestion + vbYesNo, "Aplicar clasificación") = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    ' --- 2) Creamos las carpetas (con sus "padres") --------------------------------
    For Each d In destinos.Items
        Application.StatusBar = "Creando carpeta '" & d & "'..."
        CrearCarpetaConPadres db, CStr(d)
    Next d

    ' --- 3) Copiamos / movemos cada correo ---------------------------------------
    For fila = 2 To ultimaFila
        If FilaParaAplicar(ws, fila) Then
            Application.StatusBar = "Aplicando fila " & fila & "..."
            destino = RutaCompleta(CStr(ws.Cells(fila, COL_K_CARPETA).Value))

            Set doc = Nothing
            On Error Resume Next
            Set doc = db.GetDocumentByUNID(CStr(ws.Cells(fila, COL_K_ID).Value))
            If Not doc Is Nothing Then
                Err.Clear
                doc.PutInFolder destino, True
                If CLASIFICACION_MUEVE And Err.Number = 0 Then doc.RemoveFromFolder "($Inbox)"
            End If

            If doc Is Nothing Or Err.Number <> 0 Then
                fallidos = fallidos + 1
                ws.Cells(fila, COL_K_RESULTADO).Value = "ERROR: " & _
                    IIf(doc Is Nothing, "el correo no está en el buzón", Err.Description)
                ws.Cells(fila, COL_K_RESULTADO).Font.Color = RGB(192, 0, 0)
            Else
                correctos = correctos + 1
                ws.Cells(fila, COL_K_RESULTADO).Value = IIf(CLASIFICACION_MUEVE, "Movido a '", "Copiado a '") & _
                    destino & "' " & Format$(Now, "dd/mm/yyyy hh:mm")
                ws.Cells(fila, COL_K_RESULTADO).Font.Color = RGB(0, 128, 0)
                ws.Cells(fila, COL_K_APLICAR).Value = "Hecho"
            End If
            On Error GoTo ManejoError
        End If
    Next fila

    ' Actualizamos la hoja Carpetas para que se vean las nuevas
    Application.StatusBar = "Actualizando la hoja Carpetas..."
    RellenarHojaCarpetas db

    Application.StatusBar = False
    Application.ScreenUpdating = True
    ws.Activate
    MsgBox "Hecho." & vbCrLf & vbCrLf & _
           "Correctos: " & correctos & vbCrLf & _
           "Con error: " & fallidos & IIf(fallidos > 0, "  (mira la columna 'Resultado')", "") & vbCrLf & _
           "Carpetas nuevas: " & nuevas, _
           IIf(fallidos > 0, vbExclamation, vbInformation), "Aplicar clasificación"

Salir:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Aplicar clasificación", Err.Number, Err.Description
    Resume Salir
End Sub


'==============================================================================
' FUNCIONES INTERNAS
'==============================================================================

'------------------------------------------------------------------------------
' EscribirFilaClasificacion: escribe un correo en la hoja Clasificacion.
' Si la carpeta está vacía, pone "Sin clasificar" y Aplicar = No.
'------------------------------------------------------------------------------
Private Sub EscribirFilaClasificacion(ByVal ws As Worksheet, ByVal fila As Long, ByVal sesion As Object, _
                                      ByVal doc As Object, ByVal carpeta As String, ByVal motivo As String)
    ws.Cells(fila, COL_K_FECHA).Value = FechaDelCorreo(doc)
    ws.Cells(fila, COL_K_REMITENTE).Value = TextoSeguro(NombreLegible(sesion, ValorCampo(doc, "From")))
    ws.Cells(fila, COL_K_ASUNTO).Value = TextoSeguro(CStr(ValorCampo(doc, "Subject")))
    ws.Cells(fila, COL_K_ID).Value = doc.UniversalID

    If Len(carpeta) > 0 Then
        ws.Cells(fila, COL_K_CARPETA).Value = TextoSeguro(carpeta)
        ws.Cells(fila, COL_K_MOTIVO).Value = motivo
        ws.Cells(fila, COL_K_APLICAR).Value = "Sí"
    Else
        ws.Cells(fila, COL_K_CARPETA).Value = SIN_CLASIFICAR
        ws.Cells(fila, COL_K_MOTIVO).Value = "Ninguna palabra clave"
        ws.Cells(fila, COL_K_APLICAR).Value = "No"
        ws.Range(ws.Cells(fila, COL_K_FECHA), ws.Cells(fila, COL_K_APLICAR)).Font.Color = RGB(150, 150, 150)
    End If
End Sub


'------------------------------------------------------------------------------
' LimpiarHojaClasificacion: borra los datos anteriores y pone los desplegables.
'------------------------------------------------------------------------------
Private Sub LimpiarHojaClasificacion(ByVal ws As Worksheet)
    Dim rango As Range

    Set rango = ws.Range(ws.Cells(2, COL_K_FECHA), ws.Cells(ws.Rows.Count, COL_K_ID))
    rango.ClearContents
    rango.Font.ColorIndex = xlColorIndexAutomatic
    ws.Columns(COL_K_ID).Font.Color = RGB(150, 150, 150)
    ws.Columns(COL_K_FECHA).NumberFormat = "dd/mm/yyyy hh:mm"

    With ws.Range(ws.Cells(2, COL_K_APLICAR), ws.Cells(2000, COL_K_APLICAR)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=ListaSiNo"
    End With
End Sub


'------------------------------------------------------------------------------
' FilaParaAplicar: True si la fila tiene ID, carpeta válida y Aplicar = Sí.
'------------------------------------------------------------------------------
Private Function FilaParaAplicar(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    Dim carpeta As String

    carpeta = LimpiarRuta(CStr(ws.Cells(fila, COL_K_CARPETA).Value))
    FilaParaAplicar = Len(Trim$(CStr(ws.Cells(fila, COL_K_ID).Value))) > 0 And _
                      Len(carpeta) > 0 And _
                      LCase$(carpeta) <> LCase$(SIN_CLASIFICAR) And _
                      UCase$(Left$(Trim$(CStr(ws.Cells(fila, COL_K_APLICAR).Value)), 1)) = "S"
End Function


'------------------------------------------------------------------------------
' RutaCompleta: "RRHH\Sindicatos" -> "Clasificacion nueva\RRHH\Sindicatos"
' (si ya empieza por "Clasificacion nueva", no lo repite).
'------------------------------------------------------------------------------
Private Function RutaCompleta(ByVal carpeta As String) As String
    carpeta = LimpiarRuta(carpeta)
    If LCase$(carpeta) = LCase$(CARPETA_CLASIFICACION) Or _
       LCase$(Left$(carpeta, Len(CARPETA_CLASIFICACION) + 1)) = LCase$(CARPETA_CLASIFICACION & "\") Then
        RutaCompleta = carpeta
    Else
        RutaCompleta = CARPETA_CLASIFICACION & "\" & carpeta
    End If
End Function


'------------------------------------------------------------------------------
' CrearCarpetaConPadres: crea "A\B\C" y también "A" y "A\B" si no existen,
' para que en Notes se vea un árbol completo.
'------------------------------------------------------------------------------
Private Sub CrearCarpetaConPadres(ByVal db As Object, ByVal ruta As String)
    Dim partes As Variant
    Dim parcial As String
    Dim i As Long

    partes = Split(ruta, "\")
    For i = LBound(partes) To UBound(partes)
        If Len(parcial) > 0 Then parcial = parcial & "\"
        parcial = parcial & partes(i)
        If Not ExisteCarpeta(db, parcial) Then db.EnableFolder parcial
    Next i
End Sub


'------------------------------------------------------------------------------
' CumplePalabras: True si el campo indicado del correo contiene ALGUNA de las
' palabras (sin distinguir mayúsculas ni tildes).
'------------------------------------------------------------------------------
Private Function CumplePalabras(ByVal doc As Object, ByVal campo As String, ByVal lista As Variant) As Boolean
    Dim contenido As String
    Dim i As Long
    Dim palabra As String

    Select Case campo
        Case "remitente"
            contenido = TextoCampo(doc, "From") & " " & TextoCampo(doc, "INetFrom")
        Case "cuerpo"
            contenido = TextoCuerpo(doc)
        Case "asunto o cuerpo"
            contenido = TextoCampo(doc, "Subject") & " " & TextoCuerpo(doc)
        Case Else    ' "asunto" o vacío
            contenido = TextoCampo(doc, "Subject")
    End Select
    contenido = Normalizar(contenido)

    For i = LBound(lista) To UBound(lista)
        palabra = Trim$(CStr(lista(i)))
        If Len(palabra) > 0 Then
            If InStr(1, contenido, palabra) > 0 Then
                CumplePalabras = True
                Exit Function
            End If
        End If
    Next i
End Function


'------------------------------------------------------------------------------
' Normalizar: minúsculas y sin tildes, para comparar ("Vacunación" = "vacunacion").
'------------------------------------------------------------------------------
Private Function Normalizar(ByVal texto As String) As String
    Dim conTilde As String
    Dim sinTilde As String
    Dim i As Long

    conTilde = "áàäâéèëêíìïîóòöôúùüûç"
    sinTilde = "aaaaeeeeiiiioooouuuuc"

    texto = LCase$(texto)
    For i = 1 To Len(conTilde)
        texto = Replace(texto, Mid$(conTilde, i, 1), Mid$(sinTilde, i, 1))
    Next i
    Normalizar = texto
End Function


'------------------------------------------------------------------------------
' LeerTextoUtf8: lee un archivo de texto en UTF-8 (con o sin BOM).
' Devuelve el texto con saltos de línea vbLf.
'------------------------------------------------------------------------------
Private Function LeerTextoUtf8(ByVal ruta As String) As String
    Dim flujo As Object
    Dim texto As String

    Set flujo = CreateObject("ADODB.Stream")
    flujo.Type = 2              ' 2 = texto
    flujo.Charset = "utf-8"
    flujo.Open
    flujo.LoadFromFile ruta
    texto = flujo.ReadText
    flujo.Close

    texto = Replace(texto, vbCrLf, vbLf)
    texto = Replace(texto, vbCr, vbLf)
    LeerTextoUtf8 = texto
End Function

Attribute VB_Name = "modLeerCorreo"
'==============================================================================
' MÓDULO: modLeerCorreo
'------------------------------------------------------------------------------
' Lee correos de HCL Notes y los copia en la hoja "Bandeja" con estas columnas:
'   A Fecha | B Remitente | C Asunto | D Extracto | E Nº adjuntos | F ID interno
'   G Carpeta | H Leído | I Mover a | J Resultado
'
' Macro principal: LeerBandejaEntrada  (asígnala a un botón)
'   Pregunta cuántos correos leer y si los quieres en RESUMEN (solo la hoja)
'   o COMPLETOS (la hoja + un .txt con el texto entero de cada correo).
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


' Variables "de módulo": recuerdan tus últimas respuestas mientras el libro
' esté abierto, para no tener que escribirlas cada vez.
Private mUltimoMaximo As Long          ' Último nº de correos que escribiste
Private mUltimaCarpetaTxt As String    ' Última carpeta del PC donde guardaste un .txt


'------------------------------------------------------------------------------
' LeerBandejaEntrada  -> MACRO PRINCIPAL
' Pregunta cuántos correos leer y si los quieres con el contenido completo
' (en ese caso, además de la hoja, crea un .txt donde tú elijas).
'------------------------------------------------------------------------------
Public Sub LeerBandejaEntrada()
    Dim db As Object
    Dim carpetas As New Collection
    Dim maximo As Long
    Dim rutaTxt As String
    Dim contador As Long
    Dim limiteAlcanzado As Boolean

    ' Si ocurre cualquier error, saltamos a la etiqueta "ManejoError".
    On Error GoTo ManejoError

    ' 1) Preguntamos: cuántos, resumen o completo, y dónde guardar el .txt
    If Not PedirOpcionesLectura("Leer bandeja de entrada", MAX_CORREOS, "Bandeja", _
                                maximo, rutaTxt) Then Exit Sub

    ' Desactivar el refresco de pantalla hace que la macro vaya más rápido.
    Application.ScreenUpdating = False
    Application.StatusBar = "Conectando con HCL Notes..."

    ' 2) Comprobamos que la Bandeja de entrada existe en el buzón
    Set db = AbrirBaseCorreo()
    If db.GetView("($Inbox)") Is Nothing Then
        Err.Raise ERR_BASE_CORREO, "LeerBandejaEntrada", _
                  "No se encontró la carpeta Bandeja de entrada en tu buzón."
    End If

    ' 3) Leemos solo la Bandeja de entrada, sin filtros
    carpetas.Add "($Inbox)"
    contador = LeerCorreosEnHoja(carpetas, maximo, "", "", 0, 0, limiteAlcanzado, _
                                 rutaTxt, CARPETA_BANDEJA)

    Application.StatusBar = False
    Application.ScreenUpdating = True
    ObtenerHoja(HOJA_BANDEJA).Activate
    MsgBox "Se han leído " & contador & " correo(s) de la Bandeja de entrada." & _
           TextoFinalTxt(rutaTxt), vbInformation, "Leer correo"

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
' PedirOpcionesLectura: hace las preguntas antes de leer correos.
'   titulo     -> título de las ventanas
'   sugerido   -> nº de correos que se propone (si aún no escribiste ninguno)
'   nombreBase -> parte del nombre propuesto para el .txt (ej. "Bandeja")
'   maximo     -> (de salida) nº de correos a leer
'   rutaTxt    -> (de salida) ruta del .txt, o "" si se eligió "resumen"
' Devuelve False si el usuario cancela en cualquier pregunta.
'------------------------------------------------------------------------------
Public Function PedirOpcionesLectura(ByVal titulo As String, ByVal sugerido As Long, _
                                     ByVal nombreBase As String, _
                                     ByRef maximo As Long, ByRef rutaTxt As String) As Boolean
    Dim respuesta As Variant
    Dim modo As VbMsgBoxResult

    PedirOpcionesLectura = False
    rutaTxt = ""

    ' --- 1) ¿Cuántos correos? ----------------------------------------------------
    If mUltimoMaximo > 0 Then sugerido = mUltimoMaximo

    ' Application.InputBox con Type:=1 solo acepta números y devuelve False
    ' si pulsas Cancelar.
    respuesta = Application.InputBox( _
        Prompt:="¿Cuántos correos quieres leer como máximo?" & vbCrLf & vbCrLf & _
                "(Se empieza por los más recientes.)", _
        Title:=titulo, Default:=sugerido, Type:=1)
    If VarType(respuesta) = vbBoolean Then Exit Function     ' Cancelar

    If respuesta < 1 Or respuesta <> Int(respuesta) Then
        MsgBox "Escribe un número entero mayor que 0.", vbExclamation, titulo
        Exit Function
    End If
    maximo = CLng(respuesta)
    mUltimoMaximo = maximo

    ' --- 2) ¿Resumen o contenido completo? -------------------------------------
    modo = MsgBox("¿Quieres recuperar el CONTENIDO COMPLETO de los correos?" & vbCrLf & vbCrLf & _
                  "  Sí        = completo: la hoja como siempre + un archivo .txt" & vbCrLf & _
                  "                 con el texto entero de cada correo" & vbCrLf & _
                  "  No        = resumen: solo la hoja, como hasta ahora" & vbCrLf & _
                  "  Cancelar = no leer nada", _
                  vbQuestion + vbYesNoCancel + vbDefaultButton2, titulo)
    If modo = vbCancel Then Exit Function

    ' --- 3) Si es completo: ¿dónde guardamos el .txt? --------------------------
    If modo = vbYes Then
        rutaTxt = ElegirArchivoTxt(nombreBase)
        If Len(rutaTxt) = 0 Then Exit Function               ' Cancelar
    End If

    PedirOpcionesLectura = True
End Function


'------------------------------------------------------------------------------
' ElegirArchivoTxt: muestra la ventana "Guardar como" de Windows para elegir
' el .txt. Devuelve la ruta, o "" si se cancela.
'------------------------------------------------------------------------------
Private Function ElegirArchivoTxt(ByVal nombreBase As String) As String
    Dim ruta As Variant
    Dim carpeta As String
    Dim nombre As String
    Dim c As Variant

    ' Nombre propuesto: Correos_Bandeja_2026-10-16_1030.txt
    For Each c In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|", " ")
        nombreBase = Replace(nombreBase, c, "_")
    Next c
    nombre = "Correos_" & nombreBase & "_" & Format$(Now, "yyyy-mm-dd_hhnn") & ".txt"

    ' Carpeta propuesta: la última que usaste, o la del libro
    carpeta = mUltimaCarpetaTxt
    If Len(carpeta) = 0 Then carpeta = ThisWorkbook.Path
    If Len(carpeta) > 0 Then nombre = carpeta & Application.PathSeparator & nombre

    Do
        ruta = Application.GetSaveAsFilename( _
                   InitialFileName:=nombre, _
                   FileFilter:="Archivo de texto (*.txt), *.txt", _
                   Title:="Elige dónde guardar el contenido completo de los correos")
        If VarType(ruta) = vbBoolean Then                 ' Cancelar
            ElegirArchivoTxt = ""
            Exit Function
        End If

        ' Si ya existe, preguntamos antes de sobrescribirlo
        If Len(Dir$(CStr(ruta))) = 0 Then Exit Do
        If MsgBox("El archivo ya existe:" & vbCrLf & ruta & vbCrLf & vbCrLf & _
                  "¿Quieres reemplazarlo?", vbQuestion + vbYesNo + vbDefaultButton2, _
                  "Guardar .txt") = vbYes Then Exit Do
        nombre = CStr(ruta)
    Loop

    ' Recordamos la carpeta para proponerla la próxima vez
    If InStrRev(CStr(ruta), Application.PathSeparator) > 0 Then
        mUltimaCarpetaTxt = Left$(CStr(ruta), InStrRev(CStr(ruta), Application.PathSeparator) - 1)
    End If
    ElegirArchivoTxt = CStr(ruta)
End Function


'------------------------------------------------------------------------------
' TextoFinalTxt: frase para el mensaje final ("" si no se creó .txt).
'------------------------------------------------------------------------------
Public Function TextoFinalTxt(ByVal rutaTxt As String) As String
    If Len(rutaTxt) > 0 Then
        TextoFinalTxt = vbCrLf & vbCrLf & "Contenido completo guardado en:" & vbCrLf & rutaTxt
    End If
End Function


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
'   rutaTxt   -> (opcional) si se indica, guarda AHÍ el contenido completo
'   origen    -> (opcional) descripción de lo leído, para la cabecera del .txt
'
' Si un correo está en varias de esas carpetas, solo se escribe una vez.
'------------------------------------------------------------------------------
Public Function LeerCorreosEnHoja(ByVal carpetas As Collection, ByVal maximo As Long, _
                                  ByVal texto As String, ByVal remitente As String, _
                                  ByVal desde As Date, ByVal hasta As Date, _
                                  ByRef limiteAlcanzado As Boolean, _
                                  Optional ByVal rutaTxt As String = "", _
                                  Optional ByVal origen As String = "") As Long
    Dim sesion As Object
    Dim db As Object
    Dim vista As Object
    Dim doc As Object
    Dim ws As Worksheet
    Dim vistos As Object          ' Diccionario con los correos ya revisados
    Dim bloques As Object         ' Diccionario: ID del correo -> su bloque de texto
    Dim carpeta As Variant
    Dim unid As String
    Dim fila As Long
    Dim contador As Long
    Dim revisados As Long
    Dim completo As Boolean
    Dim mimeAntes As Boolean
    Dim numError As Long
    Dim descError As String

    completo = (Len(rutaTxt) > 0)

    Set ws = PrepararHojaBandeja()
    Set sesion = ObtenerSesion()
    Set db = AbrirBaseCorreo()

    ' La columna "Mover a" necesita la lista de carpetas. Si todavía no se ha
    ' creado (primera vez), la creamos ahora.
    AsegurarListaCarpetas db

    LimpiarHojaBandeja ws

    ' Los correos en formato web (MIME/HTML) a veces devuelven el cuerpo vacío.
    ' Con ConvertMime = False Notes los convierte a texto al leerlos.
    ' Al terminar lo dejamos como estaba.
    On Error Resume Next
    mimeAntes = sesion.ConvertMime
    sesion.ConvertMime = False
    On Error GoTo Fallo

    ' Un "Dictionary" es como una lista con buscador: nos dice muy rápido si
    ' un correo ya ha salido (por si está en dos carpetas a la vez).
    Set vistos = CreateObject("Scripting.Dictionary")
    Set bloques = CreateObject("Scripting.Dictionary")

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
        On Error GoTo Fallo

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
                        If completo Then
                            bloques.Add unid, BloqueCorreo(sesion, doc, NombreVisible(CStr(carpeta)))
                        End If
                        fila = fila + 1
                        contador = contador + 1
                    End If
                End If

                ' Pasamos al correo anterior de la carpeta
                Set doc = vista.GetPrevDocument(doc)
            Loop
        End If
    Next carpeta

    ' Ordena la hoja (más recientes arriba) y pone los desplegables
    FormatearHojaBandeja ws, fila - 1

    ' El .txt sigue el MISMO orden que la hoja ya ordenada
    If completo Then
        Application.StatusBar = "Guardando el archivo .txt..."
        GuardarTxt rutaTxt, sesion, ws, bloques, contador, origen, limiteAlcanzado
    End If

    On Error Resume Next
    sesion.ConvertMime = mimeAntes
    On Error GoTo 0

    LeerCorreosEnHoja = contador
    Exit Function

Fallo:
    ' Guardamos el error, dejamos ConvertMime como estaba y lo "relanzamos"
    ' para que lo muestre la macro que nos llamó.
    numError = Err.Number
    descError = Err.Description
    On Error Resume Next
    sesion.ConvertMime = mimeAntes
    On Error GoTo 0
    Err.Raise numError, "LeerCorreosEnHoja", descError
End Function


'------------------------------------------------------------------------------
' BloqueCorreo: el texto de UN correo para el .txt (sin la línea de inicio y
' la de fin, que se añaden al guardar porque llevan el número de orden).
'------------------------------------------------------------------------------
Private Function BloqueCorreo(ByVal sesion As Object, ByVal doc As Object, _
                              ByVal carpetaVisible As String) As String
    Dim fecha As Variant
    Dim adjuntos As String
    Dim nombres As Variant
    Dim numAdjuntos As Long
    Dim i As Long
    Dim t As String

    fecha = FechaDelCorreo(doc)

    nombres = NombresAdjuntos(sesion, doc)
    For i = LBound(nombres) To UBound(nombres)
        If Len(CStr(nombres(i))) > 0 Then
            If numAdjuntos > 0 Then adjuntos = adjuntos & "; "
            adjuntos = adjuntos & CStr(nombres(i))
            numAdjuntos = numAdjuntos + 1
        End If
    Next i

    t = "ID: " & doc.UniversalID & vbCrLf
    t = t & "Fecha: " & IIf(IsDate(fecha), Format$(fecha, "yyyy-mm-dd hh:nn"), "") & vbCrLf
    t = t & "De: " & UnaLinea(TextoRemitente(sesion, doc)) & vbCrLf
    ' Para y CC NO se exportan: suelen ser listas muy largas de destinatarios
    t = t & "Asunto: " & UnaLinea(TextoCampo(doc, "Subject")) & vbCrLf
    t = t & "Carpeta: " & carpetaVisible & vbCrLf
    t = t & "Leído: " & TextoLeido(doc) & vbCrLf
    t = t & "Adjuntos (" & numAdjuntos & "): " & adjuntos & vbCrLf
    t = t & "--- CUERPO ---" & vbCrLf
    t = t & ProtegerCuerpo(TextoCuerpo(doc)) & vbCrLf
    t = t & "--- FIN CUERPO ---" & vbCrLf

    BloqueCorreo = t
End Function


'------------------------------------------------------------------------------
' GuardarTxt: escribe el .txt (en UTF-8, para que tildes y ñ salgan bien).
' Recorre la hoja ya ordenada para que el .txt tenga el mismo orden.
'------------------------------------------------------------------------------
Private Sub GuardarTxt(ByVal rutaTxt As String, ByVal sesion As Object, ByVal ws As Worksheet, _
                       ByVal bloques As Object, ByVal total As Long, ByVal origen As String, _
                       ByVal limiteAlcanzado As Boolean)
    Dim flujo As Object        ' ADODB.Stream: permite escribir texto en UTF-8
    Dim usuario As String
    Dim unid As String
    Dim fila As Long
    Dim n As Long

    On Error Resume Next
    usuario = sesion.CommonUserName
    On Error GoTo 0

    Set flujo = CreateObject("ADODB.Stream")
    flujo.Type = 2              ' 2 = texto
    flujo.Charset = "utf-8"
    flujo.Open

    ' --- Cabecera ---------------------------------------------------------------
    flujo.WriteText "=== EXPORTACIÓN HCL NOTES ===" & vbCrLf
    flujo.WriteText "Fecha exportación: " & Format$(Now, "yyyy-mm-dd hh:nn") & vbCrLf
    flujo.WriteText "Usuario: " & usuario & vbCrLf
    flujo.WriteText "Origen: " & origen & vbCrLf
    flujo.WriteText "Nº correos: " & total & vbCrLf
    flujo.WriteText "Límite alcanzado: " & IIf(limiteAlcanzado, "Sí (hay más correos que no se han exportado)", "No") & vbCrLf
    flujo.WriteText "Formato: cada correo va entre '### CORREO n DE N ###' y '### FIN CORREO n ###'." & vbCrLf
    flujo.WriteText "         El cuerpo va entre '--- CUERPO ---' y '--- FIN CUERPO ---'." & vbCrLf
    flujo.WriteText "         Fechas en formato AAAA-MM-DD hh:mm. Orden: del más reciente al más antiguo." & vbCrLf
    flujo.WriteText "==============================" & vbCrLf & vbCrLf

    ' --- Un bloque por correo, en el orden de la hoja ------------------------------
    For fila = 2 To total + 1
        unid = CStr(ws.Cells(fila, COL_B_ID).Value)
        If bloques.Exists(unid) Then
            n = n + 1
            flujo.WriteText "### CORREO " & n & " DE " & total & " ###" & vbCrLf
            flujo.WriteText bloques(unid)
            flujo.WriteText "### FIN CORREO " & n & " ###" & vbCrLf & vbCrLf
        End If
    Next fila

    flujo.SaveToFile rutaTxt, 2   ' 2 = sobrescribir si existe
    flujo.Close
End Sub


'------------------------------------------------------------------------------
' TextoRemitente: "Ana López <ana@acme.com>" (o lo que se pueda saber).
'------------------------------------------------------------------------------
Private Function TextoRemitente(ByVal sesion As Object, ByVal doc As Object) As String
    Dim nombre As String
    Dim direccion As String

    nombre = NombreLegible(sesion, ValorCampo(doc, "From"))
    direccion = Trim$(CStr(ValorCampo(doc, "INetFrom")))   ' Dirección de Internet

    If Len(direccion) > 0 And InStr(direccion, "<") > 0 Then
        TextoRemitente = direccion          ' Ya viene como "Nombre <dirección>"
    ElseIf Len(direccion) > 0 And LCase$(direccion) <> LCase$(nombre) Then
        TextoRemitente = nombre & " <" & direccion & ">"
    Else
        TextoRemitente = nombre
    End If
End Function


'------------------------------------------------------------------------------
' UnaLinea: quita los saltos de línea (para los campos de cabecera).
'------------------------------------------------------------------------------
Private Function UnaLinea(ByVal texto As String) As String
    texto = Replace(texto, vbCrLf, " ")
    texto = Replace(texto, vbCr, " ")
    texto = Replace(texto, vbLf, " ")
    UnaLinea = Trim$(texto)
End Function


'------------------------------------------------------------------------------
' ProtegerCuerpo: deja los saltos de línea como Windows (CRLF) y, si alguna
' línea del cuerpo empezara como nuestras marcas ("### " o "--- "), le pone
' un espacio delante para que no se confunda con el final del bloque.
'------------------------------------------------------------------------------
Private Function ProtegerCuerpo(ByVal texto As String) As String
    texto = Replace(texto, vbCrLf, vbLf)
    texto = Replace(texto, vbCr, vbLf)
    texto = vbLf & texto
    texto = Replace(texto, vbLf & "### ", vbLf & " ### ")
    texto = Replace(texto, vbLf & "--- ", vbLf & " --- ")
    texto = Mid$(texto, 2)
    ProtegerCuerpo = Replace(texto, vbLf, vbCrLf)
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

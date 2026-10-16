Attribute VB_Name = "modBuscar"
'==============================================================================
' MÓDULO: modBuscar
'------------------------------------------------------------------------------
' Lee carpetas completas y busca correos. El resultado siempre aparece en la
' hoja "Bandeja", así luego puedes moverlos, guardar sus adjuntos, etc.
'
'   LeerCarpeta     -> lee la carpeta seleccionada en la hoja "Carpetas"
'                      (con o sin sus subcarpetas, y con filtro de días)
'   BuscarCorreos   -> busca con los criterios escritos en la hoja "Buscar"
'   LimpiarBusqueda -> borra los criterios de la hoja "Buscar"
'
' Antes de leer se pregunta cuántos correos como máximo (se propone
' MAX_CORREOS_CARPETA) y si se quieren en resumen o con el contenido completo
' en un .txt (ver PedirOpcionesLectura en modLeerCorreo).
'==============================================================================
Option Explicit


'------------------------------------------------------------------------------
' LeerCarpeta  -> MACRO (botón "Leer carpeta" de la hoja Carpetas)
'------------------------------------------------------------------------------
Public Sub LeerCarpeta()
    Dim db As Object
    Dim ruta As String
    Dim carpetas As New Collection
    Dim subcarpetas As Collection
    Dim respuesta As VbMsgBoxResult
    Dim dias As String
    Dim desde As Date
    Dim contador As Long
    Dim limiteAlcanzado As Boolean
    Dim maximo As Long
    Dim rutaTxt As String
    Dim origen As String
    Dim s As Variant

    On Error GoTo ManejoError

    ruta = CarpetaSeleccionada()
    If Len(ruta) = 0 Then
        MsgBox "Ve a la hoja '" & HOJA_CARPETAS & "', haz clic en la fila de la carpeta " & _
               "que quieres leer y vuelve a pulsar 'Leer carpeta'." & vbCrLf & vbCrLf & _
               "(Si la hoja está vacía, pulsa 'Actualizar carpetas'.)", vbInformation, "Leer carpeta"
        Exit Sub
    End If

    Set db = AbrirBaseCorreo()
    carpetas.Add NombreInterno(ruta)

    ' ¿Incluimos las subcarpetas?
    Set subcarpetas = ListaSubcarpetas(db, ruta)
    If subcarpetas.Count > 0 Then
        respuesta = MsgBox("La carpeta '" & ruta & "' tiene " & subcarpetas.Count & " subcarpeta(s)." & _
                           vbCrLf & vbCrLf & "¿Quieres leer también sus correos?", _
                           vbQuestion + vbYesNoCancel, "Leer carpeta")
        If respuesta = vbCancel Then Exit Sub
        If respuesta = vbYes Then
            For Each s In subcarpetas
                carpetas.Add CStr(s)
            Next s
        End If
    End If

    ' ¿Filtro de fechas?
    dias = Trim$(InputBox("¿Quieres solo los correos de los últimos días?" & vbCrLf & vbCrLf & _
                          "Escribe el número de días (ej. 30)" & vbCrLf & _
                          "o déjalo vacío para leerlos todos.", "Leer carpeta"))
    If Len(dias) > 0 Then
        If Not IsNumeric(dias) Then
            MsgBox "'" & dias & "' no es un número.", vbExclamation, "Leer carpeta"
            Exit Sub
        End If
        desde = Date - CLng(dias)
    End If

    ' ¿Cuántos, resumen o completo, y dónde guardar el .txt?
    If Not PedirOpcionesLectura("Leer carpeta", MAX_CORREOS_CARPETA, ruta, maximo, rutaTxt) Then Exit Sub

    ' Descripción para la cabecera del .txt
    origen = "Carpeta '" & ruta & "'" & IIf(carpetas.Count > 1, " y sus subcarpetas", "") & _
             IIf(desde <> 0, " (desde " & Format$(desde, "yyyy-mm-dd") & ")", "")

    Application.ScreenUpdating = False
    Application.StatusBar = "Leyendo la carpeta '" & ruta & "'..."

    contador = LeerCorreosEnHoja(carpetas, maximo, "", "", desde, 0, limiteAlcanzado, rutaTxt, origen)

    Application.StatusBar = False
    Application.ScreenUpdating = True
    ObtenerHoja(HOJA_BANDEJA).Activate
    MsgBox "Se han leído " & contador & " correo(s) de '" & ruta & "'" & _
           IIf(carpetas.Count > 1, " y sus subcarpetas", "") & "." & _
           IIf(limiteAlcanzado, vbCrLf & vbCrLf & "Se ha llegado al máximo de " & maximo & _
               " correos que pediste: hay más en la carpeta.", "") & _
           TextoFinalTxt(rutaTxt), _
           vbInformation, "Leer carpeta"

Salir:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Leer carpeta", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' BuscarCorreos  -> MACRO (botón "Buscar" de la hoja Buscar)
'------------------------------------------------------------------------------
Public Sub BuscarCorreos()
    Dim ws As Worksheet
    Dim db As Object
    Dim carpetas As New Collection
    Dim texto As String
    Dim remitente As String
    Dim carpeta As String
    Dim conSubcarpetas As Boolean
    Dim desde As Date
    Dim hasta As Date
    Dim contador As Long
    Dim limiteAlcanzado As Boolean
    Dim maximo As Long
    Dim rutaTxt As String
    Dim origen As String
    Dim s As Variant

    On Error GoTo ManejoError

    Set ws = PrepararHojaBuscar()

    ' 1) Leemos los criterios de la hoja
    texto = Trim$(CStr(ws.Cells(FILA_BUS_TEXTO, 2).Value))
    remitente = Trim$(CStr(ws.Cells(FILA_BUS_REMITENTE, 2).Value))
    desde = LeerFecha(ws.Cells(FILA_BUS_DESDE, 2), "Desde")
    hasta = LeerFecha(ws.Cells(FILA_BUS_HASTA, 2), "Hasta")
    carpeta = Trim$(CStr(ws.Cells(FILA_BUS_CARPETA, 2).Value))
    conSubcarpetas = (UCase$(Left$(Trim$(CStr(ws.Cells(FILA_BUS_SUBCARPETAS, 2).Value)), 1)) <> "N")

    If Len(texto) = 0 And Len(remitente) = 0 And desde = 0 And hasta = 0 And Len(carpeta) = 0 Then
        MsgBox "Rellena al menos un criterio de búsqueda (columna B).", vbInformation, "Buscar correos"
        Exit Sub
    End If

    If desde <> 0 And hasta <> 0 And desde > hasta Then
        MsgBox "La fecha 'Desde' es posterior a la fecha 'Hasta'.", vbExclamation, "Buscar correos"
        Exit Sub
    End If

    ' ¿Cuántos, resumen o completo, y dónde guardar el .txt?
    If Not PedirOpcionesLectura("Buscar correos", MAX_CORREOS_CARPETA, "Busqueda", maximo, rutaTxt) Then Exit Sub

    ' Descripción de la búsqueda para la cabecera del .txt
    origen = "Búsqueda -> texto: '" & texto & "', remitente: '" & remitente & "', desde: " & _
             IIf(desde <> 0, Format$(desde, "yyyy-mm-dd"), "-") & ", hasta: " & _
             IIf(hasta <> 0, Format$(hasta, "yyyy-mm-dd"), "-") & ", carpeta: " & _
             IIf(Len(carpeta) > 0, "'" & carpeta & "'" & IIf(conSubcarpetas, " y subcarpetas", ""), "todas")

    Application.ScreenUpdating = False
    Application.StatusBar = "Conectando con HCL Notes..."
    Set db = AbrirBaseCorreo()

    ' 2) ¿En qué carpetas buscamos?
    If Len(carpeta) = 0 Then
        ' En todas: Bandeja de entrada, Enviados y todas tus carpetas
        carpetas.Add "($Inbox)"
        carpetas.Add "($Sent)"
        For Each s In ListaCarpetasUsuario(db)
            carpetas.Add CStr(s)
        Next s
    Else
        carpetas.Add NombreInterno(carpeta)
        If conSubcarpetas Then
            For Each s In ListaSubcarpetas(db, carpeta)
                carpetas.Add CStr(s)
            Next s
        End If
    End If

    ' 3) Buscamos
    contador = LeerCorreosEnHoja(carpetas, maximo, texto, remitente, desde, hasta, limiteAlcanzado, _
                                 rutaTxt, origen)

    Application.StatusBar = False
    Application.ScreenUpdating = True
    ObtenerHoja(HOJA_BANDEJA).Activate
    MsgBox "Se han encontrado " & contador & " correo(s)." & _
           IIf(limiteAlcanzado, vbCrLf & vbCrLf & "Se ha llegado al máximo de " & maximo & _
               " correos que pediste: puede haber más resultados.", "") & _
           TextoFinalTxt(rutaTxt), _
           vbInformation, "Buscar correos"

Salir:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Buscar correos", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' LimpiarBusqueda  -> MACRO (botón "Limpiar criterios")
'------------------------------------------------------------------------------
Public Sub LimpiarBusqueda()
    Dim ws As Worksheet

    Set ws = PrepararHojaBuscar()
    ws.Range(ws.Cells(FILA_BUS_TEXTO, 2), ws.Cells(FILA_BUS_CARPETA, 2)).ClearContents
    ws.Cells(FILA_BUS_SUBCARPETAS, 2).Value = "Sí"
    ws.Cells(FILA_BUS_TEXTO, 2).Select
End Sub


'------------------------------------------------------------------------------
' LeerFecha: devuelve la fecha de la celda, o 0 si está vacía.
' Si lo escrito no es una fecha, da un error comprensible.
'------------------------------------------------------------------------------
Private Function LeerFecha(ByVal celda As Range, ByVal nombre As String) As Date
    If Len(Trim$(CStr(celda.Value))) = 0 Then
        LeerFecha = 0
    ElseIf IsDate(celda.Value) Then
        LeerFecha = Int(CDate(celda.Value))     ' Int = sin la hora
    Else
        Err.Raise ERR_DATOS, "BuscarCorreos", _
                  "La fecha '" & nombre & "' no es válida: " & CStr(celda.Value) & vbCrLf & _
                  "Escríbela así: 15/10/2026"
    End If
End Function

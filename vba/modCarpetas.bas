Attribute VB_Name = "modCarpetas"
'==============================================================================
' MÓDULO: modCarpetas
'------------------------------------------------------------------------------
' Gestiona las CARPETAS de tu buzón de HCL Notes desde la hoja "Carpetas":
'
'   ActualizarCarpetas    -> lista todas tus carpetas en forma de árbol
'   CrearCarpeta          -> crea una carpeta (o subcarpeta de la seleccionada)
'   RenombrarCarpeta      -> cambia el nombre de la carpeta seleccionada
'   EliminarCarpeta       -> elimina la carpeta seleccionada (y sus subcarpetas)
'   CrearCarpetasDeLista  -> crea de golpe las carpetas escritas en la columna G
'
' ¿CÓMO SON LAS SUBCARPETAS EN NOTES?
'   Notes guarda la ruta completa en el nombre, separando niveles con "\".
'   Ejemplo: la carpeta "2026" dentro de "Acme" dentro de "Clientes" se llama
'   "Clientes\Acme\2026".
'
' ¿QUÉ ES UNA "AGRUPACIÓN"?
'   Si existe "Clientes\Acme" pero NO existe "Clientes" como carpeta, Notes
'   igualmente la muestra como "padre". En la hoja aparece con tipo
'   "Agrupación" (no contiene correos propios).
'==============================================================================
Option Explicit


'------------------------------------------------------------------------------
' ActualizarCarpetas  -> MACRO (botón "Actualizar carpetas")
'------------------------------------------------------------------------------
Public Sub ActualizarCarpetas()
    Dim total As Long

    On Error GoTo ManejoError

    Application.ScreenUpdating = False
    Application.StatusBar = "Leyendo las carpetas de HCL Notes..."

    total = RellenarHojaCarpetas(AbrirBaseCorreo())

    Application.StatusBar = False
    Application.ScreenUpdating = True
    ObtenerHoja(HOJA_CARPETAS).Activate
    MsgBox "Se han encontrado " & total & " carpeta(s)." & vbCrLf & vbCrLf & _
           "Selecciona una fila y usa los botones de la derecha para leerla, " & _
           "renombrarla, eliminarla o crear subcarpetas dentro.", _
           vbInformation, "Carpetas"

Salir:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Actualizar carpetas", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' CrearCarpeta  -> MACRO (botón "Crear carpeta / subcarpeta")
' Si hay una carpeta seleccionada en la hoja Carpetas, la nueva se crea
' DENTRO de ella. Si no, se crea en el primer nivel.
'------------------------------------------------------------------------------
Public Sub CrearCarpeta()
    Dim db As Object
    Dim padre As String
    Dim nombre As String
    Dim ruta As String
    Dim donde As String
    Dim motivo As String

    On Error GoTo ManejoError

    padre = CarpetaSeleccionada()
    ' Dentro de "Bandeja de entrada" o "Enviados" no se pueden crear subcarpetas
    If EsCarpetaSistema(padre) Then padre = ""

    If Len(padre) > 0 Then
        donde = "Se creará DENTRO de:  " & padre
    Else
        donde = "Se creará en el primer nivel." & vbCrLf & _
                "(Para crear una subcarpeta, selecciona antes la carpeta padre en la hoja '" & _
                HOJA_CARPETAS & "'.)"
    End If

    nombre = InputBox("Nombre de la nueva carpeta:" & vbCrLf & vbCrLf & donde & vbCrLf & vbCrLf & _
                      "Puedes crear varios niveles a la vez separándolos con \" & vbCrLf & _
                      "(ejemplo:  Facturas\2026)", "Crear carpeta")
    nombre = LimpiarRuta(nombre)
    If Len(nombre) = 0 Then Exit Sub        ' Cancelado o vacío

    If Len(padre) > 0 Then
        ruta = padre & "\" & nombre
    Else
        ruta = nombre
    End If

    If Not RutaValida(ruta, motivo) Then
        MsgBox motivo, vbExclamation, "Crear carpeta"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set db = AbrirBaseCorreo()

    If ExisteCarpeta(db, ruta) Then
        Application.ScreenUpdating = True
        MsgBox "Ya existe una carpeta llamada:" & vbCrLf & ruta, vbExclamation, "Crear carpeta"
        Exit Sub
    End If

    ' EnableFolder crea la carpeta si no existe
    db.EnableFolder ruta

    RellenarHojaCarpetas db
    SeleccionarCarpetaEnHoja ruta
    Application.ScreenUpdating = True
    MsgBox "Carpeta creada:" & vbCrLf & ruta, vbInformation, "Crear carpeta"

Salir:
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Crear carpeta", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' RenombrarCarpeta  -> MACRO (botón "Renombrar carpeta")
' Cambia el nombre de la carpeta seleccionada. Sus subcarpetas se renombran
' también para que sigan dentro de ella. Los correos no se tocan.
'------------------------------------------------------------------------------
Public Sub RenombrarCarpeta()
    Dim db As Object
    Dim vista As Object
    Dim subcarpetas As Collection
    Dim actual As String
    Dim nuevoNombre As String
    Dim rutaNueva As String
    Dim padre As String
    Dim motivo As String
    Dim s As Variant

    On Error GoTo ManejoError

    actual = CarpetaSeleccionada()
    If Not ComprobarSeleccion(actual, "Renombrar carpeta") Then Exit Sub

    nuevoNombre = InputBox("Nuevo nombre para la carpeta:" & vbCrLf & vbCrLf & actual & vbCrLf & vbCrLf & _
                           "Escribe solo el nombre (sin \).", "Renombrar carpeta", UltimoNivel(actual))
    nuevoNombre = LimpiarRuta(nuevoNombre)
    If Len(nuevoNombre) = 0 Or nuevoNombre = UltimoNivel(actual) Then Exit Sub

    If InStr(nuevoNombre, "\") > 0 Then
        MsgBox "Escribe solo el nombre nuevo, sin el carácter \.", vbExclamation, "Renombrar carpeta"
        Exit Sub
    End If

    padre = RutaPadre(actual)
    If Len(padre) > 0 Then
        rutaNueva = padre & "\" & nuevoNombre
    Else
        rutaNueva = nuevoNombre
    End If

    If Not RutaValida(rutaNueva, motivo) Then
        MsgBox motivo, vbExclamation, "Renombrar carpeta"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Set db = AbrirBaseCorreo()

    If ExisteCarpeta(db, rutaNueva) Then
        Application.ScreenUpdating = True
        MsgBox "Ya existe una carpeta llamada:" & vbCrLf & rutaNueva, vbExclamation, "Renombrar carpeta"
        Exit Sub
    End If

    Set subcarpetas = ListaSubcarpetas(db, actual)

    ' 1) La carpeta en sí (si es una "agrupación" no existe como tal)
    Set vista = db.GetView(actual)
    If Not vista Is Nothing Then vista.Name = rutaNueva

    ' 2) Sus subcarpetas: cambiamos el principio de la ruta
    '    "Clientes\Acme\2026" -> "Clientes\NuevoNombre\2026"
    For Each s In subcarpetas
        Set vista = db.GetView(CStr(s))
        If Not vista Is Nothing Then
            vista.Name = rutaNueva & Mid$(CStr(s), Len(actual) + 1)
        End If
    Next s

    RellenarHojaCarpetas db
    SeleccionarCarpetaEnHoja rutaNueva
    Application.ScreenUpdating = True
    MsgBox "Carpeta renombrada:" & vbCrLf & actual & vbCrLf & "  ->  " & rutaNueva & _
           IIf(subcarpetas.Count > 0, vbCrLf & vbCrLf & "(también sus " & subcarpetas.Count & " subcarpeta(s))", ""), _
           vbInformation, "Renombrar carpeta"

Salir:
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Renombrar carpeta", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' EliminarCarpeta  -> MACRO (botón "Eliminar carpeta")
' Elimina la carpeta seleccionada y TODAS sus subcarpetas.
' Los correos NO se borran: puedes elegir pasarlos antes a la Bandeja de
' entrada. Si no, seguirán en la vista "Todos los documentos" de Notes.
'------------------------------------------------------------------------------
Public Sub EliminarCarpeta()
    Dim db As Object
    Dim vista As Object
    Dim subcarpetas As Collection
    Dim todas As New Collection     ' La carpeta + sus subcarpetas
    Dim actual As String
    Dim respuesta As VbMsgBoxResult
    Dim totalCorreos As Long
    Dim movidos As Long
    Dim i As Long
    Dim s As Variant

    On Error GoTo ManejoError

    actual = CarpetaSeleccionada()
    If Not ComprobarSeleccion(actual, "Eliminar carpeta") Then Exit Sub

    Application.ScreenUpdating = False
    Set db = AbrirBaseCorreo()

    Set subcarpetas = ListaSubcarpetas(db, actual)
    If ExisteCarpeta(db, actual) Then todas.Add actual
    For Each s In subcarpetas
        todas.Add CStr(s)
    Next s

    If todas.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "La carpeta '" & actual & "' ya no existe." & vbCrLf & _
               "Pulsa 'Actualizar carpetas'.", vbExclamation, "Eliminar carpeta"
        Exit Sub
    End If

    For Each s In todas
        totalCorreos = totalCorreos + NumeroCorreos(db, CStr(s))
    Next s
    Application.ScreenUpdating = True

    ' --- Primera pregunta: ¿qué hacemos con los correos? ----------------------
    respuesta = MsgBox("Vas a ELIMINAR la carpeta:" & vbCrLf & "   " & actual & vbCrLf & _
        IIf(subcarpetas.Count > 0, "y sus " & subcarpetas.Count & " subcarpeta(s)." & vbCrLf, "") & vbCrLf & _
        "Contienen " & totalCorreos & " correo(s). Los correos NO se borran." & vbCrLf & vbCrLf & _
        "¿Quieres pasar esos correos a la Bandeja de entrada antes de eliminar?" & vbCrLf & vbCrLf & _
        "  Sí        = pasarlos a la Bandeja de entrada y eliminar" & vbCrLf & _
        "  No        = eliminar sin moverlos (quedarán en 'Todos los documentos')" & vbCrLf & _
        "  Cancelar = no hacer nada", _
        vbYesNoCancel + vbExclamation + vbDefaultButton3, "Eliminar carpeta")
    If respuesta = vbCancel Then Exit Sub

    ' --- Segunda pregunta: confirmación final ----------------------------------
    If MsgBox("¿Seguro? Eliminar una carpeta NO se puede deshacer.", _
              vbYesNo + vbCritical + vbDefaultButton2, "Eliminar carpeta") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    If respuesta = vbYes Then
        For Each s In todas
            Application.StatusBar = "Pasando correos de '" & s & "' a la Bandeja de entrada..."
            movidos = movidos + PasarCorreosABandeja(db, CStr(s))
        Next s
    End If

    ' Borramos de la más profunda a la menos profunda (de atrás hacia delante,
    ' porque la lista está ordenada con los padres antes que los hijos)
    For i = todas.Count To 1 Step -1
        Application.StatusBar = "Eliminando '" & todas(i) & "'..."
        Set vista = db.GetView(CStr(todas(i)))
        If Not vista Is Nothing Then vista.Remove
    Next i

    RellenarHojaCarpetas db
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Carpeta eliminada: " & actual & _
           IIf(respuesta = vbYes, vbCrLf & movidos & " correo(s) pasados a la Bandeja de entrada.", ""), _
           vbInformation, "Eliminar carpeta"

Salir:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Eliminar carpeta", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' CrearCarpetasDeLista  -> MACRO (botón "Crear carpetas de la lista")
' Crea todas las carpetas escritas en la columna G de la hoja Carpetas
' (una ruta por fila, ej. "Clientes\Acme\2026") y escribe el resultado en H.
'------------------------------------------------------------------------------
Public Sub CrearCarpetasDeLista()
    Dim ws As Worksheet
    Dim db As Object
    Dim ultimaFila As Long
    Dim fila As Long
    Dim ruta As String
    Dim motivo As String
    Dim creadas As Long
    Dim pendientes As Long

    On Error GoTo ManejoError

    Set ws = PrepararHojaCarpetas()
    ultimaFila = ws.Cells(ws.Rows.Count, COL_C_NUEVA).End(xlUp).Row

    For fila = 2 To ultimaFila
        If Len(LimpiarRuta(CStr(ws.Cells(fila, COL_C_NUEVA).Value))) > 0 Then pendientes = pendientes + 1
    Next fila

    If pendientes = 0 Then
        MsgBox "Escribe en la columna G de la hoja '" & HOJA_CARPETAS & "' las carpetas " & _
               "que quieres crear, una por fila. Ejemplo:" & vbCrLf & vbCrLf & _
               "   Clientes\Acme" & vbCrLf & "   Clientes\Acme\2026" & vbCrLf & "   Facturas", _
               vbInformation, "Crear carpetas de la lista"
        Exit Sub
    End If

    If MsgBox("Se van a crear hasta " & pendientes & " carpeta(s). ¿Continuar?", _
              vbQuestion + vbYesNo, "Crear carpetas de la lista") = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Set db = AbrirBaseCorreo()

    For fila = 2 To ultimaFila
        ruta = LimpiarRuta(CStr(ws.Cells(fila, COL_C_NUEVA).Value))
        If Len(ruta) > 0 Then
            If Not RutaValida(ruta, motivo) Then
                ws.Cells(fila, COL_C_NUEVA_RES).Value = "ERROR: " & motivo
                ws.Cells(fila, COL_C_NUEVA_RES).Font.Color = RGB(192, 0, 0)
            ElseIf ExisteCarpeta(db, ruta) Then
                ws.Cells(fila, COL_C_NUEVA_RES).Value = "Ya existía"
                ws.Cells(fila, COL_C_NUEVA_RES).Font.Color = RGB(128, 128, 128)
            Else
                db.EnableFolder ruta
                creadas = creadas + 1
                ws.Cells(fila, COL_C_NUEVA_RES).Value = "Creada " & Format$(Now, "dd/mm/yyyy hh:mm")
                ws.Cells(fila, COL_C_NUEVA_RES).Font.Color = RGB(0, 128, 0)
            End If
        End If
    Next fila

    RellenarHojaCarpetas db
    Application.ScreenUpdating = True
    MsgBox "Se han creado " & creadas & " carpeta(s). Mira la columna 'Resultado'.", _
           vbInformation, "Crear carpetas de la lista"

Salir:
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Crear carpetas de la lista", Err.Number, Err.Description
    Resume Salir
End Sub


'==============================================================================
' FUNCIONES QUE USAN OTROS MÓDULOS
'==============================================================================

'------------------------------------------------------------------------------
' RellenarHojaCarpetas: escribe el árbol de carpetas en la hoja Carpetas,
' actualiza la lista de los desplegables y devuelve el nº de carpetas.
'------------------------------------------------------------------------------
Public Function RellenarHojaCarpetas(ByVal db As Object) As Long
    Dim ws As Worksheet
    Dim escritas As Object        ' Diccionario: rutas ya escritas en la hoja
    Dim carpeta As Variant
    Dim partes As Variant
    Dim ruta As String
    Dim fila As Long
    Dim i As Long

    Set ws = PrepararHojaCarpetas()
    Set escritas = CreateObject("Scripting.Dictionary")

    ' Borramos el árbol anterior (columnas A a E; la lista de la columna G no)
    With ws.Range(ws.Cells(2, COL_C_NOMBRE), ws.Cells(ws.Rows.Count, COL_C_TIPO))
        .ClearContents
        .Font.Bold = False
        .Font.ColorIndex = xlColorIndexAutomatic
        .IndentLevel = 0
    End With

    fila = 2

    ' 1) Carpetas del sistema
    EscribirFilaCarpeta ws, fila, db, CARPETA_BANDEJA, "Sistema"
    EscribirFilaCarpeta ws, fila, db, CARPETA_ENVIADOS, "Sistema"

    ' 2) Tus carpetas, en orden de árbol
    For Each carpeta In ListaCarpetasUsuario(db)
        ' Si falta algún "padre" (agrupación), lo escribimos antes
        partes = Split(CStr(carpeta), "\")
        ruta = ""
        For i = LBound(partes) To UBound(partes) - 1
            If Len(ruta) > 0 Then ruta = ruta & "\"
            ruta = ruta & partes(i)
            If Not escritas.Exists(LCase$(ruta)) Then
                escritas.Add LCase$(ruta), True
                EscribirFilaCarpeta ws, fila, db, ruta, "Agrupación"
            End If
        Next i

        If Not escritas.Exists(LCase$(CStr(carpeta))) Then
            escritas.Add LCase$(CStr(carpeta)), True
            EscribirFilaCarpeta ws, fila, db, CStr(carpeta), "Carpeta"
        End If
    Next carpeta

    ' 3) El rango con nombre que usan los desplegables: columna B (rutas)
    ThisWorkbook.Names.Add Name:=NOMBRE_LISTA_CARPETAS, _
        RefersTo:="='" & ws.Name & "'!$B$2:$B$" & (fila - 1)

    ' 4) Ponemos el desplegable en las hojas Reglas y Buscar (si existen)
    If HojaExiste(HOJA_REGLAS) Then
        With ThisWorkbook.Worksheets(HOJA_REGLAS)
            AplicarDesplegableCarpetas .Range(.Cells(2, COL_R_DESTINO), .Cells(200, COL_R_DESTINO))
        End With
    End If
    If HojaExiste(HOJA_BUSCAR) Then
        With ThisWorkbook.Worksheets(HOJA_BUSCAR)
            AplicarDesplegableCarpetas .Cells(FILA_BUS_CARPETA, 2)
        End With
    End If

    RellenarHojaCarpetas = fila - 2
End Function


'------------------------------------------------------------------------------
' AsegurarListaCarpetas: si todavía no existe la lista de carpetas para los
' desplegables, la crea leyendo las carpetas de Notes.
'------------------------------------------------------------------------------
Public Sub AsegurarListaCarpetas(ByVal db As Object)
    Dim nombreRango As Object

    On Error Resume Next
    Set nombreRango = ThisWorkbook.Names(NOMBRE_LISTA_CARPETAS)
    On Error GoTo 0

    If nombreRango Is Nothing Then RellenarHojaCarpetas db
End Sub


'------------------------------------------------------------------------------
' AplicarDesplegableCarpetas: pone en las celdas un desplegable con la lista
' de carpetas. Si la lista aún no existe, no hace nada.
'------------------------------------------------------------------------------
Public Sub AplicarDesplegableCarpetas(ByVal rango As Range)
    Dim nombreRango As Object

    On Error Resume Next
    Set nombreRango = ThisWorkbook.Names(NOMBRE_LISTA_CARPETAS)
    On Error GoTo 0
    If nombreRango Is Nothing Then Exit Sub

    With rango.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & NOMBRE_LISTA_CARPETAS
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Carpeta no válida"
        .ErrorMessage = "Elige una carpeta de la lista." & vbCrLf & _
                        "Si no aparece, pulsa 'Actualizar carpetas' en la hoja " & HOJA_CARPETAS & "."
    End With
End Sub


'------------------------------------------------------------------------------
' ListaCarpetasUsuario: tus carpetas (sin las del sistema), ordenadas en
' forma de árbol: cada carpeta va justo antes de sus subcarpetas.
'------------------------------------------------------------------------------
Public Function ListaCarpetasUsuario(ByVal db As Object) As Collection
    Dim resultado As New Collection
    Dim vistas As Variant
    Dim v As Variant
    Dim nombres() As String
    Dim nombre As String
    Dim esCarpeta As Boolean
    Dim n As Long
    Dim i As Long

    vistas = db.Views     ' Todas las vistas y carpetas del buzón
    If Not IsArray(vistas) Then
        Set ListaCarpetasUsuario = resultado
        Exit Function
    End If

    ReDim nombres(0 To UBound(vistas) - LBound(vistas))

    For Each v In vistas
        esCarpeta = False
        nombre = ""
        On Error Resume Next
        esCarpeta = v.IsFolder
        nombre = v.Name
        On Error GoTo 0

        ' Las carpetas "internas" de Notes empiezan por "(" -> las saltamos
        If esCarpeta And Len(nombre) > 0 Then
            If Left$(nombre, 1) <> "(" Then
                nombres(n) = nombre
                n = n + 1
            End If
        End If
    Next v

    If n > 0 Then
        ReDim Preserve nombres(0 To n - 1)
        OrdenarRutas nombres
        For i = 0 To n - 1
            resultado.Add nombres(i)
        Next i
    End If

    Set ListaCarpetasUsuario = resultado
End Function


'------------------------------------------------------------------------------
' ListaSubcarpetas: todas las subcarpetas (de cualquier nivel) de una carpeta.
'------------------------------------------------------------------------------
Public Function ListaSubcarpetas(ByVal db As Object, ByVal rutaVisible As String) As Collection
    Dim resultado As New Collection
    Dim carpeta As Variant
    Dim prefijo As String

    If Len(rutaVisible) > 0 And Not EsCarpetaSistema(rutaVisible) Then
        prefijo = LCase$(rutaVisible) & "\"
        For Each carpeta In ListaCarpetasUsuario(db)
            If Left$(LCase$(CStr(carpeta)), Len(prefijo)) = prefijo Then
                resultado.Add CStr(carpeta)
            End If
        Next carpeta
    End If

    Set ListaSubcarpetas = resultado
End Function


'------------------------------------------------------------------------------
' CarpetaSeleccionada: ruta de la carpeta de la fila seleccionada en la hoja
' Carpetas. Devuelve "" si no estás en esa hoja o no hay fila válida.
'------------------------------------------------------------------------------
Public Function CarpetaSeleccionada() As String
    Dim fila As Long

    If ActiveSheet Is Nothing Then Exit Function
    If ActiveSheet.Name <> HOJA_CARPETAS Then Exit Function
    If TypeName(Selection) <> "Range" Then Exit Function

    fila = ActiveCell.Row
    If fila < 2 Then Exit Function

    CarpetaSeleccionada = Trim$(CStr(ActiveSheet.Cells(fila, COL_C_RUTA).Value))
End Function


'------------------------------------------------------------------------------
' NombreInterno / NombreVisible: traducen entre el nombre que ves en Excel y
' el nombre que usa Notes por dentro.
'   "Bandeja de entrada" <-> "($Inbox)"      "Enviados" <-> "($Sent)"
'------------------------------------------------------------------------------
Public Function NombreInterno(ByVal visible As String) As String
    Select Case LCase$(visible)
        Case LCase$(CARPETA_BANDEJA): NombreInterno = "($Inbox)"
        Case LCase$(CARPETA_ENVIADOS): NombreInterno = "($Sent)"
        Case Else: NombreInterno = visible
    End Select
End Function

Public Function NombreVisible(ByVal interno As String) As String
    Select Case LCase$(interno)
        Case "($inbox)": NombreVisible = CARPETA_BANDEJA
        Case "($sent)": NombreVisible = CARPETA_ENVIADOS
        Case Else: NombreVisible = interno
    End Select
End Function


'------------------------------------------------------------------------------
' EsCarpetaSistema: True para "Bandeja de entrada" y "Enviados".
'------------------------------------------------------------------------------
Public Function EsCarpetaSistema(ByVal visible As String) As Boolean
    EsCarpetaSistema = (LCase$(visible) = LCase$(CARPETA_BANDEJA) Or _
                        LCase$(visible) = LCase$(CARPETA_ENVIADOS))
End Function


'------------------------------------------------------------------------------
' ExisteCarpeta: True si existe la carpeta (nombre visible).
'------------------------------------------------------------------------------
Public Function ExisteCarpeta(ByVal db As Object, ByVal rutaVisible As String) As Boolean
    Dim vista As Object

    On Error Resume Next
    Set vista = db.GetView(NombreInterno(rutaVisible))
    On Error GoTo 0

    ExisteCarpeta = Not vista Is Nothing
End Function


'------------------------------------------------------------------------------
' HojaExiste: True si el libro tiene una hoja con ese nombre (sin crearla).
'------------------------------------------------------------------------------
Public Function HojaExiste(ByVal nombre As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0

    HojaExiste = Not ws Is Nothing
End Function


'==============================================================================
' FUNCIONES INTERNAS
'==============================================================================

'------------------------------------------------------------------------------
' EscribirFilaCarpeta: escribe una carpeta en la hoja y avanza la fila.
'------------------------------------------------------------------------------
Private Sub EscribirFilaCarpeta(ByVal ws As Worksheet, ByRef fila As Long, ByVal db As Object, _
                                ByVal rutaVisible As String, ByVal tipo As String)
    Dim nivel As Long
    Dim total As Variant
    Dim noLeidos As Variant

    nivel = UBound(Split(rutaVisible, "\"))     ' 0 = primer nivel, 1 = subcarpeta...

    With ws.Cells(fila, COL_C_NOMBRE)
        .Value = TextoSeguro(UltimoNivel(rutaVisible))
        .IndentLevel = Application.WorksheetFunction.Min(nivel * 2, 15)   ' Sangría = árbol
        .Font.Bold = (nivel = 0)
    End With
    ws.Cells(fila, COL_C_RUTA).Value = TextoSeguro(rutaVisible)
    ws.Cells(fila, COL_C_TIPO).Value = tipo

    If tipo <> "Agrupación" Then
        ContarCorreos db, NombreInterno(rutaVisible), total, noLeidos
        ws.Cells(fila, COL_C_TOTAL).Value = total
        ws.Cells(fila, COL_C_NOLEIDOS).Value = noLeidos
    End If

    ' Gris para las del sistema y las agrupaciones
    If tipo <> "Carpeta" Then
        ws.Range(ws.Cells(fila, COL_C_NOMBRE), ws.Cells(fila, COL_C_TIPO)).Font.Color = RGB(110, 110, 110)
    End If

    fila = fila + 1
End Sub


'------------------------------------------------------------------------------
' ContarCorreos: nº de correos y de no leídos de una carpeta.
' Si Notes no puede dar el dato (versiones antiguas), lo deja vacío.
'------------------------------------------------------------------------------
Private Sub ContarCorreos(ByVal db As Object, ByVal interno As String, _
                          ByRef total As Variant, ByRef noLeidos As Variant)
    Dim vista As Object
    Dim noLeidas As Object

    total = ""
    noLeidos = ""

    On Error Resume Next
    Set vista = db.GetView(interno)
    If vista Is Nothing Then Exit Sub

    total = vista.EntryCount
    Set noLeidas = vista.GetAllUnreadEntries()
    If Not noLeidas Is Nothing Then noLeidos = noLeidas.Count
    On Error GoTo 0
End Sub


'------------------------------------------------------------------------------
' NumeroCorreos: nº de correos de una carpeta (0 si no se puede saber).
'------------------------------------------------------------------------------
Private Function NumeroCorreos(ByVal db As Object, ByVal rutaVisible As String) As Long
    Dim total As Variant
    Dim noLeidos As Variant

    ContarCorreos db, NombreInterno(rutaVisible), total, noLeidos
    If IsNumeric(total) Then NumeroCorreos = CLng(total)
End Function


'------------------------------------------------------------------------------
' PasarCorreosABandeja: pone en la Bandeja de entrada todos los correos de una
' carpeta. Devuelve cuántos.
'------------------------------------------------------------------------------
Private Function PasarCorreosABandeja(ByVal db As Object, ByVal rutaVisible As String) As Long
    Dim vista As Object
    Dim doc As Object
    Dim docs As New Collection
    Dim d As Variant

    Set vista = db.GetView(NombreInterno(rutaVisible))
    If vista Is Nothing Then Exit Function

    ' Primero apuntamos los correos y DESPUÉS los movemos: si los movemos
    ' mientras recorremos la carpeta, Notes puede "perder el sitio".
    vista.AutoUpdate = False
    Set doc = vista.GetFirstDocument()
    Do While Not doc Is Nothing
        docs.Add doc
        Set doc = vista.GetNextDocument(doc)
    Loop

    For Each d In docs
        d.PutInFolder "($Inbox)"
    Next d

    PasarCorreosABandeja = docs.Count
End Function


'------------------------------------------------------------------------------
' ComprobarSeleccion: avisa si no hay carpeta seleccionada o si es del sistema.
'------------------------------------------------------------------------------
Private Function ComprobarSeleccion(ByVal ruta As String, ByVal titulo As String) As Boolean
    ComprobarSeleccion = False

    If Len(ruta) = 0 Then
        MsgBox "Primero ve a la hoja '" & HOJA_CARPETAS & "' y haz clic en la fila de la carpeta." & _
               vbCrLf & vbCrLf & "(Si la hoja está vacía, pulsa 'Actualizar carpetas'.)", _
               vbInformation, titulo
    ElseIf EsCarpetaSistema(ruta) Then
        MsgBox "'" & ruta & "' es una carpeta del sistema y no se puede modificar.", _
               vbExclamation, titulo
    Else
        ComprobarSeleccion = True
    End If
End Function


'------------------------------------------------------------------------------
' SeleccionarCarpetaEnHoja: activa la hoja Carpetas y selecciona esa ruta.
'------------------------------------------------------------------------------
Private Sub SeleccionarCarpetaEnHoja(ByVal rutaVisible As String)
    Dim ws As Worksheet
    Dim celda As Range

    Set ws = ObtenerHoja(HOJA_CARPETAS)
    ws.Activate
    Set celda = ws.Columns(COL_C_RUTA).Find(What:=rutaVisible, LookIn:=xlValues, LookAt:=xlWhole, _
                                            MatchCase:=False)
    If Not celda Is Nothing Then ws.Cells(celda.Row, COL_C_NOMBRE).Select
End Sub


'------------------------------------------------------------------------------
' LimpiarRuta: quita espacios sobrantes y niveles vacíos.
'   "  Clientes \\ Acme\ "  ->  "Clientes\Acme"
'------------------------------------------------------------------------------
Public Function LimpiarRuta(ByVal ruta As String) As String
    Dim partes As Variant
    Dim resultado As String
    Dim i As Long

    partes = Split(Replace(ruta, "/", "\"), "\")
    For i = LBound(partes) To UBound(partes)
        If Len(Trim$(partes(i))) > 0 Then
            If Len(resultado) > 0 Then resultado = resultado & "\"
            resultado = resultado & Trim$(partes(i))
        End If
    Next i

    LimpiarRuta = resultado
End Function


'------------------------------------------------------------------------------
' RutaValida: comprueba que una ruta de carpeta se puede usar en Notes.
'------------------------------------------------------------------------------
Private Function RutaValida(ByVal ruta As String, ByRef motivo As String) As Boolean
    RutaValida = False

    If Len(ruta) = 0 Then
        motivo = "El nombre está vacío."
    ElseIf Left$(ruta, 1) = "(" Or Left$(ruta, 1) = "$" Then
        motivo = "El nombre no puede empezar por ( ni por $ (son nombres reservados de Notes)."
    ElseIf InStr(ruta, "|") > 0 Then
        motivo = "El nombre no puede contener el carácter |"
    ElseIf EsCarpetaSistema(Split(ruta, "\")(0)) Then
        motivo = "No se pueden crear carpetas dentro de '" & Split(ruta, "\")(0) & "'."
    Else
        RutaValida = True
    End If
End Function


'------------------------------------------------------------------------------
' UltimoNivel: "Clientes\Acme\2026" -> "2026"
' RutaPadre:   "Clientes\Acme\2026" -> "Clientes\Acme"
'------------------------------------------------------------------------------
Public Function UltimoNivel(ByVal ruta As String) As String
    Dim posicion As Long
    posicion = InStrRev(ruta, "\")
    UltimoNivel = Mid$(ruta, posicion + 1)
End Function

Public Function RutaPadre(ByVal ruta As String) As String
    Dim posicion As Long
    posicion = InStrRev(ruta, "\")
    If posicion > 0 Then RutaPadre = Left$(ruta, posicion - 1)
End Function


'------------------------------------------------------------------------------
' OrdenarRutas: ordena alfabéticamente de forma que cada carpeta quede justo
' antes de sus subcarpetas. (Ordenación por inserción: sencilla y suficiente
' para unos cientos de carpetas.)
'------------------------------------------------------------------------------
Private Sub OrdenarRutas(ByRef nombres() As String)
    Dim i As Long
    Dim j As Long
    Dim actual As String

    For i = LBound(nombres) + 1 To UBound(nombres)
        actual = nombres(i)
        j = i - 1
        Do While j >= LBound(nombres)
            If ClaveOrden(nombres(j)) <= ClaveOrden(actual) Then Exit Do
            nombres(j + 1) = nombres(j)
            j = j - 1
        Loop
        nombres(j + 1) = actual
    Next i
End Sub

' Cambiamos "\" por un carácter muy "pequeño" para que "A\B" vaya antes que
' "A B" y el árbol no se desordene.
Private Function ClaveOrden(ByVal ruta As String) As String
    ClaveOrden = LCase$(Replace(ruta, "\", Chr$(1)))
End Function

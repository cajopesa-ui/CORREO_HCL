Attribute VB_Name = "modAdjuntos"
'==============================================================================
' MÓDULO: modAdjuntos
'------------------------------------------------------------------------------
' Guarda en una carpeta de tu PC los archivos adjuntos de los correos que
' aparecen en la hoja "Bandeja" (primero ejecuta LeerBandejaEntrada).
'
' Puedes elegir:
'   - Solo los correos de las filas que tengas SELECCIONADAS en la hoja, o
'   - Todos los correos de la lista.
'
' Si ya existe un archivo con el mismo nombre en la carpeta, NO se
' sobrescribe: se guarda como "nombre (2).ext", "nombre (3).ext", etc.
'
' Macro principal: GuardarAdjuntos  (asígnala a un botón)
'
' Además, GuardarAdjuntosCarpeta guarda los adjuntos de una carpeta ENTERA de
' Notes (y de sus subcarpetas), creando en tu PC la misma estructura de
' carpetas. Se usa desde la hoja "Carpetas".
'==============================================================================
Option Explicit


'------------------------------------------------------------------------------
' GuardarAdjuntos  -> MACRO PRINCIPAL
'------------------------------------------------------------------------------
Public Sub GuardarAdjuntos()
    Dim ws As Worksheet
    Dim filas As Collection     ' Números de fila a procesar
    Dim carpeta As String
    Dim respuesta As VbMsgBoxResult
    Dim sesion As Object
    Dim db As Object
    Dim doc As Object
    Dim adjunto As Object
    Dim nombres As Variant
    Dim fila As Variant
    Dim unid As String
    Dim i As Long
    Dim totalArchivos As Long
    Dim correosNoEncontrados As Long

    On Error GoTo ManejoError

    Set ws = PrepararHojaBandeja()

    ' --- 1) ¿Qué correos? ------------------------------------------------------
    respuesta = MsgBox("¿Guardar los adjuntos SOLO de las filas seleccionadas?" & vbCrLf & vbCrLf & _
                       "  Sí  = solo las filas seleccionadas en la hoja '" & HOJA_BANDEJA & "'" & vbCrLf & _
                       "  No  = todos los correos de la lista", _
                       vbQuestion + vbYesNoCancel, "Guardar adjuntos")
    If respuesta = vbCancel Then Exit Sub

    If respuesta = vbYes Then
        Set filas = FilasSeleccionadas(ws)
    Else
        Set filas = TodasLasFilas(ws)
    End If

    If filas.Count = 0 Then
        MsgBox "No hay correos para procesar." & vbCrLf & _
               "Primero pulsa 'Leer bandeja de entrada' y, si elegiste 'Sí', " & _
               "selecciona alguna fila con correos.", vbInformation, "Guardar adjuntos"
        Exit Sub
    End If

    ' --- 2) ¿En qué carpeta? ---------------------------------------------------
    carpeta = ElegirCarpeta()
    If Len(carpeta) = 0 Then Exit Sub     ' El usuario canceló

    ' --- 3) Conectar con Notes -------------------------------------------------
    Application.ScreenUpdating = False
    Application.StatusBar = "Conectando con HCL Notes..."
    Set sesion = ObtenerSesion()
    Set db = AbrirBaseCorreo()

    ' --- 4) Recorrer los correos y extraer sus adjuntos ------------------------
    For Each fila In filas
        unid = CStr(ws.Cells(fila, COL_B_ID).Value)

        ' Solo si la fila tiene ID y la columna de adjuntos indica alguno
        If Len(unid) > 0 And Val(ws.Cells(fila, COL_B_ADJUNTOS).Value) > 0 Then
            Application.StatusBar = "Guardando adjuntos del correo de la fila " & fila & "..."

            ' Buscamos el correo por su identificador único
            Set doc = Nothing
            On Error Resume Next
            Set doc = db.GetDocumentByUNID(unid)
            On Error GoTo ManejoError

            If doc Is Nothing Then
                ' Puede que se haya borrado o movido desde que se leyó la bandeja
                correosNoEncontrados = correosNoEncontrados + 1
            Else
                nombres = NombresAdjuntos(sesion, doc)
                For i = LBound(nombres) To UBound(nombres)
                    If Len(CStr(nombres(i))) > 0 Then
                        Set adjunto = doc.GetAttachment(CStr(nombres(i)))
                        If Not adjunto Is Nothing Then
                            adjunto.ExtractFile RutaLibre(carpeta, CStr(nombres(i)))
                            totalArchivos = totalArchivos + 1
                        End If
                    End If
                Next i
            End If
        End If
    Next fila

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Se han guardado " & totalArchivos & " archivo(s) en:" & vbCrLf & carpeta & _
           IIf(correosNoEncontrados > 0, vbCrLf & vbCrLf & correosNoEncontrados & _
               " correo(s) ya no se encontraron en el buzón (¿borrados o movidos?).", ""), _
           vbInformation, "Guardar adjuntos"

Salir:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Guardar adjuntos", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' GuardarAdjuntosCarpeta  -> MACRO (botón de la hoja Carpetas)
' Guarda los adjuntos de TODOS los correos de la carpeta seleccionada en la
' hoja Carpetas (y, si quieres, de sus subcarpetas). En tu PC se crea la misma
' estructura:  <carpeta elegida>\Clientes\Acme\2026\factura.pdf
'------------------------------------------------------------------------------
Public Sub GuardarAdjuntosCarpeta()
    Dim sesion As Object
    Dim db As Object
    Dim vista As Object
    Dim doc As Object
    Dim adjunto As Object
    Dim fso As Object
    Dim carpetas As New Collection
    Dim subcarpetas As Collection
    Dim ruta As String
    Dim carpetaBase As String
    Dim destinoPC As String
    Dim respuesta As VbMsgBoxResult
    Dim nombres As Variant
    Dim carpeta As Variant
    Dim s As Variant
    Dim i As Long
    Dim totalArchivos As Long
    Dim totalCorreos As Long

    On Error GoTo ManejoError

    ruta = CarpetaSeleccionada()
    If Len(ruta) = 0 Then
        MsgBox "Ve a la hoja '" & HOJA_CARPETAS & "', haz clic en la fila de la carpeta " & _
               "y vuelve a pulsar el botón.", vbInformation, "Guardar adjuntos de la carpeta"
        Exit Sub
    End If

    Set sesion = ObtenerSesion()
    Set db = AbrirBaseCorreo()
    carpetas.Add ruta

    Set subcarpetas = ListaSubcarpetas(db, ruta)
    If subcarpetas.Count > 0 Then
        respuesta = MsgBox("La carpeta '" & ruta & "' tiene " & subcarpetas.Count & " subcarpeta(s)." & _
                           vbCrLf & vbCrLf & "¿Guardar también sus adjuntos?", _
                           vbQuestion + vbYesNoCancel, "Guardar adjuntos de la carpeta")
        If respuesta = vbCancel Then Exit Sub
        If respuesta = vbYes Then
            For Each s In subcarpetas
                carpetas.Add CStr(s)
            Next s
        End If
    End If

    carpetaBase = ElegirCarpeta()
    If Len(carpetaBase) = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each carpeta In carpetas
        Set vista = Nothing
        On Error Resume Next
        Set vista = db.GetView(NombreInterno(CStr(carpeta)))
        On Error GoTo ManejoError

        If Not vista Is Nothing Then
            ' En el PC: <base>\Clientes\Acme\2026
            destinoPC = fso.BuildPath(carpetaBase, RutaParaWindows(CStr(carpeta)))
            vista.AutoUpdate = False

            Set doc = vista.GetFirstDocument()
            Do While Not doc Is Nothing
                totalCorreos = totalCorreos + 1
                If totalCorreos Mod 20 = 0 Then
                    Application.StatusBar = "Carpeta '" & carpeta & "': " & totalCorreos & _
                                            " correos revisados, " & totalArchivos & " archivos guardados..."
                End If

                nombres = NombresAdjuntos(sesion, doc)
                For i = LBound(nombres) To UBound(nombres)
                    If Len(CStr(nombres(i))) > 0 Then
                        Set adjunto = doc.GetAttachment(CStr(nombres(i)))
                        If Not adjunto Is Nothing Then
                            ' La carpeta del PC solo se crea si hay algo que guardar
                            CrearCarpetaPC destinoPC
                            adjunto.ExtractFile RutaLibre(destinoPC, CStr(nombres(i)))
                            totalArchivos = totalArchivos + 1
                        End If
                    End If
                Next i

                Set doc = vista.GetNextDocument(doc)
            Loop
        End If
    Next carpeta

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Se han revisado " & totalCorreos & " correo(s) y guardado " & totalArchivos & _
           " archivo(s) en:" & vbCrLf & carpetaBase, vbInformation, "Guardar adjuntos de la carpeta"

Salir:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ManejoError:
    MostrarError "Guardar adjuntos de la carpeta", Err.Number, Err.Description
    Resume Salir
End Sub


'------------------------------------------------------------------------------
' RutaParaWindows: convierte la ruta de una carpeta de Notes en una ruta
' válida para Windows, cambiando por "_" los caracteres no permitidos.
'   "Clientes\Acme: 2026"  ->  "Clientes\Acme_ 2026"
'------------------------------------------------------------------------------
Private Function RutaParaWindows(ByVal rutaNotes As String) As String
    Dim partes As Variant
    Dim parte As String
    Dim caracter As Variant
    Dim resultado As String
    Dim i As Long

    partes = Split(rutaNotes, "\")
    For i = LBound(partes) To UBound(partes)
        parte = partes(i)
        For Each caracter In Array("/", ":", "*", "?", Chr$(34), "<", ">", "|")
            parte = Replace(parte, caracter, "_")
        Next caracter
        ' Windows no admite nombres que acaben en punto o espacio
        Do While Len(parte) > 0 And (Right$(parte, 1) = "." Or Right$(parte, 1) = " ")
            parte = Left$(parte, Len(parte) - 1)
        Loop
        If Len(parte) = 0 Then parte = "_"

        If Len(resultado) > 0 Then resultado = resultado & "\"
        resultado = resultado & parte
    Next i

    RutaParaWindows = resultado
End Function


'------------------------------------------------------------------------------
' CrearCarpetaPC: crea una carpeta en el PC, y sus "padres" si no existen.
'------------------------------------------------------------------------------
Private Sub CrearCarpetaPC(ByVal ruta As String)
    Dim fso As Object

    If Len(ruta) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(ruta) Then Exit Sub

    CrearCarpetaPC fso.GetParentFolderName(ruta)   ' Primero el padre
    fso.CreateFolder ruta
End Sub


'------------------------------------------------------------------------------
' ElegirCarpeta: muestra la ventana de Windows para elegir una carpeta.
' Devuelve la ruta elegida, o "" si se cancela.
'------------------------------------------------------------------------------
Public Function ElegirCarpeta() As String
    ' msoFileDialogFolderPicker = 4 (selector de carpetas)
    With Application.FileDialog(4)
        .Title = "Elige la carpeta donde guardar los adjuntos"
        .AllowMultiSelect = False
        If .Show = -1 Then             ' -1 = el usuario pulsó Aceptar
            ElegirCarpeta = .SelectedItems(1)
        Else
            ElegirCarpeta = ""
        End If
    End With
End Function


'------------------------------------------------------------------------------
' RutaLibre: devuelve una ruta en "carpeta" para "nombreArchivo" que NO exista
' todavía, añadiendo " (2)", " (3)"... si hace falta.
'------------------------------------------------------------------------------
Public Function RutaLibre(ByVal carpeta As String, ByVal nombreArchivo As String) As String
    Dim fso As Object
    Dim baseNombre As String
    Dim extension As String
    Dim ruta As String
    Dim n As Long

    Set fso = CreateObject("Scripting.FileSystemObject")

    baseNombre = fso.GetBaseName(nombreArchivo)    ' "informe"
    extension = fso.GetExtensionName(nombreArchivo) ' "pdf"
    If Len(extension) > 0 Then extension = "." & extension

    ruta = fso.BuildPath(carpeta, nombreArchivo)
    n = 1
    Do While fso.FileExists(ruta)
        n = n + 1
        ruta = fso.BuildPath(carpeta, baseNombre & " (" & n & ")" & extension)
    Loop

    RutaLibre = ruta
End Function


'------------------------------------------------------------------------------
' FilasSeleccionadas: números de fila (a partir de la 2) seleccionados en la
' hoja Bandeja. Si la hoja activa no es la Bandeja devuelve una lista vacía.
'------------------------------------------------------------------------------
Public Function FilasSeleccionadas(ByVal ws As Worksheet) As Collection
    Dim resultado As New Collection
    Dim area As Range
    Dim filaRango As Range

    If ActiveSheet.Name = ws.Name And TypeName(Selection) = "Range" Then
        ' Una selección puede tener varias zonas (con Ctrl+clic): las recorremos
        For Each area In Selection.Areas
            For Each filaRango In area.Rows
                If filaRango.Row >= 2 Then
                    ' La "clave" evita repetir la misma fila dos veces
                    On Error Resume Next
                    resultado.Add filaRango.Row, CStr(filaRango.Row)
                    On Error GoTo 0
                End If
            Next filaRango
        Next area
    End If

    Set FilasSeleccionadas = resultado
End Function


'------------------------------------------------------------------------------
' TodasLasFilas: todas las filas con correo de la hoja Bandeja.
'------------------------------------------------------------------------------
Private Function TodasLasFilas(ByVal ws As Worksheet) As Collection
    Dim resultado As New Collection
    Dim ultimaFila As Long
    Dim fila As Long

    ultimaFila = ws.Cells(ws.Rows.Count, COL_B_ID).End(xlUp).Row
    For fila = 2 To ultimaFila
        resultado.Add fila
    Next fila

    Set TodasLasFilas = resultado
End Function

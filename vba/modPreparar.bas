Attribute VB_Name = "modPreparar"
'==============================================================================
' MÓDULO: modPreparar
'------------------------------------------------------------------------------
' Prepara el libro de Excel AUTOMÁTICAMENTE:
'   - Crea las hojas "Bandeja" y "Enviar" (si no existen)
'   - Escribe los encabezados de las columnas
'   - Añade los botones que ejecutan las macros
'
' Macro principal: PrepararLibro
'   Ejecútala UNA VEZ después de importar los módulos:
'   Alt+F8 > PrepararLibro > Ejecutar
'   (Si la vuelves a ejecutar no pasa nada: no borra tus datos.)
'==============================================================================
Option Explicit


'------------------------------------------------------------------------------
' PrepararLibro  -> MACRO PRINCIPAL
'------------------------------------------------------------------------------
Public Sub PrepararLibro()
    Dim wsBandeja As Worksheet
    Dim wsEnviar As Worksheet

    On Error GoTo ManejoError

    Set wsBandeja = PrepararHojaBandeja()
    Set wsEnviar = PrepararHojaEnviar()

    ' Borramos el antiguo botón "Desconectar de Notes" (versiones anteriores),
    ' que ahora se sustituye por el botón Conectar/Desconectar.
    BorrarBoton wsBandeja, "btn_DesconectarNotes"
    BorrarBoton wsEnviar, "btn_DesconectarNotes"

    ' --- Botones de la hoja Bandeja (a la derecha de los datos, columna H) ---
    CrearBoton wsBandeja, 1, "Leer bandeja de entrada", "LeerBandejaEntrada"
    CrearBoton wsBandeja, 2, "Guardar adjuntos", "GuardarAdjuntos"
    CrearBoton wsBandeja, 3, "Enviar correos", "EnviarCorreos"
    CrearBoton wsBandeja, 4, TEXTO_CONECTAR, "ConectarDesconectarNotes"

    ' --- Botones de la hoja Enviar --------------------------------------------
    CrearBoton wsEnviar, 1, "Enviar correos", "EnviarCorreos"
    CrearBoton wsEnviar, 2, TEXTO_CONECTAR, "ConectarDesconectarNotes"

    ' Ponemos el texto correcto en el botón de conexión (por si ya estabas conectado)
    ActualizarBotonesConexion

    ' Fila de ejemplo en la hoja Enviar, solo si está vacía
    If Len(CStr(wsEnviar.Cells(2, COL_E_PARA).Value)) = 0 Then
        wsEnviar.Cells(2, COL_E_PARA).Value = "destinatario@empresa.com"
        wsEnviar.Cells(2, COL_E_CC).Value = ""
        wsEnviar.Cells(2, COL_E_ASUNTO).Value = "Prueba desde Excel"
        wsEnviar.Cells(2, COL_E_CUERPO).Value = "Hola," & vbLf & vbLf & _
            "Este es un correo de prueba enviado desde Excel con HCL Notes." & vbLf & vbLf & "Saludos"
        wsEnviar.Cells(2, COL_E_ADJUNTO).Value = ""
    End If

    wsBandeja.Activate
    MsgBox "¡Libro preparado!" & vbCrLf & vbCrLf & _
           "Hoja '" & HOJA_BANDEJA & "': pulsa 'Leer bandeja de entrada'." & vbCrLf & _
           "Hoja '" & HOJA_ENVIAR & "': rellena los correos y pulsa 'Enviar correos'." & vbCrLf & vbCrLf & _
           "Recuerda guardar el libro como 'Libro de Excel habilitado para macros (*.xlsm)'.", _
           vbInformation, "Preparar libro"
    Exit Sub

ManejoError:
    MostrarError "Preparar libro", Err.Number, Err.Description
End Sub


'------------------------------------------------------------------------------
' PrepararHojaBandeja: devuelve la hoja Bandeja, creándola y poniendo los
' encabezados si hace falta. La usan también otras macros.
'------------------------------------------------------------------------------
Public Function PrepararHojaBandeja() As Worksheet
    Dim ws As Worksheet
    Set ws = ObtenerHoja(HOJA_BANDEJA)

    ws.Cells(1, COL_B_FECHA).Value = "Fecha"
    ws.Cells(1, COL_B_REMITENTE).Value = "Remitente"
    ws.Cells(1, COL_B_ASUNTO).Value = "Asunto"
    ws.Cells(1, COL_B_EXTRACTO).Value = "Extracto del cuerpo"
    ws.Cells(1, COL_B_ADJUNTOS).Value = "Nº adjuntos"
    ws.Cells(1, COL_B_ID).Value = "ID (no tocar)"

    FormatoEncabezado ws.Range(ws.Cells(1, COL_B_FECHA), ws.Cells(1, COL_B_ID))

    ' Anchos de columna razonables (solo la primera vez, si siguen por defecto)
    If ws.Columns(COL_B_ASUNTO).ColumnWidth < 20 Then
        ws.Columns(COL_B_FECHA).ColumnWidth = 17
        ws.Columns(COL_B_REMITENTE).ColumnWidth = 25
        ws.Columns(COL_B_ASUNTO).ColumnWidth = 40
        ws.Columns(COL_B_EXTRACTO).ColumnWidth = 60
        ws.Columns(COL_B_ADJUNTOS).ColumnWidth = 12
        ws.Columns(COL_B_ID).ColumnWidth = 10
        ws.Columns(COL_B_ID).Font.Color = RGB(150, 150, 150)   ' Gris: es interna
    End If

    Set PrepararHojaBandeja = ws
End Function


'------------------------------------------------------------------------------
' PrepararHojaEnviar: igual que la anterior, para la hoja Enviar.
'------------------------------------------------------------------------------
Public Function PrepararHojaEnviar() As Worksheet
    Dim ws As Worksheet
    Set ws = ObtenerHoja(HOJA_ENVIAR)

    ws.Cells(1, COL_E_PARA).Value = "Para"
    ws.Cells(1, COL_E_CC).Value = "CC"
    ws.Cells(1, COL_E_ASUNTO).Value = "Asunto"
    ws.Cells(1, COL_E_CUERPO).Value = "Cuerpo"
    ws.Cells(1, COL_E_ADJUNTO).Value = "Adjunto (ruta completa)"
    ws.Cells(1, COL_E_ESTADO).Value = "Estado"

    FormatoEncabezado ws.Range(ws.Cells(1, COL_E_PARA), ws.Cells(1, COL_E_ESTADO))

    If ws.Columns(COL_E_CUERPO).ColumnWidth < 20 Then
        ws.Columns(COL_E_PARA).ColumnWidth = 30
        ws.Columns(COL_E_CC).ColumnWidth = 25
        ws.Columns(COL_E_ASUNTO).ColumnWidth = 30
        ws.Columns(COL_E_CUERPO).ColumnWidth = 50
        ws.Columns(COL_E_CUERPO).WrapText = True    ' Para ver los saltos de línea
        ws.Columns(COL_E_ADJUNTO).ColumnWidth = 35
        ws.Columns(COL_E_ESTADO).ColumnWidth = 30
    End If

    Set PrepararHojaEnviar = ws
End Function


'------------------------------------------------------------------------------
' ObtenerHoja: devuelve la hoja con ese nombre. Si no existe, la crea al final.
'------------------------------------------------------------------------------
Public Function ObtenerHoja(ByVal nombre As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombre
    End If

    Set ObtenerHoja = ws
End Function


'------------------------------------------------------------------------------
' FormatoEncabezado: negrita, fondo azul y letra blanca. (Uso interno)
'------------------------------------------------------------------------------
Private Sub FormatoEncabezado(ByVal rango As Range)
    With rango
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
End Sub


'------------------------------------------------------------------------------
' CrearBoton: añade un botón en la columna H de la hoja.
'   posicion -> 1 = primer botón, 2 = segundo (debajo), etc.
'   macro    -> nombre de la macro que se ejecuta al pulsarlo
'------------------------------------------------------------------------------
Private Sub CrearBoton(ByVal ws As Worksheet, ByVal posicion As Long, _
                       ByVal texto As String, ByVal macro As String)
    Dim boton As Object
    Dim izquierda As Double
    Dim arriba As Double

    izquierda = ws.Columns(8).Left + 5           ' Columna H
    arriba = ws.Rows(1).Top + 5 + (posicion - 1) * 35

    ' Si el botón ya existía (porque se ejecutó PrepararLibro antes), lo
    ' borramos para no tenerlo duplicado. Si no existía, no pasa nada.
    BorrarBoton ws, "btn_" & macro

    Set boton = ws.Buttons.Add(izquierda, arriba, 170, 28)
    boton.Name = "btn_" & macro
    boton.Caption = texto
    boton.OnAction = macro
End Sub


'------------------------------------------------------------------------------
' BorrarBoton: borra el botón con ese nombre, si existe. (Uso interno)
'------------------------------------------------------------------------------
Private Sub BorrarBoton(ByVal ws As Worksheet, ByVal nombre As String)
    On Error Resume Next
    ws.Buttons(nombre).Delete
    On Error GoTo 0
End Sub

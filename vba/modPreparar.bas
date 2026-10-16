Attribute VB_Name = "modPreparar"
'==============================================================================
' MÓDULO: modPreparar
'------------------------------------------------------------------------------
' Prepara el libro de Excel AUTOMÁTICAMENTE:
'   - Crea las hojas "Bandeja", "Enviar", "Carpetas", "Buscar" y "Reglas"
'     (si no existen) y una hoja oculta "Listas" para los desplegables
'   - Escribe los encabezados de las columnas
'   - Añade los botones que ejecutan las macros
'
' Macro principal: PrepararLibro
'   Ejecútala UNA VEZ después de importar los módulos:
'   Alt+F8 > PrepararLibro > Ejecutar
'   (Si la vuelves a ejecutar no pasa nada: no borra tus datos.)
'==============================================================================
Option Explicit

' Columna donde se colocan los botones en cada hoja (1 = A, 2 = B, ...)
Private Const BOTONES_BANDEJA As Long = 12    ' L
Private Const BOTONES_ENVIAR As Long = 8      ' H
Private Const BOTONES_CARPETAS As Long = 10   ' J
Private Const BOTONES_BUSCAR As Long = 4      ' D
Private Const BOTONES_REGLAS As Long = 7      ' G


'------------------------------------------------------------------------------
' PrepararLibro  -> MACRO PRINCIPAL
'------------------------------------------------------------------------------
Public Sub PrepararLibro()
    Dim wsBandeja As Worksheet
    Dim wsEnviar As Worksheet
    Dim wsCarpetas As Worksheet
    Dim wsBuscar As Worksheet
    Dim wsReglas As Worksheet

    On Error GoTo ManejoError

    Set wsBandeja = PrepararHojaBandeja()
    Set wsEnviar = PrepararHojaEnviar()
    Set wsCarpetas = PrepararHojaCarpetas()
    Set wsBuscar = PrepararHojaBuscar()
    Set wsReglas = PrepararHojaReglas()

    ' Borramos el antiguo botón "Desconectar de Notes" (versiones anteriores),
    ' que ahora se sustituye por el botón Conectar/Desconectar.
    BorrarBoton wsBandeja, "btn_DesconectarNotes"
    BorrarBoton wsEnviar, "btn_DesconectarNotes"

    ' --- Botones de la hoja Bandeja --------------------------------------------
    CrearBoton wsBandeja, BOTONES_BANDEJA, 1, "Leer bandeja de entrada", "LeerBandejaEntrada"
    CrearBoton wsBandeja, BOTONES_BANDEJA, 2, "Mover a carpeta", "MoverCorreos"
    CrearBoton wsBandeja, BOTONES_BANDEJA, 3, "Copiar a carpeta", "CopiarCorreos"
    CrearBoton wsBandeja, BOTONES_BANDEJA, 4, "Marcar como leído", "MarcarLeidos"
    CrearBoton wsBandeja, BOTONES_BANDEJA, 5, "Marcar como no leído", "MarcarNoLeidos"
    CrearBoton wsBandeja, BOTONES_BANDEJA, 6, "Guardar adjuntos", "GuardarAdjuntos"
    CrearBoton wsBandeja, BOTONES_BANDEJA, 7, "Enviar correos", "EnviarCorreos"
    CrearBoton wsBandeja, BOTONES_BANDEJA, 8, TEXTO_CONECTAR, "ConectarDesconectarNotes"

    ' --- Botones de la hoja Enviar --------------------------------------------
    CrearBoton wsEnviar, BOTONES_ENVIAR, 1, "Enviar correos", "EnviarCorreos"
    CrearBoton wsEnviar, BOTONES_ENVIAR, 2, TEXTO_CONECTAR, "ConectarDesconectarNotes"

    ' --- Botones de la hoja Carpetas ------------------------------------------
    CrearBoton wsCarpetas, BOTONES_CARPETAS, 1, "Actualizar carpetas", "ActualizarCarpetas"
    CrearBoton wsCarpetas, BOTONES_CARPETAS, 2, "Leer carpeta", "LeerCarpeta"
    CrearBoton wsCarpetas, BOTONES_CARPETAS, 3, "Crear carpeta / subcarpeta", "CrearCarpeta"
    CrearBoton wsCarpetas, BOTONES_CARPETAS, 4, "Renombrar carpeta", "RenombrarCarpeta"
    CrearBoton wsCarpetas, BOTONES_CARPETAS, 5, "Eliminar carpeta", "EliminarCarpeta"
    CrearBoton wsCarpetas, BOTONES_CARPETAS, 6, "Crear carpetas de la lista (col. G)", "CrearCarpetasDeLista"
    CrearBoton wsCarpetas, BOTONES_CARPETAS, 7, "Guardar adjuntos de la carpeta", "GuardarAdjuntosCarpeta"
    CrearBoton wsCarpetas, BOTONES_CARPETAS, 8, TEXTO_CONECTAR, "ConectarDesconectarNotes"

    ' --- Botones de la hoja Buscar --------------------------------------------
    CrearBoton wsBuscar, BOTONES_BUSCAR, 1, "Buscar", "BuscarCorreos"
    CrearBoton wsBuscar, BOTONES_BUSCAR, 2, "Limpiar criterios", "LimpiarBusqueda"
    CrearBoton wsBuscar, BOTONES_BUSCAR, 3, TEXTO_CONECTAR, "ConectarDesconectarNotes"

    ' --- Botones de la hoja Reglas --------------------------------------------
    CrearBoton wsReglas, BOTONES_REGLAS, 1, "Aplicar reglas a la Bandeja", "AplicarReglas"
    CrearBoton wsReglas, BOTONES_REGLAS, 2, TEXTO_CONECTAR, "ConectarDesconectarNotes"

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
           "Hoja '" & HOJA_BANDEJA & "': lee, mueve, copia y marca correos." & vbCrLf & _
           "Hoja '" & HOJA_ENVIAR & "': rellena los correos y pulsa 'Enviar correos'." & vbCrLf & _
           "Hoja '" & HOJA_CARPETAS & "': pulsa 'Actualizar carpetas' para ver tu árbol de carpetas." & vbCrLf & _
           "Hoja '" & HOJA_BUSCAR & "': busca en todas tus carpetas." & vbCrLf & _
           "Hoja '" & HOJA_REGLAS & "': archiva la Bandeja de entrada automáticamente." & vbCrLf & vbCrLf & _
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
    ws.Cells(1, COL_B_CARPETA).Value = "Carpeta"
    ws.Cells(1, COL_B_LEIDO).Value = "Leído"
    ws.Cells(1, COL_B_MOVER).Value = "Mover a (elige)"
    ws.Cells(1, COL_B_RESULTADO).Value = "Resultado"

    FormatoEncabezado ws.Range(ws.Cells(1, COL_B_FECHA), ws.Cells(1, COL_B_RESULTADO))
    ' La columna "Mover a" en otro color, para que se vea que es para rellenar
    ws.Cells(1, COL_B_MOVER).Interior.Color = RGB(197, 90, 17)

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
    If ws.Columns(COL_B_CARPETA).ColumnWidth < 20 Then
        ws.Columns(COL_B_CARPETA).ColumnWidth = 25
        ws.Columns(COL_B_LEIDO).ColumnWidth = 7
        ws.Columns(COL_B_MOVER).ColumnWidth = 28
        ws.Columns(COL_B_RESULTADO).ColumnWidth = 35
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
' PrepararHojaCarpetas: hoja con el árbol de carpetas.
'------------------------------------------------------------------------------
Public Function PrepararHojaCarpetas() As Worksheet
    Dim ws As Worksheet
    Set ws = ObtenerHoja(HOJA_CARPETAS)

    ws.Cells(1, COL_C_NOMBRE).Value = "Carpeta"
    ws.Cells(1, COL_C_RUTA).Value = "Ruta completa"
    ws.Cells(1, COL_C_TOTAL).Value = "Nº correos"
    ws.Cells(1, COL_C_NOLEIDOS).Value = "No leídos"
    ws.Cells(1, COL_C_TIPO).Value = "Tipo"
    ws.Cells(1, COL_C_NUEVA).Value = "Crear estas carpetas (una por fila)"
    ws.Cells(1, COL_C_NUEVA_RES).Value = "Resultado"

    FormatoEncabezado ws.Range(ws.Cells(1, COL_C_NOMBRE), ws.Cells(1, COL_C_TIPO))
    FormatoEncabezado ws.Range(ws.Cells(1, COL_C_NUEVA), ws.Cells(1, COL_C_NUEVA_RES))

    If ws.Columns(COL_C_RUTA).ColumnWidth < 20 Then
        ws.Columns(COL_C_NOMBRE).ColumnWidth = 35
        ws.Columns(COL_C_RUTA).ColumnWidth = 40
        ws.Columns(COL_C_TOTAL).ColumnWidth = 11
        ws.Columns(COL_C_NOLEIDOS).ColumnWidth = 10
        ws.Columns(COL_C_TIPO).ColumnWidth = 12
        ws.Columns(COL_C_NUEVA).ColumnWidth = 36
        ws.Columns(COL_C_NUEVA_RES).ColumnWidth = 25
        ' Ejemplo en la lista de carpetas a crear
        If Len(CStr(ws.Cells(2, COL_C_NUEVA).Value)) = 0 Then
            ws.Cells(2, COL_C_NUEVA).Value = "Ejemplo\Subcarpeta (borra esta fila)"
        End If
    End If

    Set PrepararHojaCarpetas = ws
End Function


'------------------------------------------------------------------------------
' PrepararHojaBuscar: hoja con los criterios de búsqueda (valores en col. B).
'------------------------------------------------------------------------------
Public Function PrepararHojaBuscar() As Worksheet
    Dim ws As Worksheet
    Set ws = ObtenerHoja(HOJA_BUSCAR)

    PrepararListas

    ws.Cells(1, 1).Value = "Buscar correos (rellena solo lo que quieras; los resultados salen en la hoja '" & _
                           HOJA_BANDEJA & "')"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12

    ws.Cells(FILA_BUS_TEXTO, 1).Value = "Texto (en asunto o cuerpo)"
    ws.Cells(FILA_BUS_REMITENTE, 1).Value = "Remitente contiene"
    ws.Cells(FILA_BUS_DESDE, 1).Value = "Desde (fecha)"
    ws.Cells(FILA_BUS_HASTA, 1).Value = "Hasta (fecha)"
    ws.Cells(FILA_BUS_CARPETA, 1).Value = "Carpeta (vacía = todas)"
    ws.Cells(FILA_BUS_SUBCARPETAS, 1).Value = "Incluir subcarpetas"
    ws.Range(ws.Cells(FILA_BUS_TEXTO, 1), ws.Cells(FILA_BUS_SUBCARPETAS, 1)).Font.Bold = True

    ' Celdas para rellenar: fondo amarillo claro
    ws.Range(ws.Cells(FILA_BUS_TEXTO, 2), ws.Cells(FILA_BUS_SUBCARPETAS, 2)).Interior.Color = RGB(255, 242, 204)
    ws.Range(ws.Cells(FILA_BUS_DESDE, 2), ws.Cells(FILA_BUS_HASTA, 2)).NumberFormat = "dd/mm/yyyy"

    ws.Cells(FILA_BUS_SUBCARPETAS + 2, 1).Value = _
        "Nota: buscar texto en el cuerpo de muchos correos puede tardar. " & _
        "Al pulsar 'Buscar' eliges cuántos resultados quieres y si los quieres completos (.txt)."
    ws.Cells(FILA_BUS_SUBCARPETAS + 2, 1).Font.Italic = True

    ' Desplegables: Sí/No y lista de carpetas (si ya existe)
    With ws.Cells(FILA_BUS_SUBCARPETAS, 2).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=ListaSiNo"
    End With
    If Len(CStr(ws.Cells(FILA_BUS_SUBCARPETAS, 2).Value)) = 0 Then ws.Cells(FILA_BUS_SUBCARPETAS, 2).Value = "Sí"
    AplicarDesplegableCarpetas ws.Cells(FILA_BUS_CARPETA, 2)

    If ws.Columns(2).ColumnWidth < 20 Then
        ws.Columns(1).ColumnWidth = 28
        ws.Columns(2).ColumnWidth = 40
    End If

    Set PrepararHojaBuscar = ws
End Function


'------------------------------------------------------------------------------
' PrepararHojaReglas: hoja con las reglas de archivado automático.
'------------------------------------------------------------------------------
Public Function PrepararHojaReglas() As Worksheet
    Dim ws As Worksheet
    Set ws = ObtenerHoja(HOJA_REGLAS)

    PrepararListas

    ws.Cells(1, COL_R_CAMPO).Value = "Buscar en"
    ws.Cells(1, COL_R_CONTIENE).Value = "Contiene (texto)"
    ws.Cells(1, COL_R_DESTINO).Value = "Mover a carpeta"
    ws.Cells(1, COL_R_ACTIVA).Value = "Activa"
    ws.Cells(1, COL_R_RESULTADO).Value = "Resultado (última vez)"
    FormatoEncabezado ws.Range(ws.Cells(1, COL_R_CAMPO), ws.Cells(1, COL_R_RESULTADO))

    ' Desplegables para las 199 primeras reglas
    With ws.Range(ws.Cells(2, COL_R_CAMPO), ws.Cells(200, COL_R_CAMPO)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=ListaCamposRegla"
    End With
    With ws.Range(ws.Cells(2, COL_R_ACTIVA), ws.Cells(200, COL_R_ACTIVA)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=ListaSiNo"
    End With
    AplicarDesplegableCarpetas ws.Range(ws.Cells(2, COL_R_DESTINO), ws.Cells(200, COL_R_DESTINO))

    If ws.Columns(COL_R_CONTIENE).ColumnWidth < 20 Then
        ws.Columns(COL_R_CAMPO).ColumnWidth = 14
        ws.Columns(COL_R_CONTIENE).ColumnWidth = 30
        ws.Columns(COL_R_DESTINO).ColumnWidth = 35
        ws.Columns(COL_R_ACTIVA).ColumnWidth = 8
        ws.Columns(COL_R_RESULTADO).ColumnWidth = 35
        ' Regla de ejemplo, desactivada
        If Len(CStr(ws.Cells(2, COL_R_CONTIENE).Value)) = 0 Then
            ws.Cells(2, COL_R_CAMPO).Value = "Remitente"
            ws.Cells(2, COL_R_CONTIENE).Value = "@ejemplo.com"
            ws.Cells(2, COL_R_ACTIVA).Value = "No"
            ws.Cells(2, COL_R_RESULTADO).Value = "(Ejemplo: elige la carpeta y pon Activa = Sí)"
        End If
    End If

    Set PrepararHojaReglas = ws
End Function


'------------------------------------------------------------------------------
' PrepararListas: hoja OCULTA "Listas" con las opciones de los desplegables
' (Sí/No y los campos de las reglas) y sus nombres de rango.
'------------------------------------------------------------------------------
Public Sub PrepararListas()
    Dim ws As Worksheet

    Set ws = ObtenerHoja(HOJA_LISTAS)

    ws.Range("A1").Value = "Sí"
    ws.Range("A2").Value = "No"
    ws.Range("B1").Value = "Remitente"
    ws.Range("B2").Value = "Para/CC"
    ws.Range("B3").Value = "Asunto"
    ws.Range("B4").Value = "Cuerpo"

    ThisWorkbook.Names.Add Name:="ListaSiNo", RefersTo:="='" & ws.Name & "'!$A$1:$A$2"
    ThisWorkbook.Names.Add Name:="ListaCamposRegla", RefersTo:="='" & ws.Name & "'!$B$1:$B$4"

    If ws.Visible = xlSheetVisible Then ws.Visible = xlSheetHidden
End Sub


'------------------------------------------------------------------------------
' ObtenerHoja: devuelve la hoja con ese nombre. Si no existe, la crea al final
' (sin cambiar la hoja en la que estás).
'------------------------------------------------------------------------------
Public Function ObtenerHoja(ByVal nombre As String) As Worksheet
    Dim ws As Worksheet
    Dim anterior As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0

    If ws Is Nothing Then
        Set anterior = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombre
        ' Worksheets.Add activa la hoja nueva: volvemos a la que estaba
        If Not anterior Is Nothing Then anterior.Activate
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
' CrearBoton: añade un botón en la hoja.
'   columna  -> columna donde va el botón (1 = A, 2 = B, ...)
'   posicion -> 1 = primer botón, 2 = segundo (debajo), etc.
'   macro    -> nombre de la macro que se ejecuta al pulsarlo
'------------------------------------------------------------------------------
Private Sub CrearBoton(ByVal ws As Worksheet, ByVal columna As Long, ByVal posicion As Long, _
                       ByVal texto As String, ByVal macro As String)
    Dim boton As Object
    Dim izquierda As Double
    Dim arriba As Double

    izquierda = ws.Columns(columna).Left + 5
    arriba = ws.Rows(1).Top + 5 + (posicion - 1) * 35

    ' Si el botón ya existía (porque se ejecutó PrepararLibro antes), lo
    ' borramos para no tenerlo duplicado. Si no existía, no pasa nada.
    BorrarBoton ws, "btn_" & macro

    Set boton = ws.Buttons.Add(izquierda, arriba, 200, 28)
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

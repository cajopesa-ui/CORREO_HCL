Attribute VB_Name = "modAyuda"
'==============================================================================
' MÓDULO: modAyuda
'------------------------------------------------------------------------------
' Ayuda interactiva integrada en el libro:
'
'   PrepararHojaAyuda -> crea la hoja "Ayuda" (la primera del libro) con un
'                        índice con enlaces y, para cada tema: para qué
'                        sirve, los pasos y un ejemplo.
'   AyudaVerPasos     -> (botón) salta al tema elegido en el desplegable
'   AyudaIrAHoja      -> (botón) te lleva a la hoja del tema y te recuerda
'                        los pasos en una ventana
'   AyudaVisita       -> (botón) visita guiada: recorre las hojas una a una
'   AyudaHoja         -> (botón "? Ayuda de esta hoja" en cada hoja)
'                        ayuda según la hoja en la que estés
'   PonerNotasAyuda   -> pone notas en los encabezados de las columnas
'                        (aparecen al pasar el ratón por encima)
'
' TODOS los textos de la ayuda están en CargarTemas. Si quieres cambiar o
' añadir algo, hazlo ahí y vuelve a ejecutar PrepararLibro.
'==============================================================================
Option Explicit

' Cada tema es un Array(titulo, hoja, celda, descripcion, pasos, ejemplo)
' donde "pasos" y "ejemplo" son colecciones de líneas de texto.
Private mTemas As Collection
Private mPasos As Collection        ' Pasos del tema que se está definiendo
Private mEjemplo As Collection      ' Ejemplo del tema que se está definiendo

Private Const T_TITULO As Long = 0
Private Const T_HOJA As Long = 1
Private Const T_CELDA As Long = 2
Private Const T_DESC As Long = 3
Private Const T_PASOS As Long = 4
Private Const T_EJEMPLO As Long = 5

Private Const CELDA_ELEGIR As String = "B4"        ' Desplegable "¿Qué quieres hacer?"
Private Const TEXTO_ELEGIR As String = "» Haz clic aquí y elige en la lista qué quieres hacer"


'==============================================================================
' TEXTOS DE LA AYUDA
'==============================================================================
Private Sub CargarTemas()
    If Not mTemas Is Nothing Then Exit Sub
    Set mTemas = New Collection

    '--------------------------------------------------------------------------
    Tema "Conectar y desconectar de Notes", HOJA_BANDEJA, "A1", _
         "El botón 'Conectar a Notes' abre tu sesión de HCL Notes y pasa a decir 'Desconectar de Notes'. " & _
         "No es obligatorio: los demás botones se conectan solos cuando lo necesitan."
    Paso "Abre el cliente HCL Notes (recomendado, aunque no imprescindible)."
    Paso "Pulsa 'Conectar a Notes' en cualquier hoja."
    Paso "Si te la pide, escribe tu contraseña de Notes. OJO: se ve mientras la escribes, que nadie mire."
    Paso "El botón pasa a decir 'Desconectar de Notes'. Púlsalo cuando quieras cerrar la sesión (por ejemplo, si cambias de usuario)."
    Ejemplo "Antes de conectar:   [ Conectar a Notes ]"
    Ejemplo "Ya conectado:        [ Desconectar de Notes ]"

    '--------------------------------------------------------------------------
    Tema "Leer la Bandeja de entrada", HOJA_BANDEJA, "A2", _
         "Copia tus correos más recientes a la hoja Bandeja: fecha, remitente, asunto, extracto, " & _
         "nº de adjuntos, carpeta y si está leído."
    Paso "Ve a la hoja Bandeja y pulsa 'Leer bandeja de entrada'."
    Paso "Escribe cuántos correos quieres leer (por ejemplo 50). Se empieza por los más recientes."
    Paso "¿Contenido completo? Responde No para ver solo el resumen en la hoja, o Sí para crear además un .txt con el texto entero."
    Paso "Espera a que termine (abajo, en la barra de estado, ves por dónde va). Los más recientes quedan arriba."
    Ejemplo "Fecha             Remitente      Asunto                   Nº adj.  Carpeta"
    Ejemplo "16/10/2026 09:47  Ana López      Presupuesto octubre      2        Bandeja de entrada"
    Ejemplo "15/10/2026 13:59  Xavier Garza   Distribución tabletas    1        Bandeja de entrada"

    '--------------------------------------------------------------------------
    Tema "Guardar el contenido completo en un .txt", HOJA_BANDEJA, "A2", _
         "Además de rellenar la hoja, guarda en un archivo de texto el cuerpo entero de cada correo, " & _
         "ordenado y fácil de leer para ti o para Claude (por ejemplo, para resumirlos o clasificarlos)."
    Paso "Pulsa 'Leer bandeja de entrada' (hoja Bandeja), 'Leer carpeta' (hoja Carpetas) o 'Buscar' (hoja Buscar)."
    Paso "Cuando pregunte '¿contenido completo?', responde Sí."
    Paso "En la ventana 'Guardar como' elige la carpeta de tu PC y el nombre del archivo."
    Paso "Ábrelo con el Bloc de notas, o pásaselo a Claude."
    Ejemplo "### CORREO 1 DE 25 ###"
    Ejemplo "Fecha: 2026-10-15 09:12"
    Ejemplo "De: Ana López <ana@acme.com>"
    Ejemplo "Asunto: Presupuesto octubre"
    Ejemplo "Adjuntos (1): presupuesto.pdf"
    Ejemplo "--- CUERPO ---"
    Ejemplo "Hola Juan, te adjunto el presupuesto..."
    Ejemplo "--- FIN CUERPO ---"
    Ejemplo "### FIN CORREO 1 ###"

    '--------------------------------------------------------------------------
    Tema "Enviar correos", HOJA_ENVIAR, "A2", _
         "Envía con Notes los correos que escribas en la hoja Enviar: una fila por correo."
    Paso "Ve a la hoja Enviar."
    Paso "Rellena al menos 'Para' y 'Asunto'. 'CC', 'Cuerpo' y 'Adjunto' son opcionales."
    Paso "Varios destinatarios o varios adjuntos: sepáralos con punto y coma ( ; )."
    Paso "Deja vacía la columna 'Estado' y pulsa 'Enviar correos'. Confirma."
    Paso "La columna 'Estado' dice 'Enviado' (verde) o 'ERROR' (rojo). Para reenviar una fila, borra su Estado."
    Ejemplo "Para                          Asunto     Cuerpo              Adjunto"
    Ejemplo "ana@empresa.com; luis@...     Informe    Hola, adjunto...    C:\Informes\semana.pdf"
    Ejemplo "Truco: Alt+Intro dentro de una celda = salto de línea en el cuerpo."

    '--------------------------------------------------------------------------
    Tema "Guardar los adjuntos", HOJA_BANDEJA, "A2", _
         "Guarda en tu PC los archivos adjuntos de los correos de la hoja Bandeja. " & _
         "Desde la hoja Carpetas puedes guardar los de una carpeta entera."
    Paso "Lee la bandeja (o una carpeta, o haz una búsqueda)."
    Paso "Opcional: selecciona las filas de los correos que te interesan."
    Paso "Pulsa 'Guardar adjuntos'. Sí = solo las filas seleccionadas; No = todos los de la lista."
    Paso "Elige la carpeta de tu PC. Nunca se sobrescribe nada: si ya existe, se guarda como 'nombre (2)'."
    Ejemplo "C:\Adjuntos\presupuesto.pdf"
    Ejemplo "C:\Adjuntos\presupuesto (2).pdf     <- si ya había uno con ese nombre"

    '--------------------------------------------------------------------------
    Tema "Ver, crear, renombrar y eliminar carpetas", HOJA_CARPETAS, "A2", _
         "La hoja Carpetas muestra el árbol de carpetas de tu buzón con el nº de correos. " & _
         "Desde ella puedes crearlas, renombrarlas y eliminarlas."
    Paso "Ve a la hoja Carpetas y pulsa 'Actualizar carpetas'."
    Paso "Haz clic en la fila de una carpeta para elegirla."
    Paso "'Crear carpeta / subcarpeta' crea la nueva DENTRO de la seleccionada (sin selección, en el primer nivel)."
    Paso "'Renombrar carpeta' cambia el nombre; sus subcarpetas la siguen."
    Paso "'Eliminar carpeta' pide dos confirmaciones. Los correos nunca se borran del buzón."
    Ejemplo "Clientes                 <- carpeta (ruta: Clientes)"
    Ejemplo "    Acme                 <- subcarpeta (ruta: Clientes\Acme)"
    Ejemplo "        2026             <- sub-subcarpeta (ruta: Clientes\Acme\2026)"

    '--------------------------------------------------------------------------
    Tema "Crear muchas carpetas de golpe", HOJA_CARPETAS, "G2", _
         "Escribe una ruta por fila en la columna G de la hoja Carpetas y créalas todas con un solo botón."
    Paso "Ve a la hoja Carpetas, columna G ('Crear estas carpetas')."
    Paso "Escribe una ruta por fila. Separa los niveles con \ (barra invertida)."
    Paso "Pulsa 'Crear carpetas de la lista (col. G)'."
    Paso "La columna H dice 'Creada', 'Ya existía' o el error."
    Ejemplo "Clientes"
    Ejemplo "Clientes\Acme"
    Ejemplo "Clientes\Acme\2026"
    Ejemplo "Facturas\2026"

    '--------------------------------------------------------------------------
    Tema "Leer una carpeta completa (con subcarpetas)", HOJA_CARPETAS, "A2", _
         "Copia a la hoja Bandeja los correos de una carpeta. Puedes incluir sus subcarpetas y quedarte solo con los últimos días."
    Paso "En la hoja Carpetas, haz clic en la fila de la carpeta."
    Paso "Pulsa 'Leer carpeta'. Si tiene subcarpetas, te pregunta si incluirlas."
    Paso "¿Solo los últimos días? Escribe un número (ej. 30) o déjalo vacío para todos."
    Paso "Di cuántos correos y si quieres el contenido completo (.txt)."
    Paso "Los correos aparecen en la hoja Bandeja, con su carpeta en la columna G."
    Ejemplo "Carpeta: Clientes    Subcarpetas: Sí    Días: 30"
    Ejemplo "=> correos de Clientes, Clientes\Acme y Clientes\Acme\2026 de los últimos 30 días"

    '--------------------------------------------------------------------------
    Tema "Mover o copiar correos a una carpeta", HOJA_BANDEJA, "I2", _
         "Elige la carpeta de destino en el desplegable de la columna 'Mover a' (naranja) y pulsa un botón."
    Paso "Lee la bandeja, una carpeta o haz una búsqueda."
    Paso "En la columna 'Mover a' (I) haz clic en la celda y elige la carpeta en el desplegable."
    Paso "'Mover a carpeta': el correo pasa a la nueva carpeta y SALE de la suya (columna G)."
    Paso "'Copiar a carpeta': el correo está en las dos carpetas a la vez."
    Paso "La columna 'Resultado' dice qué ha pasado. Si no ves una carpeta nueva: hoja Carpetas > 'Actualizar carpetas'."
    Ejemplo "Asunto             Carpeta              Mover a          Resultado"
    Ejemplo "Factura octubre    Bandeja de entrada   Facturas\2026    Movido a 'Facturas\2026'"

    '--------------------------------------------------------------------------
    Tema "Marcar como leído o no leído", HOJA_BANDEJA, "A2", _
         "Cambia el estado de lectura de los correos seleccionados en la hoja Bandeja (necesita Notes 8 o superior)."
    Paso "Selecciona las filas de los correos. Para varias sueltas usa Ctrl+clic; para un bloque, Mayús+clic."
    Paso "Pulsa 'Marcar como leído' o 'Marcar como no leído'."
    Ejemplo "Ctrl+clic en las filas 3, 5 y 8  ->  'Marcar como leído'"

    '--------------------------------------------------------------------------
    Tema "Buscar correos", HOJA_BUSCAR, "B3", _
         "Busca por texto, remitente, fechas y carpeta. Los resultados salen en la hoja Bandeja, " & _
         "desde donde puedes moverlos, marcarlos o guardar sus adjuntos."
    Paso "Ve a la hoja Buscar y rellena solo lo que quieras (celdas amarillas)."
    Paso "Carpeta vacía = busca en la Bandeja de entrada, Enviados y todas tus carpetas."
    Paso "Pulsa 'Buscar', di cuántos resultados y si quieres el contenido completo (.txt)."
    Paso "'Limpiar criterios' vacía las celdas para una búsqueda nueva."
    Ejemplo "Texto: presupuesto   Remitente: acme.com   Desde: 01/01/2026   Carpeta: (vacía)"
    Ejemplo "Todo el año en un .txt:  Desde = 01/01/2026, lo demás vacío, 5000 correos, completo = Sí"

    '--------------------------------------------------------------------------
    Tema "Reglas de archivado automático", HOJA_REGLAS, "A2", _
         "Mueve automáticamente los correos de la Bandeja de entrada a sus carpetas según el remitente, " & _
         "los destinatarios, el asunto o el cuerpo."
    Paso "Ve a la hoja Reglas. Una regla por fila: 'Buscar en', 'Contiene', 'Mover a carpeta' y 'Activa' = Sí."
    Paso "Pulsa 'Aplicar reglas a la Bandeja'."
    Paso "Te enseña un resumen de lo que va a mover. Confirma si estás de acuerdo."
    Paso "Las reglas se revisan de arriba abajo: si un correo cumple varias, gana la primera."
    Ejemplo "Buscar en   Contiene       Mover a carpeta    Activa"
    Ejemplo "Remitente   @acme.com      Clientes\Acme      Sí"
    Ejemplo "Asunto      factura        Facturas\2026      Sí"

    '--------------------------------------------------------------------------
    Tema "Clasificar por temas con Claude", HOJA_CLASIFICACION, "A2", _
         "Claude lee tus correos y propone a qué carpeta temática va cada uno. Excel crea las carpetas " & _
         "dentro de '" & CARPETA_CLASIFICACION & "' y copia allí los correos."
    Paso "Exporta los correos con contenido completo (.txt): Leer bandeja, Leer carpeta o Buscar."
    Paso "Pásale el .txt a Claude y pídele 'el archivo de clasificación'."
    Paso "En la hoja Clasificacion pulsa 'Importar clasificación de Claude' y elige ese archivo."
    Paso "Revisa: puedes cambiar la carpeta (columna D) o poner Aplicar = No (columna F)."
    Paso "Pulsa 'Aplicar clasificación'. Los correos se COPIAN: siguen también en la Bandeja."
    Ejemplo "Archivo de Claude (una línea por correo):  ID;Carpeta;Asunto"
    Ejemplo "8A6F5119...C05D;RRHH;Jubilación voluntaria"
    Ejemplo "3E6426EF...989C;Informatica\Incidencias y Serman;CASO 3039034"
    Ejemplo "=> " & CARPETA_CLASIFICACION & "\RRHH   y   " & CARPETA_CLASIFICACION & "\Informatica\Incidencias y Serman"

    '--------------------------------------------------------------------------
    Tema "Clasificar con palabras clave", HOJA_PALABRAS, "A2", _
         "Clasifica la Bandeja de entrada sin salir de tu PC, con las reglas de la hoja Palabras clave."
    Paso "Revisa la hoja Palabras clave: carpeta, dónde buscar y palabras separadas por ; (punto y coma)."
    Paso "En la hoja Clasificacion pulsa 'Clasificar con palabras clave' y di cuántos correos."
    Paso "Los que no cumplen ninguna regla salen como 'Sin clasificar' con Aplicar = No."
    Paso "Revisa la hoja y pulsa 'Aplicar clasificación'."
    Ejemplo "Carpeta                        Buscar en   Palabras"
    Ejemplo "Informatica\Despliegues CEAR   Asunto      CEAR;I26-;parches;BIOS"
    Ejemplo "RRHH\Sindicatos                Asunto      UCESHA;sindicat;huelga"
    Ejemplo "No distingue mayúsculas ni tildes: 'vacunacion' encuentra 'VACUNACIÓN'."

    '--------------------------------------------------------------------------
    Tema "Si algo falla", HOJA_AYUDA, "A1", _
         "Los problemas más habituales y cómo resolverlos."
    Paso "'Error 429' o 'No se pudo conectar': Excel y Notes deben ser los dos de 32 bits o los dos de 64 bits. Abre Notes al menos una vez."
    Paso "Pide la contraseña cada vez: en Notes activa la opción de compartir la contraseña con otros programas."
    Paso "No aparece una carpeta en un desplegable: hoja Carpetas > 'Actualizar carpetas'."
    Paso "Notes no muestra los cambios: pulsa F9 en Notes o cierra y abre la carpeta."
    Paso "Los botones no hacen nada: al abrir el libro pulsa 'Habilitar contenido' (macros)."
    Paso "Va muy lento: pide menos correos o usa el filtro de fechas."
End Sub


'------------------------------------------------------------------------------
' Tema / Paso / Ejemplo: ayudas para escribir los textos de arriba.
'------------------------------------------------------------------------------
Private Sub Tema(ByVal titulo As String, ByVal hoja As String, ByVal celda As String, ByVal descripcion As String)
    Set mPasos = New Collection
    Set mEjemplo = New Collection
    mTemas.Add Array(titulo, hoja, celda, descripcion, mPasos, mEjemplo)
End Sub

Private Sub Paso(ByVal texto As String)
    mPasos.Add texto
End Sub

Private Sub Ejemplo(ByVal texto As String)
    mEjemplo.Add texto
End Sub


'==============================================================================
' HOJA "AYUDA"
'==============================================================================

'------------------------------------------------------------------------------
' PrepararHojaAyuda: (re)escribe la hoja Ayuda completa y la pone la primera.
'------------------------------------------------------------------------------
Public Function PrepararHojaAyuda() As Worksheet
    Dim ws As Worksheet
    Dim wsListas As Worksheet
    Dim datos As Variant
    Dim linea As Variant
    Dim fila As Long
    Dim filaIndice As Long
    Dim n As Long
    Dim i As Long

    CargarTemas
    PrepararListas

    Set ws = ObtenerHoja(HOJA_AYUDA)
    If ws.Index <> 1 Then ws.Move Before:=ThisWorkbook.Worksheets(1)   ' Que sea la primera hoja

    ' Empezamos de cero (la hoja solo tiene textos de ayuda)
    ws.Hyperlinks.Delete
    ws.Cells.Clear
    ws.Columns(1).ColumnWidth = 3
    ws.Columns(2).ColumnWidth = 95
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11
    ws.Cells.VerticalAlignment = xlTop

    ' --- Cabecera --------------------------------------------------------------------
    With ws.Range("B1")
        .Value = "AYUDA  ·  Correo HCL Notes desde Excel"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With
    ws.Range("B2").Value = "Haz clic en un tema del ÍNDICE, o elige qué quieres hacer en la celda amarilla y usa los botones de la derecha. " & _
                           "En cada hoja tienes el botón '? Ayuda de esta hoja' y notas en los encabezados (pasa el ratón por encima)."
    ws.Range("B2").WrapText = True
    ws.Range("B2").Font.Italic = True
    ws.Range("B2").Font.Color = RGB(89, 89, 89)

    ws.Range("B3").Value = "¿Qué quieres hacer?"
    ws.Range("B3").Font.Bold = True

    ' --- Desplegable con los temas (las opciones van en la hoja oculta Listas) --------
    Set wsListas = ObtenerHoja(HOJA_LISTAS)
    wsListas.Range("D1:D100").ClearContents
    For i = 1 To mTemas.Count
        wsListas.Cells(i, 4).Value = NombreTema(i)
    Next i
    ThisWorkbook.Names.Add Name:="ListaAyuda", RefersTo:="='" & wsListas.Name & "'!$D$1:$D$" & mTemas.Count

    With ws.Range(CELDA_ELEGIR)
        .Value = TEXTO_ELEGIR
        .Interior.Color = RGB(255, 242, 204)
        .Font.Bold = True
        .Borders.Color = RGB(197, 90, 17)
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                        Operator:=xlBetween, Formula1:="=ListaAyuda"
        .Validation.InputTitle = "¿Qué quieres hacer?"
        .Validation.InputMessage = "Elige un tema en la flecha de la derecha y pulsa 'Ver los pasos de este tema' o 'Llévame a la hoja'."
    End With

    ' --- Índice ------------------------------------------------------------------------
    fila = 6
    Encabezado ws, fila, "ÍNDICE"
    filaIndice = fila
    fila = fila + 1
    For i = 1 To mTemas.Count
        ws.Cells(fila, 2).Value = NombreTema(i)
        fila = fila + 1
    Next i
    fila = fila + 1

    ' --- Un bloque por tema --------------------------------------------------------------
    For i = 1 To mTemas.Count
        datos = mTemas(i)

        ' Enlace desde el índice hasta este bloque
        ws.Hyperlinks.Add Anchor:=ws.Cells(filaIndice + i, 2), Address:="", _
                          SubAddress:="'" & ws.Name & "'!B" & fila, TextToDisplay:=NombreTema(i)

        Encabezado ws, fila, NombreTema(i)
        fila = fila + 1

        ws.Cells(fila, 2).Value = "Para qué sirve: " & datos(T_DESC)
        ws.Cells(fila, 2).WrapText = True
        fila = fila + 1

        ws.Cells(fila, 2).Value = "Pasos:"
        ws.Cells(fila, 2).Font.Bold = True
        fila = fila + 1
        n = 0
        For Each linea In datos(T_PASOS)
            n = n + 1
            ws.Cells(fila, 2).Value = "   " & n & ".  " & linea
            ws.Cells(fila, 2).WrapText = True
            fila = fila + 1
        Next linea

        If datos(T_EJEMPLO).Count > 0 Then
            ws.Cells(fila, 2).Value = "Ejemplo:"
            ws.Cells(fila, 2).Font.Bold = True
            fila = fila + 1
            For Each linea In datos(T_EJEMPLO)
                With ws.Cells(fila, 2)
                    .Value = "'" & linea      ' El apóstrofo evita que Excel lo tome como fórmula
                    .Font.Name = "Consolas"
                    .Font.Size = 10
                    .Interior.Color = RGB(255, 242, 204)
                End With
                fila = fila + 1
            Next linea
        End If

        ' Enlaces: ir a la hoja y volver al índice
        If datos(T_HOJA) <> HOJA_AYUDA Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(fila, 2), Address:="", _
                              SubAddress:="'" & datos(T_HOJA) & "'!" & datos(T_CELDA), _
                              TextToDisplay:="» Ir a la hoja '" & datos(T_HOJA) & "'"
            fila = fila + 1
        End If
        ws.Hyperlinks.Add Anchor:=ws.Cells(fila, 2), Address:="", _
                          SubAddress:="'" & ws.Name & "'!B" & filaIndice, TextToDisplay:="^ Volver al índice"
        ws.Cells(fila, 2).Font.Size = 9
        fila = fila + 2
    Next i

    ws.Rows.AutoFit
    ws.Rows(1).RowHeight = 30

    ' Sin cuadrícula en la hoja de ayuda: se lee mejor
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range(CELDA_ELEGIR).Select

    Set PrepararHojaAyuda = ws
End Function


'------------------------------------------------------------------------------
' Encabezado: fila con fondo azul y letra blanca.
'------------------------------------------------------------------------------
Private Sub Encabezado(ByVal ws As Worksheet, ByVal fila As Long, ByVal texto As String)
    With ws.Cells(fila, 2)
        .Value = texto
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
End Sub


'------------------------------------------------------------------------------
' NombreTema: "3. Guardar el contenido completo en un .txt"
'------------------------------------------------------------------------------
Private Function NombreTema(ByVal i As Long) As String
    NombreTema = i & ". " & mTemas(i)(T_TITULO)
End Function


'------------------------------------------------------------------------------
' TemaElegido: nº del tema elegido en el desplegable (0 si no hay ninguno).
'------------------------------------------------------------------------------
Private Function TemaElegido() As Long
    Dim valor As String
    Dim i As Long

    CargarTemas
    valor = CStr(ObtenerHoja(HOJA_AYUDA).Range(CELDA_ELEGIR).Value)
    For i = 1 To mTemas.Count
        If valor = NombreTema(i) Then
            TemaElegido = i
            Exit Function
        End If
    Next i
End Function


'==============================================================================
' BOTONES
'==============================================================================

'------------------------------------------------------------------------------
' AyudaVerPasos  -> salta, dentro de la hoja Ayuda, al tema elegido.
'------------------------------------------------------------------------------
Public Sub AyudaVerPasos()
    Dim i As Long

    i = TemaElegido()
    If i = 0 Then
        AvisarElegir
        Exit Sub
    End If
    IrAlTema i
End Sub


'------------------------------------------------------------------------------
' AyudaIrAHoja  -> abre la hoja del tema elegido y recuerda los pasos.
'------------------------------------------------------------------------------
Public Sub AyudaIrAHoja()
    Dim i As Long

    i = TemaElegido()
    If i = 0 Then
        AvisarElegir
        Exit Sub
    End If
    LlevarATema i
End Sub


'------------------------------------------------------------------------------
' AyudaHoja  -> botón "? Ayuda de esta hoja": ayuda de la hoja en la que estás.
'------------------------------------------------------------------------------
Public Sub AyudaHoja()
    Dim temas As Variant
    Dim texto As String
    Dim i As Long
    Dim n As Long
    Dim respuesta As VbMsgBoxResult

    CargarTemas

    ' Temas que corresponden a cada hoja (números del índice)
    Select Case ActiveSheet.Name
        Case HOJA_BANDEJA:       temas = Array(2, 9, 10, 5, 3, 1)
        Case HOJA_ENVIAR:        temas = Array(4)
        Case HOJA_CARPETAS:      temas = Array(6, 8, 7, 5)
        Case HOJA_BUSCAR:        temas = Array(11, 3)
        Case HOJA_REGLAS:        temas = Array(12)
        Case HOJA_CLASIFICACION: temas = Array(13, 14)
        Case HOJA_PALABRAS:      temas = Array(14)
        Case Else:               temas = Array(15)
    End Select

    texto = "En esta hoja ('" & ActiveSheet.Name & "') puedes:" & vbCrLf & vbCrLf
    For i = LBound(temas) To UBound(temas)
        n = temas(i)
        texto = texto & "• " & mTemas(n)(T_TITULO) & vbCrLf
    Next i

    ' Si la hoja tiene un solo tema, mostramos también sus pasos
    If UBound(temas) = LBound(temas) Then
        texto = texto & vbCrLf & TextoPasos(CLng(temas(LBound(temas))))
    End If

    texto = texto & vbCrLf & "¿Quieres ver la ayuda completa, con pasos y ejemplos?"
    respuesta = MsgBox(texto, vbQuestion + vbYesNo, "Ayuda de esta hoja")
    If respuesta = vbYes Then IrAlTema CLng(temas(LBound(temas)))
End Sub


'------------------------------------------------------------------------------
' AyudaVisita  -> visita guiada: recorre las hojas explicando cada una.
'------------------------------------------------------------------------------
Public Sub AyudaVisita()
    Dim paradas As Variant
    Dim i As Long
    Dim hoja As String

    paradas = Array( _
        Array(HOJA_BANDEJA, "Aquí aparecen los correos que lees o buscas. Desde esta hoja puedes moverlos o copiarlos " & _
              "a carpetas (columna naranja 'Mover a'), marcarlos como leídos y guardar sus adjuntos."), _
        Array(HOJA_ENVIAR, "Escribe aquí los correos que quieres enviar: una fila por correo. Pulsa 'Enviar correos' " & _
              "y mira la columna Estado."), _
        Array(HOJA_CARPETAS, "Tu árbol de carpetas de Notes. Pulsa 'Actualizar carpetas', haz clic en una fila y " & _
              "usa los botones para leerla, crear subcarpetas, renombrarla o eliminarla."), _
        Array(HOJA_BUSCAR, "Busca por texto, remitente, fechas y carpeta. Rellena las celdas amarillas y pulsa 'Buscar'."), _
        Array(HOJA_REGLAS, "Reglas para archivar automáticamente la Bandeja de entrada: 'si el remitente contiene X, mover a la carpeta Y'."), _
        Array(HOJA_CLASIFICACION, "Clasificación por temas: importa la propuesta de Claude o usa las palabras clave, " & _
              "revísala y pulsa 'Aplicar clasificación'. Las carpetas se crean dentro de '" & CARPETA_CLASIFICACION & "'."), _
        Array(HOJA_PALABRAS, "Las palabras clave de la clasificación automática. Puedes añadir o cambiar reglas."), _
        Array(HOJA_AYUDA, "Y esta es la Ayuda: índice, pasos y ejemplos de todo. ¡Fin de la visita!"))

    For i = LBound(paradas) To UBound(paradas)
        hoja = paradas(i)(0)
        If HojaExiste(hoja) Then
            ThisWorkbook.Worksheets(hoja).Activate
            ThisWorkbook.Worksheets(hoja).Range("A1").Select
            If MsgBox("Hoja '" & hoja & "'  (" & (i + 1) & " de " & (UBound(paradas) + 1) & ")" & vbCrLf & vbCrLf & _
                      paradas(i)(1) & vbCrLf & vbCrLf & _
                      "Aceptar = siguiente hoja      Cancelar = terminar la visita", _
                      vbInformation + vbOKCancel, "Visita guiada") = vbCancel Then Exit Sub
        End If
    Next i
End Sub


'==============================================================================
' FUNCIONES INTERNAS
'==============================================================================

'------------------------------------------------------------------------------
' IrAlTema: muestra en la hoja Ayuda el bloque del tema nº i.
'------------------------------------------------------------------------------
Private Sub IrAlTema(ByVal i As Long)
    Dim ws As Worksheet
    Dim celda As Range

    Set ws = ObtenerHoja(HOJA_AYUDA)
    ' Buscamos el encabezado del tema (la 2ª vez que aparece: la 1ª es el índice)
    Set celda = ws.Columns(2).Find(What:=NombreTema(i), LookIn:=xlValues, LookAt:=xlWhole)
    If Not celda Is Nothing Then
        Set celda = ws.Columns(2).FindNext(celda)
    End If

    ws.Activate
    If celda Is Nothing Then
        MsgBox "No encuentro ese tema en la hoja Ayuda. Ejecuta PrepararLibro para regenerarla.", _
               vbExclamation, "Ayuda"
    Else
        Application.Goto ws.Cells(celda.Row, 2), Scroll:=True
    End If
End Sub


'------------------------------------------------------------------------------
' LlevarATema: abre la hoja del tema, selecciona la celda y recuerda los pasos.
'------------------------------------------------------------------------------
Private Sub LlevarATema(ByVal i As Long)
    Dim datos As Variant

    CargarTemas
    datos = mTemas(i)

    If HojaExiste(CStr(datos(T_HOJA))) Then
        Application.Goto ThisWorkbook.Worksheets(CStr(datos(T_HOJA))).Range(CStr(datos(T_CELDA))), Scroll:=False
    End If
    MsgBox TextoPasos(i), vbInformation, NombreTema(i)
End Sub


'------------------------------------------------------------------------------
' TextoPasos: los pasos de un tema, numerados, para una ventana de mensaje.
'------------------------------------------------------------------------------
Private Function TextoPasos(ByVal i As Long) As String
    Dim datos As Variant
    Dim linea As Variant
    Dim texto As String
    Dim n As Long

    datos = mTemas(i)
    texto = datos(T_DESC) & vbCrLf & vbCrLf & "Pasos:" & vbCrLf
    For Each linea In datos(T_PASOS)
        n = n + 1
        texto = texto & "  " & n & ". " & linea & vbCrLf
    Next linea

    ' Las ventanas de mensaje admiten unos 1.000 caracteres
    If Len(texto) > 950 Then texto = Left$(texto, 940) & "..."
    TextoPasos = texto
End Function


Private Sub AvisarElegir()
    ObtenerHoja(HOJA_AYUDA).Activate
    ObtenerHoja(HOJA_AYUDA).Range(CELDA_ELEGIR).Select
    MsgBox "Primero elige un tema en la celda amarilla (haz clic en ella y usa la flecha de la derecha).", _
           vbInformation, "Ayuda"
End Sub


'==============================================================================
' NOTAS EN LOS ENCABEZADOS (aparecen al pasar el ratón por encima)
'==============================================================================
Public Sub PonerNotasAyuda()
    ' --- Bandeja ---
    Nota HOJA_BANDEJA, 1, COL_B_ID, "ID interno del correo." & vbLf & "No lo borres: lo usan Mover, Marcar y Guardar adjuntos."
    Nota HOJA_BANDEJA, 1, COL_B_CARPETA, "Carpeta en la que está el correo." & vbLf & "Al MOVER, el correo sale de esta carpeta."
    Nota HOJA_BANDEJA, 1, COL_B_LEIDO, "Sí / No. Puede salir vacío en versiones antiguas de Notes."
    Nota HOJA_BANDEJA, 1, COL_B_MOVER, "Elige aquí la carpeta de destino con el desplegable." & vbLf & _
                                        "Luego pulsa 'Mover a carpeta' o 'Copiar a carpeta'." & vbLf & "Ej.: Facturas\2026"
    Nota HOJA_BANDEJA, 1, COL_B_RESULTADO, "Lo rellena la macro: 'Movido a...', 'Copiado a...' o el error."

    ' --- Enviar ---
    Nota HOJA_ENVIAR, 1, COL_E_PARA, "Obligatorio. Varios separados por ;" & vbLf & "Ej.: ana@empresa.com; luis@empresa.com"
    Nota HOJA_ENVIAR, 1, COL_E_CC, "Opcional. Igual que 'Para'."
    Nota HOJA_ENVIAR, 1, COL_E_CUERPO, "Texto del correo." & vbLf & "Alt+Intro = salto de línea dentro de la celda."
    Nota HOJA_ENVIAR, 1, COL_E_ADJUNTO, "Opcional. Ruta completa del archivo; varios separados por ;" & vbLf & "Ej.: C:\Informes\semana.pdf"
    Nota HOJA_ENVIAR, 1, COL_E_ESTADO, "Déjala VACÍA para que se envíe." & vbLf & "La macro escribe 'Enviado' o 'ERROR'." & vbLf & "Para reenviar, bórrala."

    ' --- Carpetas ---
    Nota HOJA_CARPETAS, 1, COL_C_RUTA, "Ruta completa de la carpeta; los niveles se separan con \" & vbLf & "Ej.: Clientes\Acme\2026"
    Nota HOJA_CARPETAS, 1, COL_C_TIPO, "Sistema = Bandeja / Enviados" & vbLf & "Carpeta = tuya" & vbLf & "Agrupación = 'padre' que no existe como carpeta"
    Nota HOJA_CARPETAS, 1, COL_C_NUEVA, "Escribe una ruta por fila y pulsa 'Crear carpetas de la lista'." & vbLf & "Ej.: Clientes\Acme"

    ' --- Buscar ---
    Nota HOJA_BUSCAR, FILA_BUS_TEXTO, 1, "Busca en el asunto y en el cuerpo. No distingue mayúsculas." & vbLf & "Ej.: presupuesto"
    Nota HOJA_BUSCAR, FILA_BUS_REMITENTE, 1, "Vale el nombre o la dirección." & vbLf & "Ej.: acme.com   o   Juan"
    Nota HOJA_BUSCAR, FILA_BUS_DESDE, 1, "Fecha inicial, incluida." & vbLf & "Ej.: 01/01/2026"
    Nota HOJA_BUSCAR, FILA_BUS_HASTA, 1, "Fecha final, incluida. Vacía = hasta hoy."
    Nota HOJA_BUSCAR, FILA_BUS_CARPETA, 1, "Vacía = Bandeja, Enviados y todas tus carpetas."
    Nota HOJA_BUSCAR, FILA_BUS_SUBCARPETAS, 1, "Solo cuenta si eliges una carpeta."

    ' --- Reglas ---
    Nota HOJA_REGLAS, 1, COL_R_CAMPO, "Dónde buscar: Remitente, Para/CC, Asunto o Cuerpo."
    Nota HOJA_REGLAS, 1, COL_R_CONTIENE, "Texto a buscar (sin distinguir mayúsculas)." & vbLf & "Ej.: @acme.com"
    Nota HOJA_REGLAS, 1, COL_R_DESTINO, "Carpeta a la que se mueve el correo (desplegable)."
    Nota HOJA_REGLAS, 1, COL_R_ACTIVA, "No = la regla no se aplica (sin borrarla)."

    ' --- Clasificacion ---
    Nota HOJA_CLASIFICACION, 1, COL_K_CARPETA, "Carpeta dentro de '" & CARPETA_CLASIFICACION & "'. Puedes cambiarla." & vbLf & "Ej.: RRHH\Sindicatos"
    Nota HOJA_CLASIFICACION, 1, COL_K_APLICAR, "Sí = se copia al aplicar." & vbLf & "No = no se toca." & vbLf & "Hecho = ya aplicado."
    Nota HOJA_CLASIFICACION, 1, COL_K_MOTIVO, "Por qué se propone esa carpeta: archivo de Claude o fila de 'Palabras clave'."

    ' --- Palabras clave ---
    Nota HOJA_PALABRAS, 1, COL_P_CARPETA, "Carpeta dentro de '" & CARPETA_CLASIFICACION & "'." & vbLf & "Ej.: Informatica\Despliegues CEAR"
    Nota HOJA_PALABRAS, 1, COL_P_CAMPO, "Asunto, Remitente, Cuerpo o 'Asunto o cuerpo'."
    Nota HOJA_PALABRAS, 1, COL_P_PALABRAS, "Palabras separadas por ;  Basta con que aparezca UNA." & vbLf & _
                                            "Sin distinguir mayúsculas ni tildes." & vbLf & "Ej.: CEAR;I26-;parches"
    Nota HOJA_PALABRAS, 1, COL_P_ACTIVA, "No = la regla no se usa." & vbLf & "Gana la PRIMERA regla (de arriba abajo) que se cumple."
End Sub


'------------------------------------------------------------------------------
' Nota: pone (o reemplaza) la nota de una celda, si la hoja existe.
'------------------------------------------------------------------------------
Private Sub Nota(ByVal hoja As String, ByVal fila As Long, ByVal columna As Long, ByVal texto As String)
    Dim celda As Range

    If Not HojaExiste(hoja) Then Exit Sub
    On Error Resume Next
    Set celda = ThisWorkbook.Worksheets(hoja).Cells(fila, columna)
    celda.ClearComments
    celda.AddComment texto
    celda.Comment.Shape.TextFrame.AutoSize = True
    On Error GoTo 0
End Sub

Attribute VB_Name = "modConfiguracion"
'==============================================================================
' MÓDULO: modConfiguracion
'------------------------------------------------------------------------------
' Aquí están TODOS los valores que quizá quieras cambiar.
' Si necesitas ajustar algo (nombre de una hoja, número de correos a leer...)
' cámbialo SOLO aquí y el resto del código lo usará automáticamente.
'
' "Const" significa "constante": un valor con nombre que no cambia mientras
' se ejecuta la macro. "Public" hace que se pueda usar desde otros módulos.
'==============================================================================
Option Explicit   ' Obliga a declarar todas las variables (evita errores de escritura)

'--- Nombres de las hojas de Excel ------------------------------------------
Public Const HOJA_BANDEJA As String = "Bandeja"   ' Hoja donde se vuelcan los correos leídos
Public Const HOJA_ENVIAR As String = "Enviar"     ' Hoja con los correos a enviar
Public Const HOJA_CARPETAS As String = "Carpetas" ' Árbol de carpetas de tu buzón
Public Const HOJA_BUSCAR As String = "Buscar"     ' Criterios de búsqueda
Public Const HOJA_REGLAS As String = "Reglas"     ' Reglas de archivado automático
Public Const HOJA_LISTAS As String = "Listas"     ' Hoja OCULTA con las opciones de los desplegables

'--- Lectura de la bandeja de entrada -----------------------------------------
' Número máximo de correos que se leen (empezando por los más recientes).
' Si tienes miles de correos, un número alto hará que tarde mucho.
Public Const MAX_CORREOS As Long = 100

' Número de caracteres del cuerpo del correo que se copian a la hoja.
Public Const LONGITUD_EXTRACTO As Long = 200

'--- Lectura de carpetas y búsquedas ------------------------------------------
' Máximo de correos que se copian a la hoja al leer una carpeta (con sus
' subcarpetas) o al buscar. Las carpetas grandes pueden tardar bastante.
Public Const MAX_CORREOS_CARPETA As Long = 500

'--- Carpetas del sistema -----------------------------------------------------
' Nombres "bonitos" con los que aparecen en Excel las carpetas especiales de
' Notes (por dentro se llaman "($Inbox)" y "($Sent)").
Public Const CARPETA_BANDEJA As String = "Bandeja de entrada"
Public Const CARPETA_ENVIADOS As String = "Enviados"

' Nombre del rango con la lista de carpetas que usan los desplegables (no tocar)
Public Const NOMBRE_LISTA_CARPETAS As String = "ListaCarpetas"

'--- Base de correo -----------------------------------------------------------
' Normalmente se dejan VACÍOS ("") y el programa averigua solo cuál es tu
' base de correo (leyendo la configuración de tu cliente de Notes).
' Rellénalos solo si la detección automática falla. Ejemplo:
'     SERVIDOR_CORREO = "ServidorCorreo/MiEmpresa"
'     ARCHIVO_CORREO  = "mail\jperez.nsf"
' Para una réplica LOCAL deja SERVIDOR_CORREO = "" y pon solo el archivo.
Public Const SERVIDOR_CORREO As String = ""
Public Const ARCHIVO_CORREO As String = ""

'--- Botón de conexión ---------------------------------------------------------
' El mismo botón sirve para conectar y desconectar: su texto cambia solo
' según el estado de la conexión.
Public Const TEXTO_CONECTAR As String = "Conectar a Notes"
Public Const TEXTO_DESCONECTAR As String = "Desconectar de Notes"
Public Const BOTON_CONEXION As String = "btn_ConectarDesconectarNotes"   ' Nombre interno (no tocar)

'--- Columnas de la hoja "Bandeja" (1 = A, 2 = B, ...) ------------------------
Public Const COL_B_FECHA As Long = 1       ' A: Fecha
Public Const COL_B_REMITENTE As Long = 2   ' B: Remitente
Public Const COL_B_ASUNTO As Long = 3      ' C: Asunto
Public Const COL_B_EXTRACTO As Long = 4    ' D: Extracto del cuerpo
Public Const COL_B_ADJUNTOS As Long = 5    ' E: Nº de adjuntos
Public Const COL_B_ID As Long = 6          ' F: Identificador interno del correo (no borrar)
Public Const COL_B_CARPETA As Long = 7     ' G: Carpeta en la que está el correo
Public Const COL_B_LEIDO As Long = 8       ' H: Leído (Sí / No)
Public Const COL_B_MOVER As Long = 9       ' I: Mover a (desplegable con tus carpetas)
Public Const COL_B_RESULTADO As Long = 10  ' J: Resultado de mover / copiar / marcar

'--- Columnas de la hoja "Carpetas" -------------------------------------------
Public Const COL_C_NOMBRE As Long = 1      ' A: Nombre (con sangría según el nivel)
Public Const COL_C_RUTA As Long = 2        ' B: Ruta completa (ej. Clientes\Acme\2026)
Public Const COL_C_TOTAL As Long = 3       ' C: Nº de correos
Public Const COL_C_NOLEIDOS As Long = 4    ' D: No leídos
Public Const COL_C_TIPO As Long = 5        ' E: Tipo (Sistema / Carpeta / Agrupación)
Public Const COL_C_NUEVA As Long = 7       ' G: Lista de carpetas a crear de golpe
Public Const COL_C_NUEVA_RES As Long = 8   ' H: Resultado de crearlas

'--- Columnas de la hoja "Reglas" ---------------------------------------------
Public Const COL_R_CAMPO As Long = 1       ' A: Dónde buscar (Remitente, Para/CC, Asunto, Cuerpo)
Public Const COL_R_CONTIENE As Long = 2    ' B: Texto que debe contener
Public Const COL_R_DESTINO As Long = 3     ' C: Carpeta de destino (desplegable)
Public Const COL_R_ACTIVA As Long = 4      ' D: Activa (Sí / No)
Public Const COL_R_RESULTADO As Long = 5   ' E: Resultado de la última vez

'--- Filas de la hoja "Buscar" (los valores van en la columna B) --------------
Public Const FILA_BUS_TEXTO As Long = 3
Public Const FILA_BUS_REMITENTE As Long = 4
Public Const FILA_BUS_DESDE As Long = 5
Public Const FILA_BUS_HASTA As Long = 6
Public Const FILA_BUS_CARPETA As Long = 7
Public Const FILA_BUS_SUBCARPETAS As Long = 8

'--- Columnas de la hoja "Enviar" ---------------------------------------------
Public Const COL_E_PARA As Long = 1        ' A: Para (varias direcciones separadas por ;)
Public Const COL_E_CC As Long = 2          ' B: CC   (opcional)
Public Const COL_E_ASUNTO As Long = 3      ' C: Asunto
Public Const COL_E_CUERPO As Long = 4      ' D: Cuerpo del mensaje
Public Const COL_E_ADJUNTO As Long = 5     ' E: Ruta del adjunto (opcional, varias separadas por ;)
Public Const COL_E_ESTADO As Long = 6      ' F: Estado (lo rellena la macro)

'--- Códigos de error propios -------------------------------------------------
' vbObjectError es la base que VBA recomienda para errores "inventados" por
' nosotros, así no se confunden con los errores propios de VBA.
Public Const ERR_CANCELADO As Long = vbObjectError + 1001       ' El usuario canceló
Public Const ERR_SIN_NOTES As Long = vbObjectError + 1002       ' No se encuentra Notes
Public Const ERR_CONTRASENA As Long = vbObjectError + 1003      ' Contraseña incorrecta
Public Const ERR_BASE_CORREO As Long = vbObjectError + 1004     ' No se abre la base de correo
Public Const ERR_DATOS As Long = vbObjectError + 1005           ' Datos de la hoja incorrectos

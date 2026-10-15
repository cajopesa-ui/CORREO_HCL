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

'--- Lectura de la bandeja de entrada -----------------------------------------
' Número máximo de correos que se leen (empezando por los más recientes).
' Si tienes miles de correos, un número alto hará que tarde mucho.
Public Const MAX_CORREOS As Long = 100

' Número de caracteres del cuerpo del correo que se copian a la hoja.
Public Const LONGITUD_EXTRACTO As Long = 200

'--- Base de correo -----------------------------------------------------------
' Normalmente se dejan VACÍOS ("") y el programa averigua solo cuál es tu
' base de correo (leyendo la configuración de tu cliente de Notes).
' Rellénalos solo si la detección automática falla. Ejemplo:
'     SERVIDOR_CORREO = "ServidorCorreo/MiEmpresa"
'     ARCHIVO_CORREO  = "mail\jperez.nsf"
' Para una réplica LOCAL deja SERVIDOR_CORREO = "" y pon solo el archivo.
Public Const SERVIDOR_CORREO As String = ""
Public Const ARCHIVO_CORREO As String = ""

'--- Columnas de la hoja "Bandeja" (1 = A, 2 = B, ...) ------------------------
Public Const COL_B_FECHA As Long = 1       ' A: Fecha
Public Const COL_B_REMITENTE As Long = 2   ' B: Remitente
Public Const COL_B_ASUNTO As Long = 3      ' C: Asunto
Public Const COL_B_EXTRACTO As Long = 4    ' D: Extracto del cuerpo
Public Const COL_B_ADJUNTOS As Long = 5    ' E: Nº de adjuntos
Public Const COL_B_ID As Long = 6          ' F: Identificador interno del correo (no borrar)

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

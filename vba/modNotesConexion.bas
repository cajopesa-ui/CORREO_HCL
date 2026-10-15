Attribute VB_Name = "modNotesConexion"
'==============================================================================
' MÓDULO: modNotesConexion
'------------------------------------------------------------------------------
' Se encarga de "hablar" con HCL Notes:
'   - ObtenerSesion    -> abre una sesión de Notes (pide contraseña si hace falta)
'   - AbrirBaseCorreo  -> abre tu base de correo (tu buzón, el archivo .nsf)
'   - DesconectarNotes -> olvida la sesión (útil si cambias de usuario)
'   - MostrarError     -> muestra los errores en español con pistas de solución
'
' ¿QUÉ ES "LATE BINDING"?
'   Para usar Notes desde Excel podríamos añadir una "referencia" en
'   Herramientas > Referencias. Pero si otro PC tiene otra versión de Notes,
'   el libro fallaría. Por eso declaramos los objetos "As Object" y los creamos
'   con CreateObject(...). Así el libro funciona sin tocar referencias.
'   La única pega: el editor no te sugiere los métodos mientras escribes.
'==============================================================================
Option Explicit

' Constantes de la API de Notes. Con late binding no conocemos sus nombres
' oficiales, así que las definimos nosotros con el mismo valor numérico.
Public Const NOTES_RICHTEXT As Long = 1             ' Tipo de campo "texto enriquecido"
Public Const NOTES_EMBED_ATTACHMENT As Long = 1454  ' Tipo de objeto "archivo adjunto"

' Variables "de módulo" (declaradas fuera de cualquier Sub). Conservan su valor
' mientras el libro esté abierto, así NO pedimos la contraseña en cada macro.
Private mSesion As Object        ' La sesión de Notes
Private mBaseCorreo As Object    ' La base de correo ya abierta


'------------------------------------------------------------------------------
' ObtenerSesion
' Devuelve una sesión de Notes lista para usar. Si ya había una, la reutiliza.
'------------------------------------------------------------------------------
Public Function ObtenerSesion() As Object
    Dim sesion As Object
    Dim clave As String
    Dim numError As Long
    Dim descError As String

    ' Si ya tenemos sesión abierta, la devolvemos y terminamos.
    If Not mSesion Is Nothing Then
        Set ObtenerSesion = mSesion
        Exit Function
    End If

    '--- 1) Intento con "Lotus.NotesSession" (API COM, no abre ventanas) ------
    ' "On Error Resume Next" = si algo falla, sigue en la línea siguiente
    ' sin detenerse. Luego miramos Err.Number para saber si falló.
    On Error Resume Next
    Set sesion = CreateObject("Lotus.NotesSession")
    On Error GoTo 0   ' Volvemos al tratamiento de errores normal

    If Not sesion Is Nothing Then
        ' Esta clase hay que "inicializarla". Primero lo intentamos SIN
        ' contraseña: funciona si el cliente Notes ya está abierto y tiene
        ' activado "Compartir la contraseña con programas de Notes", o si
        ' Notes muestra su propia ventana para pedirla.
        On Error Resume Next
        sesion.Initialize
        numError = Err.Number
        On Error GoTo 0

        If numError <> 0 Then
            ' No ha funcionado: pedimos la contraseña al usuario.
            clave = PedirContrasena()
            If Len(clave) = 0 Then
                Err.Raise ERR_CANCELADO, "ObtenerSesion", _
                          "Operación cancelada: no se introdujo la contraseña."
            End If

            On Error Resume Next
            sesion.Initialize clave
            numError = Err.Number
            descError = Err.Description
            On Error GoTo 0
            clave = ""   ' Borramos la contraseña de la memoria cuanto antes

            If numError <> 0 Then
                Err.Raise ERR_CONTRASENA, "ObtenerSesion", _
                          "No se pudo iniciar sesión en HCL Notes. " & _
                          "¿La contraseña es correcta?" & vbCrLf & _
                          "Detalle técnico: " & descError
            End If
        End If

    Else
        '--- 2) Plan B: "Notes.NotesSession" (abre el cliente Notes) ----------
        ' Esta clase usa el programa Notes completo. Es más lenta, pero suele
        ' funcionar aunque Excel sea de 64 bits y Notes de 32 bits.
        ' Notes pedirá la contraseña en su propia ventana si hace falta.
        On Error Resume Next
        Set sesion = CreateObject("Notes.NotesSession")
        On Error GoTo 0

        If sesion Is Nothing Then
            Err.Raise ERR_SIN_NOTES, "ObtenerSesion", _
                      "No se pudo conectar con HCL Notes desde Excel."
        End If
    End If

    ' Guardamos la sesión para la próxima vez y la devolvemos.
    Set mSesion = sesion
    Set ObtenerSesion = sesion
End Function


'------------------------------------------------------------------------------
' AbrirBaseCorreo
' Devuelve tu base de correo (buzón) abierta.
'------------------------------------------------------------------------------
Public Function AbrirBaseCorreo() As Object
    Dim sesion As Object
    Dim db As Object
    Dim directorio As Object
    Dim servidor As String
    Dim archivo As String

    ' ¿Ya estaba abierta? La reutilizamos.
    If Not mBaseCorreo Is Nothing Then
        Set AbrirBaseCorreo = mBaseCorreo
        Exit Function
    End If

    Set sesion = ObtenerSesion()

    ' ¿Dónde está el buzón? Si no se indicó en modConfiguracion, lo leemos de
    ' la configuración del cliente Notes (archivo notes.ini):
    '   MailServer = servidor de correo,  MailFile = ruta del archivo .nsf
    If Len(ARCHIVO_CORREO) > 0 Then
        servidor = SERVIDOR_CORREO
        archivo = ARCHIVO_CORREO
    Else
        On Error Resume Next
        servidor = sesion.GetEnvironmentString("MailServer", True)
        archivo = sesion.GetEnvironmentString("MailFile", True)
        On Error GoTo 0
    End If

    ' Intento 1: abrir con servidor + archivo.
    If Len(archivo) > 0 Then
        On Error Resume Next
        Set db = sesion.GetDatabase(servidor, archivo, False)
        If Not db Is Nothing Then
            If Not db.IsOpen Then db.Open servidor, archivo
        End If
        On Error GoTo 0
    End If

    ' Intento 2: pedir a Notes directamente "la base de correo del usuario".
    If Not BaseAbierta(db) Then
        On Error Resume Next
        Set directorio = sesion.GetDbDirectory("")
        Set db = directorio.OpenMailDatabase()
        On Error GoTo 0
    End If

    If Not BaseAbierta(db) Then
        Err.Raise ERR_BASE_CORREO, "AbrirBaseCorreo", _
                  "No se pudo abrir tu base de correo." & vbCrLf & _
                  "Servidor: " & IIf(Len(servidor) = 0, "(local)", servidor) & vbCrLf & _
                  "Archivo: " & IIf(Len(archivo) = 0, "(desconocido)", archivo)
    End If

    Set mBaseCorreo = db
    Set AbrirBaseCorreo = db
End Function


'------------------------------------------------------------------------------
' BaseAbierta: True si la base existe y está abierta. (Uso interno)
'------------------------------------------------------------------------------
Private Function BaseAbierta(ByVal db As Object) As Boolean
    BaseAbierta = False
    If db Is Nothing Then Exit Function
    On Error Resume Next
    BaseAbierta = db.IsOpen
    On Error GoTo 0
End Function


'------------------------------------------------------------------------------
' PedirContrasena: pregunta la contraseña de Notes. (Uso interno)
' OJO: el InputBox de Excel NO oculta lo que escribes con asteriscos.
' Devuelve "" si el usuario pulsa Cancelar.
'------------------------------------------------------------------------------
Private Function PedirContrasena() As String
    PedirContrasena = InputBox( _
        "Introduce tu contraseña de HCL Notes." & vbCrLf & vbCrLf & _
        "(Aviso: el texto se verá mientras lo escribes)", _
        "Conexión con HCL Notes")
End Function


'------------------------------------------------------------------------------
' DesconectarNotes
' Olvida la sesión guardada. La próxima macro volverá a conectarse.
' Puedes asignarla a un botón o ejecutarla con Alt+F8.
'------------------------------------------------------------------------------
Public Sub DesconectarNotes()
    Set mBaseCorreo = Nothing
    Set mSesion = Nothing
    MsgBox "Desconectado de HCL Notes.", vbInformation, "HCL Notes"
End Sub


'------------------------------------------------------------------------------
' MostrarError
' Muestra un mensaje de error comprensible, con pistas para solucionarlo.
'   titulo      -> qué estábamos haciendo (ej. "Leer la bandeja de entrada")
'   numero      -> Err.Number
'   descripcion -> Err.Description
'------------------------------------------------------------------------------
Public Sub MostrarError(ByVal titulo As String, ByVal numero As Long, _
                        ByVal descripcion As String)
    Dim pista As String
    Dim textoMinusculas As String

    ' Si el usuario canceló, no es un error: solo avisamos suavemente.
    If numero = ERR_CANCELADO Then
        MsgBox descripcion, vbInformation, titulo
        Exit Sub
    End If

    textoMinusculas = LCase$(descripcion)

    Select Case numero
        Case 429, ERR_SIN_NOTES
            ' 429 = "ActiveX no puede crear el objeto"
            pista = "Comprueba que:" & vbCrLf & _
                    "  - El cliente HCL Notes está instalado y configurado en este PC." & vbCrLf & _
                    "  - Excel y Notes son de la misma arquitectura (32 o 64 bits)." & vbCrLf & _
                    "  - Has abierto Notes al menos una vez con tu usuario." & vbCrLf & _
                    "Consulta el apartado 'Solución de problemas' del README."
        Case ERR_CONTRASENA
            pista = "Revisa la contraseña (mayúsculas/minúsculas) y vuelve a intentarlo."
        Case ERR_BASE_CORREO
            pista = "Comprueba que tienes conexión con el servidor de correo, " & _
                    "o indica SERVIDOR_CORREO y ARCHIVO_CORREO en el módulo modConfiguracion."
            ' Si no se pudo abrir, olvidamos la sesión para empezar limpio.
            Set mBaseCorreo = Nothing
        Case ERR_DATOS
            pista = "Revisa los datos de la hoja."
        Case Else
            ' Errores que vienen de Notes: buscamos palabras clave.
            If InStr(textoMinusculas, "password") > 0 Then
                pista = "Parece un problema con la contraseña de Notes."
            ElseIf InStr(textoMinusculas, "server") > 0 Then
                pista = "Parece un problema de conexión con el servidor de Notes."
            ElseIf InStr(textoMinusculas, "not authorized") > 0 Or _
                   InStr(textoMinusculas, "no autorizado") > 0 Then
                pista = "Tu usuario no tiene permisos para esta operación."
            Else
                pista = "Si el problema continúa, cierra Excel y Notes y vuelve a abrirlos."
            End If
            ' Un error inesperado puede dejar la sesión en mal estado: la olvidamos.
            Set mBaseCorreo = Nothing
            Set mSesion = Nothing
    End Select

    MsgBox "Se ha producido un error." & vbCrLf & vbCrLf & _
           descripcion & vbCrLf & vbCrLf & _
           pista & vbCrLf & vbCrLf & _
           "(Código de error: " & numero & ")", _
           vbExclamation, titulo
End Sub

# Correo HCL Notes desde Excel (VBA)

Macros de Excel que se conectan a **HCL Notes** (antes Lotus Notes) para:

- 📥 **Leer la Bandeja de entrada** y volcarla en una hoja: Fecha, Remitente, Asunto, extracto del cuerpo y nº de adjuntos.
- 📤 **Enviar correos** (Para, CC, Asunto, Cuerpo y adjuntos opcionales) a partir de las filas de una hoja.
- 📎 **Guardar los adjuntos** de los correos leídos en una carpeta de tu PC.

El código está muy comentado en español y pensado para quien empieza con VBA.

---

## 1. Requisitos

| Requisito | Detalle |
|---|---|
| Sistema operativo | **Windows** (no funciona en Mac ni en Excel Online). |
| Excel | **Excel de escritorio** (Microsoft 365, 2016, 2019, 2021…). |
| HCL Notes | El **cliente HCL Notes instalado y configurado** con tu usuario en ese mismo PC. Ábrelo al menos una vez y comprueba que ves tu correo. |
| Arquitectura | **Excel y Notes deben ser de la misma arquitectura (32 o 64 bits)**. Ver [cómo comprobarlo](#cómo-saber-si-excel-y-notes-son-de-32-o-64-bits). |
| Macros | Tener permitidas las macros en Excel. |

> No hace falta añadir ninguna referencia en *Herramientas > Referencias*: el código usa *late binding* (`CreateObject`), así funciona con cualquier versión de Notes.

---

## 2. Archivos del proyecto

```
vba/
├── modConfiguracion.bas   ← valores que puedes cambiar (nombres de hojas, nº de correos…)
├── modNotesConexion.bas   ← conexión con Notes, contraseña y mensajes de error
├── modLeerCorreo.bas      ← macro LeerBandejaEntrada
├── modEnviarCorreo.bas    ← macro EnviarCorreos (y la función EnviarUnCorreo)
├── modAdjuntos.bas        ← macro GuardarAdjuntos
└── modPreparar.bas        ← macro PrepararLibro (crea hojas y botones)
```

> Un archivo `.xlsm` es binario y no se puede generar de forma fiable aquí, por eso
> los módulos se entregan como texto (`.bas`) para que los importes tú en Excel.
> Los `.bas` están guardados con codificación **Windows-1252** y saltos de línea
> **CRLF**, que es lo que espera el editor de VBA (así se ven bien las tildes y la ñ).

---

## 3. Instalación paso a paso

### 3.1 Descargar los archivos

En GitHub pulsa **Code > Download ZIP** y descomprímelo. (Si copias el código
a mano desde la web, las tildes podrían verse mal; es mejor importar los `.bas`.)

### 3.2 Crear el libro e importar los módulos

1. Abre Excel y crea un **libro en blanco**.
2. Pulsa **Alt + F11** para abrir el editor de VBA.
3. En el menú: **Archivo > Importar archivo…** (o *File > Import File…*).
4. Selecciona `modConfiguracion.bas` y pulsa **Abrir**.
5. Repite el paso 3–4 con **cada uno** de los 6 archivos `.bas` de la carpeta `vba`.
6. En el panel izquierdo (*Explorador de proyectos*) verás una carpeta **Módulos** con los 6 módulos.
7. *(Recomendado)* Menú **Depuración > Compilar VBAProject**. Si no sale ningún mensaje, todo está bien.

### 3.3 Guardar como libro con macros (.xlsm)

1. Vuelve a Excel (**Alt + F11** otra vez, o cierra el editor).
2. **Archivo > Guardar como**.
3. En *Tipo*, elige **Libro de Excel habilitado para macros (\*.xlsm)**.
4. Guarda. ⚠️ Si lo guardas como `.xlsx` **se perderán las macros**.

### 3.4 Preparar las hojas y los botones (automático)

1. Pulsa **Alt + F8**, elige **PrepararLibro** y pulsa **Ejecutar**.
2. Se crearán:
   - La hoja **Bandeja** con sus encabezados y los botones *Leer bandeja de entrada*, *Guardar adjuntos*, *Enviar correos* y *Conectar a Notes*.
   - La hoja **Enviar** con sus encabezados, una fila de ejemplo y los botones *Enviar correos* y *Conectar a Notes*.
3. Guarda el libro (**Ctrl + G** / **Ctrl + S**).

<details>
<summary>¿Prefieres crear los botones a mano?</summary>

1. Activa la pestaña **Programador** (*Archivo > Opciones > Personalizar cinta de opciones* > marca *Programador*).
2. **Programador > Insertar > Botón (control de formulario)**.
3. Dibuja el botón en la hoja; Excel te pedirá la macro: elige `LeerBandejaEntrada`, `EnviarCorreos`, `GuardarAdjuntos` o `ConectarDesconectarNotes`.
4. Clic derecho en el botón > *Modificar texto* para cambiarle el nombre.
</details>

---

## 4. Uso

### 4.1 Leer la bandeja de entrada

1. Ve a la hoja **Bandeja** y pulsa **Leer bandeja de entrada**.
2. La primera vez se conectará con Notes:
   - Si Notes está abierto y tiene activada la opción de compartir la contraseña, entra directamente.
   - Si no, te pedirá la **contraseña de Notes** (en una ventana de Notes o de Excel).
3. Se mostrarán los **100 correos más recientes** (cambia `MAX_CORREOS` en `modConfiguracion` si quieres más o menos).

| A Fecha | B Remitente | C Asunto | D Extracto del cuerpo | E Nº adjuntos | F ID (no tocar) |
|---|---|---|---|---|---|

> La columna **F** guarda un identificador interno del correo. La necesita la macro de guardar adjuntos: no la borres.

### 4.2 Enviar correos

Rellena la hoja **Enviar**, **una fila por correo**:

| Columna | Contenido | Ejemplo |
|---|---|---|
| A **Para** | Destinatario(s), separados por `;` | `ana@empresa.com; luis@empresa.com` |
| B **CC** | Copia (opcional) | `jefe@empresa.com` |
| C **Asunto** | Asunto | `Informe semanal` |
| D **Cuerpo** | Texto. Usa **Alt + Intro** para saltos de línea | `Hola, adjunto el informe…` |
| E **Adjunto** | Ruta completa del archivo (opcional). Varios separados por `;` | `C:\Informes\semana.pdf` |
| F **Estado** | **Déjala vacía**. La macro escribe aquí el resultado | `Enviado 15/10/2026 10:30` |

Pulsa **Enviar correos**. La macro:

- Solo envía las filas con *Para* relleno y *Estado* **vacío** (así no envía dos veces el mismo correo).
- Pide confirmación antes de enviar.
- Escribe en *Estado* «Enviado + fecha» (verde) o «ERROR: …» (rojo). Para reenviar una fila, borra su *Estado*.
- Guarda una copia en tu carpeta **Enviados** de Notes.

**Desde tus propias macros** puedes enviar un correo con una sola línea:

```vb
EnviarUnCorreo "ana@empresa.com", "", "Hola", "Texto del mensaje", "C:\informe.pdf"
```

### 4.3 Guardar adjuntos

1. Primero lee la bandeja (apartado 4.1).
2. *(Opcional)* Selecciona las filas de los correos que te interesan.
3. Pulsa **Guardar adjuntos** y responde:
   - **Sí** → solo las filas seleccionadas.
   - **No** → todos los correos de la lista.
4. Elige la carpeta de destino. Si un archivo ya existe, se guarda como `nombre (2).ext` (nunca se sobrescribe).

### 4.4 Conectar y desconectar

La sesión de Notes se recuerda mientras el libro está abierto para no pedir la
contraseña cada vez. Un mismo botón sirve para las dos cosas y **cambia de texto solo**:

- **Conectar a Notes** → abre la sesión (pide la contraseña si hace falta). Después el botón pasa a decir *Desconectar de Notes*.
- **Desconectar de Notes** → olvida la sesión (útil si cambias de usuario). Después el botón vuelve a decir *Conectar a Notes*.

No es obligatorio conectar antes: *Leer bandeja*, *Enviar correos* y *Guardar adjuntos* se conectan solos, y el botón también se actualiza.

> El botón siempre mira el estado **real** de la conexión. Si alguna vez el texto
> no coincide (por ejemplo, al abrir un libro que se guardó estando conectado),
> púlsalo igualmente: hará lo correcto y el texto se corregirá.

<details>
<summary>Opcional: que el botón diga «Conectar» nada más abrir el libro</summary>

1. **Alt + F11** y, a la izquierda, doble clic en **ThisWorkbook** (*EsteLibro*).
2. Pega este código y guarda:

```vb
Private Sub Workbook_Open()
    ActualizarBotonesConexion
End Sub
```
</details>

### 4.5 Actualizar desde una versión anterior

Si ya tenías el libro montado con el botón antiguo *Desconectar de Notes*:

1. **Alt + F11**. En el panel izquierdo, clic derecho en **modConfiguracion**, **modNotesConexion** y **modPreparar** > *Quitar…* > **No** (no hace falta exportarlos).
2. **Archivo > Importar archivo** y vuelve a importar esos tres `.bas`.
3. **Depuración > Compilar VBAProject**.
4. **Alt + F8 > PrepararLibro > Ejecutar**. Borra el botón antiguo y crea los nuevos. Tus datos no se tocan.

---

## 5. Configuración

Todo lo que puedes ajustar está en el módulo **`modConfiguracion`**:

| Constante | Para qué sirve | Valor por defecto |
|---|---|---|
| `HOJA_BANDEJA` / `HOJA_ENVIAR` | Nombres de las hojas | `"Bandeja"` / `"Enviar"` |
| `MAX_CORREOS` | Nº máximo de correos que se leen | `100` |
| `LONGITUD_EXTRACTO` | Caracteres del cuerpo que se copian | `200` |
| `SERVIDOR_CORREO` / `ARCHIVO_CORREO` | Ubicación del buzón si la detección automática falla | `""` (automático) |
| `TEXTO_CONECTAR` / `TEXTO_DESCONECTAR` | Textos del botón de conexión | `"Conectar a Notes"` / `"Desconectar de Notes"` |

Ejemplo si tu buzón no se detecta solo (pregúntale los datos a tu departamento de IT o
míralos en Notes: *Archivo > Preferencias > Ubicaciones > Editar ubicación > Correo*):

```vb
Public Const SERVIDOR_CORREO As String = "ServidorCorreo/MiEmpresa"
Public Const ARCHIVO_CORREO As String = "mail\jperez.nsf"
```

---

## 6. Solución de problemas frecuentes

### «No se pudo conectar con HCL Notes desde Excel» o error 429 («ActiveX no puede crear el objeto»)

Es el problema más habitual. Causas:

1. **Notes no está instalado** o nunca se ha configurado en este PC → instala/configura Notes y ábrelo una vez.
2. **Excel y Notes tienen distinta arquitectura** (por ejemplo, Excel 64 bits y Notes 32 bits).
   La API COM de Notes (`Lotus.NotesSession`) solo funciona si ambos coinciden.
   - El código intenta automáticamente un «plan B» (`Notes.NotesSession`), que abre el cliente Notes y a veces funciona aunque no coincidan.
   - Si tampoco funciona: instala la versión de Excel (Office) o de Notes con la misma arquitectura que la otra. Consulta con tu departamento de IT.
3. **La librería COM de Notes no está registrada.** Con permisos de administrador, abre una consola (`cmd`) y ejecuta (ajusta la ruta a tu instalación):
   ```
   regsvr32 "C:\Program Files (x86)\HCL\Notes\nlsxbe.dll"
   ```

#### Cómo saber si Excel y Notes son de 32 o 64 bits

- **Excel:** *Archivo > Cuenta > Acerca de Excel*. En la primera línea pone «32 bits» o «64 bits».
- **Notes:** mira la carpeta de instalación: `C:\Program Files (x86)\...` suele ser **32 bits**; `C:\Program Files\...` suele ser **64 bits**. También en Notes: *Ayuda > Acerca de HCL Notes*.
  (HCL Notes es de 32 bits hasta la versión 12.0.1; desde la 12.0.2 existe también en 64 bits.)

### Me pide la contraseña cada vez / la contraseña no funciona

- Revisa mayúsculas y minúsculas.
- Para no tener que escribirla en Excel: en Notes activa *Archivo > Seguridad > Seguridad del usuario > «No pedir contraseña a otros programas de Notes»* (el nombre exacto varía según la versión) y deja Notes abierto.
- ⚠️ La ventana de contraseña de Excel (`InputBox`) **no oculta** lo que escribes. Asegúrate de que nadie mira la pantalla.
- La ventana de contraseña de Notes puede aparecer **detrás de Excel**: revisa la barra de tareas.

### «No se pudo abrir tu base de correo»

- Comprueba que tienes conexión con el servidor (red / VPN) y que puedes abrir tu correo en Notes.
- Si trabajas con una réplica local o una configuración especial, indica `SERVIDOR_CORREO` y `ARCHIVO_CORREO` en `modConfiguracion` (ver apartado 5).

### «No se encontró la carpeta Bandeja de entrada»

Tu buzón usa una plantilla poco habitual. Abre `modLeerCorreo` y cambia `"($Inbox)"` por el nombre de la vista que quieras leer.

### Las macros no se ejecutan / aparece una barra amarilla de seguridad

- Pulsa **Habilitar contenido** en la barra amarilla.
- Si el archivo viene de Internet o del correo, Windows lo puede bloquear: clic derecho en el `.xlsm` > *Propiedades* > marca **Desbloquear** > *Aceptar*.
- Revisa *Archivo > Opciones > Centro de confianza > Configuración del Centro de confianza > Configuración de macros*.

### Las tildes y la ñ se ven raras en el editor de VBA

Has copiado/pegado el código desde la web o has importado un archivo guardado en UTF-8.
Importa los `.bas` **originales** del ZIP descargado (están en Windows-1252).

### Error «Compile error» / «Error de compilación» al compilar

- Asegúrate de haber importado **los 6 módulos** (unos usan funciones de otros).
- No tengas dos módulos con el mismo nombre (si importaste dos veces, borra el duplicado: clic derecho > *Quitar*).

### El cuerpo del correo sale vacío en la hoja

Algunos correos en formato HTML/MIME no exponen el texto de forma sencilla. El resto de columnas sí se rellenan.

### Va muy lento

Reduce `MAX_CORREOS` en `modConfiguracion`. Leer cientos de correos de un servidor remoto lleva su tiempo.

---

## 7. Cómo funciona (para curiosos)

- `CreateObject("Lotus.NotesSession")` crea la sesión COM de Notes y `Initialize` hace el login.
- `GetEnvironmentString("MailServer"/"MailFile")` lee de `notes.ini` dónde está tu buzón; `GetDatabase` lo abre.
- La Bandeja de entrada es la carpeta `($Inbox)`; se recorre con `GetLastDocument` / `GetPrevDocument` (de más nuevo a más antiguo).
- Cada correo es un documento con campos: `From`, `Subject`, `Body`, `DeliveredDate`…
- Para enviar: `CreateDocument`, se rellenan `Form = "Memo"`, `SendTo`, `CopyTo`, `Subject`, un campo `Body` de texto enriquecido con `EmbedObject` para los adjuntos, y `Send`.

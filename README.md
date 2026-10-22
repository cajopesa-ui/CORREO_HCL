# Correo HCL Notes desde Excel (VBA)

Macros de Excel que se conectan a **HCL Notes** (antes Lotus Notes) para:

- 📥 **Leer la Bandeja de entrada** y volcarla en una hoja: Fecha, Remitente, Asunto, extracto del cuerpo y nº de adjuntos.
- 📤 **Enviar correos** (Para, CC, Asunto, Cuerpo y adjuntos opcionales) a partir de las filas de una hoja.
- 📎 **Guardar los adjuntos** de los correos leídos en una carpeta de tu PC.
- 📁 **Gestionar carpetas**: ver el árbol de carpetas, crear carpetas y subcarpetas, renombrarlas y eliminarlas.
- ↪️ **Mover y copiar correos** a tus carpetas eligiendo el destino en un desplegable.
- 📂 **Leer carpetas completas** (con sus subcarpetas) y **buscar** en todo el buzón por texto, remitente y fechas.
- 🤖 **Reglas de archivado**: mueve automáticamente los correos de la Bandeja de entrada según el remitente, el asunto…
- ✉️ **Marcar como leído / no leído**.
- ❓ **Ayuda interactiva** dentro del libro: hoja *Ayuda* con índice, pasos y ejemplos, botón *? Ayuda de esta hoja*, notas en los encabezados y visita guiada.
- 🗂️ **Clasificar por temas**: guarda los correos en carpetas dentro de *Clasificacion nueva*, con la clasificación que prepara Claude o con palabras clave.

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
├── modLeerCorreo.bas      ← macro LeerBandejaEntrada (y el "motor" de lectura)
├── modEnviarCorreo.bas    ← macro EnviarCorreos (y la función EnviarUnCorreo)
├── modAdjuntos.bas        ← macros GuardarAdjuntos y GuardarAdjuntosCarpeta
├── modCarpetas.bas        ← árbol de carpetas: actualizar, crear, renombrar, eliminar
├── modMoverCorreo.bas     ← mover / copiar correos, marcar leído / no leído
├── modBuscar.bas          ← leer una carpeta completa y buscar correos
├── modReglas.bas          ← reglas de archivado automático
├── modClasificar.bas      ← clasificación por temas (Claude o palabras clave)
├── modAyuda.bas           ← ayuda interactiva (hoja Ayuda, notas, visita guiada)
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
5. Repite el paso 3–4 con **cada uno** de los 12 archivos `.bas` de la carpeta `vba`.
6. En el panel izquierdo (*Explorador de proyectos*) verás una carpeta **Módulos** con los 12 módulos.
7. *(Recomendado)* Menú **Depuración > Compilar VBAProject**. Si no sale ningún mensaje, todo está bien.

### 3.3 Guardar como libro con macros (.xlsm)

1. Vuelve a Excel (**Alt + F11** otra vez, o cierra el editor).
2. **Archivo > Guardar como**.
3. En *Tipo*, elige **Libro de Excel habilitado para macros (\*.xlsm)**.
4. Guarda. ⚠️ Si lo guardas como `.xlsx` **se perderán las macros**.

### 3.4 Preparar las hojas y los botones (automático)

1. Pulsa **Alt + F8**, elige **PrepararLibro** y pulsa **Ejecutar**.
2. Se crearán (el libro se abre en la hoja **Ayuda**):
   - **Ayuda** (la primera): índice con enlaces y, para cada tema, para qué sirve, los pasos y un ejemplo. Arriba hay un desplegable *«¿Qué quieres hacer?»* y los botones *Ver los pasos de este tema*, *Llévame a la hoja* y *Visita guiada por las hojas*.
   - **Bandeja**: correos leídos. Botones *Leer bandeja de entrada*, *Mover a carpeta*, *Copiar a carpeta*, *Marcar como leído*, *Marcar como no leído*, *Guardar adjuntos*, *Enviar correos* y *Conectar a Notes*.
   - **Enviar**: correos para enviar, con una fila de ejemplo. Botones *Enviar correos* y *Conectar a Notes*.
   - **Carpetas**: tu árbol de carpetas. Botones *Actualizar carpetas*, *Leer carpeta*, *Crear carpeta / subcarpeta*, *Renombrar carpeta*, *Eliminar carpeta*, *Crear carpetas de la lista*, *Guardar adjuntos de la carpeta* y *Conectar a Notes*.
   - **Buscar**: criterios de búsqueda. Botones *Buscar*, *Limpiar criterios* y *Conectar a Notes*.
   - **Reglas**: reglas de archivado, con una regla de ejemplo desactivada. Botones *Aplicar reglas a la Bandeja* y *Conectar a Notes*.
   - **Clasificacion**: revisión de la clasificación por temas. Botones *Clasificar con palabras clave*, *Importar clasificación de Claude*, *Aplicar clasificación* y *Conectar a Notes*.
   - **Palabras clave**: las reglas de la clasificación automática (ya rellenas con una propuesta).
   - **Listas** (oculta): las opciones de los desplegables. No la borres.
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
2. **¿Cuántos correos?** Escribe el número máximo (se empieza por los más recientes). Se propone `MAX_CORREOS` (100) y después el último número que escribiste.
3. **¿Contenido completo?**
   - **No** → *resumen*: solo la hoja, como siempre (extracto de 200 caracteres).
   - **Sí** → *completo*: la hoja como siempre **y además** un archivo `.txt` con el texto entero de cada correo (ver [4.10](#410-contenido-completo-en-un-txt)).
   - **Cancelar** → no se lee nada.
4. Si elegiste *completo*, se abre la ventana **Guardar como** de Windows para que elijas carpeta y nombre del `.txt`.
5. La primera vez se conectará con Notes:
   - Si Notes está abierto y tiene activada la opción de compartir la contraseña, entra directamente.
   - Si no, te pedirá la **contraseña de Notes** (en una ventana de Notes o de Excel).

| A Fecha | B Remitente | C Asunto | D Extracto | E Nº adjuntos | F ID (no tocar) | G Carpeta | H Leído | I Mover a | J Resultado |
|---|---|---|---|---|---|---|---|---|---|

> La columna **F** guarda un identificador interno del correo. La necesitan las macros de adjuntos, mover y marcar: no la borres.
> La columna **G** indica en qué carpeta está el correo (la usa *Mover* para saber de dónde sacarlo).

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

### 4.5 Carpetas (hoja «Carpetas»)

Pulsa **Actualizar carpetas**. La hoja muestra tu árbol de carpetas: las subcarpetas aparecen **con sangría** debajo de su carpeta.

| A Carpeta | B Ruta completa | C Nº correos | D No leídos | E Tipo |
|---|---|---|---|---|
| Bandeja de entrada | Bandeja de entrada | 120 | 5 | Sistema |
| **Clientes** | Clientes | 3 | 0 | Carpeta |
| &nbsp;&nbsp;&nbsp;&nbsp;Acme | Clientes\Acme | 40 | 2 | Carpeta |
| &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2026 | Clientes\Acme\2026 | 8 | 1 | Carpeta |

- En Notes las subcarpetas se escriben con `\`: `Clientes\Acme\2026` es la carpeta *2026* dentro de *Acme* dentro de *Clientes*.
- **Agrupación** = un «padre» que Notes muestra pero que no existe como carpeta (por ejemplo, existe `Clientes\Acme` pero no `Clientes`).
- *No leídos* puede salir vacío en versiones antiguas de Notes.

Para usar los botones, **primero haz clic en la fila de la carpeta**:

| Botón | Qué hace |
|---|---|
| **Leer carpeta** | Copia sus correos a la hoja *Bandeja*. Pregunta si incluir las **subcarpetas**, si quieres solo los **últimos N días**, **cuántos correos** como máximo (se propone 500) y si los quieres en **resumen o completos** (`.txt`). |
| **Crear carpeta / subcarpeta** | Crea una carpeta **dentro de la seleccionada**. Sin selección, la crea en el primer nivel. Puedes crear varios niveles de golpe: `Facturas\2026`. |
| **Renombrar carpeta** | Cambia el nombre. Sus subcarpetas la siguen. Los correos no se tocan. |
| **Eliminar carpeta** | Elimina la carpeta **y sus subcarpetas**. Antes te pregunta si quieres pasar sus correos a la Bandeja de entrada y te pide una segunda confirmación. **Los correos nunca se borran**: si no los pasas a la Bandeja, siguen en la vista *Todos los documentos* de Notes. |
| **Crear carpetas de la lista** | Escribe en la **columna G** una ruta por fila (`Clientes\Acme`, `Clientes\Acme\2026`…) y pulsa: las crea todas y escribe el resultado en la columna H. |
| **Guardar adjuntos de la carpeta** | Guarda en tu PC los adjuntos de todos los correos de la carpeta (y, si quieres, de sus subcarpetas) **recreando la misma estructura de carpetas**: `C:\Destino\Clientes\Acme\2026\factura.pdf`. |

> «Bandeja de entrada» y «Enviados» son carpetas del sistema: se pueden leer, pero no renombrar, eliminar ni tener subcarpetas.

### 4.6 Mover y copiar correos (hoja «Bandeja»)

1. Lee la bandeja, una carpeta o haz una búsqueda: los correos aparecen en la hoja *Bandeja*.
2. En la columna **I «Mover a»** (encabezado naranja) haz clic en la celda y **elige la carpeta en el desplegable**. Hazlo en todas las filas que quieras.
3. Pulsa:
   - **Mover a carpeta**: el correo pasa a la carpeta elegida y **sale** de la carpeta en la que estaba (columna G).
   - **Copiar a carpeta**: el correo se añade a la carpeta elegida y **sigue también** en la suya (en Notes un correo puede estar en varias carpetas).
4. La columna **J «Resultado»** dice qué ha pasado con cada fila (verde = bien, rojo = error).

> Si creas una carpeta nueva y no aparece en el desplegable, pulsa **Actualizar carpetas** en la hoja *Carpetas*.

### 4.7 Marcar como leído / no leído

Selecciona las filas en la hoja *Bandeja* (puedes usar **Ctrl + clic** o **Mayús + clic**) y pulsa **Marcar como leído** o **Marcar como no leído**. La columna **H «Leído»** se actualiza. Necesita HCL Notes 8 o superior.

### 4.8 Buscar correos (hoja «Buscar»)

Rellena solo los criterios que quieras (celdas amarillas) y pulsa **Buscar**. Igual que al leer la bandeja, te preguntará **cuántos resultados** quieres como máximo y si los quieres en **resumen o completos** (`.txt`).

| Criterio | Ejemplo | Nota |
|---|---|---|
| Texto (en asunto o cuerpo) | `presupuesto` | No distingue mayúsculas. |
| Remitente contiene | `acme.com` o `Juan` | Vale el nombre o la dirección. |
| Desde / Hasta (fecha) | `01/01/2026` | Las dos incluidas. |
| Carpeta | *(desplegable)* | **Vacía = busca en todas**: Bandeja de entrada, Enviados y todas tus carpetas. |
| Incluir subcarpetas | `Sí` / `No` | Solo cuenta si eliges una carpeta. |

Los resultados aparecen en la hoja *Bandeja*, ordenados del más reciente al más antiguo, y desde ahí puedes moverlos, copiarlos, marcarlos o guardar sus adjuntos. **Limpiar criterios** vacía las celdas.

> Buscar texto dentro del **cuerpo** de miles de correos puede tardar varios minutos. La barra de estado de Excel (abajo) muestra el progreso.

### 4.9 Reglas de archivado automático (hoja «Reglas»)

Cada fila es una regla:

| A Buscar en | B Contiene | C Mover a carpeta | D Activa | E Resultado |
|---|---|---|---|---|
| Remitente | `@acme.com` | Clientes\Acme | Sí | *(lo rellena la macro)* |
| Asunto | `factura` | Facturas\2026 | Sí | |
| Para/CC | `ventas@miempresa.com` | Ventas | No | |

- **Buscar en**: *Remitente*, *Para/CC*, *Asunto* o *Cuerpo* (desplegable).
- **Mover a carpeta**: desplegable con tus carpetas.
- **Activa = No** desactiva la regla sin borrarla.

Pulsa **Aplicar reglas a la Bandeja**:

1. Revisa **todos** los correos de la Bandeja de entrada (sin mover nada todavía).
2. Te muestra un **resumen** («Fila 2: 5 correos → Clientes\Acme…») y te pide confirmación.
3. Si aceptas, mueve cada correo a su carpeta y lo saca de la Bandeja de entrada.

Las reglas se revisan **de arriba abajo**: si un correo cumple varias, manda la primera.

### 4.10 Contenido completo en un .txt

Al leer la bandeja, una carpeta o al buscar, si respondes **Sí** a «¿contenido completo?» se crea un `.txt`
(en **UTF-8**, con tildes y ñ correctas) con esta estructura, pensada para que la lea una persona **o un programa / una IA**:

```
=== EXPORTACIÓN HCL NOTES ===
Fecha exportación: 2026-10-16 10:30
Usuario: Juan Pérez
Origen: Bandeja de entrada
Nº correos: 25
Límite alcanzado: No
Formato: cada correo va entre '### CORREO n DE N ###' y '### FIN CORREO n ###'.
         El cuerpo va entre '--- CUERPO ---' y '--- FIN CUERPO ---'.
         Fechas en formato AAAA-MM-DD hh:mm. Orden: del más reciente al más antiguo.
==============================

### CORREO 1 DE 25 ###
ID: 8A3F0C...C21
Fecha: 2026-10-15 09:12
De: Ana López <ana@acme.com>
Asunto: Presupuesto octubre
Carpeta: Bandeja de entrada
Leído: Sí
Adjuntos (2): presupuesto.pdf; logo.png
--- CUERPO ---
Hola Juan,

Te adjunto el presupuesto...
--- FIN CUERPO ---
### FIN CORREO 1 ###
```

- Los correos van en el **mismo orden que en la hoja** (más recientes primero).
- Las etiquetas (`De:`, `Asunto:`…) son siempre las mismas y cada una ocupa una sola línea.
- Los destinatarios (**Para** y **CC**) **no se exportan**, porque suelen ser listas muy largas. Tampoco los de los correos **reenviados o respondidos** que Notes copia dentro del cuerpo (las líneas `Para:`, `cc:`, `To:`… y su continuación).
- El cuerpo conserva sus saltos de línea. Si alguna línea del cuerpo empezase por `### ` o `--- `, se le añade un espacio delante para que no se confunda con las marcas.
- Los adjuntos **no** van dentro del `.txt` (solo sus nombres). Para guardarlos usa *Guardar adjuntos*.
- Los correos en formato web (HTML) se convierten a texto antes de exportarlos.

### 4.11 Clasificar por temas (hojas «Clasificacion» y «Palabras clave»)

Organiza los correos **por temas** en carpetas que cuelgan de **`Clasificacion nueva`**, para que no se mezclen con tus carpetas:

```
Clasificacion nueva
├── Informatica
│   ├── Despliegues CEAR
│   ├── Equipos NPT y material
│   ├── Cover Tamper HP
│   ├── Impresoras y consumibles
│   ├── Incidencias y Serman
│   ├── Redes y comunicaciones
│   ├── Zoom y telefonia
│   ├── Obsolescencia
│   └── Atencion al contribuyente
├── Edificios y seguridad
├── RRHH
│   └── Sindicatos
└── Personal
```

Hay **dos formas** de obtener la clasificación. Las dos terminan en la hoja **Clasificacion**, donde la **revisas antes de aplicarla**:

**A) La prepara Claude (más precisa)**
1. Lee la bandeja con **contenido completo** y guarda el `.txt` (apartado 4.10).
2. Pásale el `.txt` a Claude y pídele la clasificación. Te devolverá un archivo con una línea por correo: `ID;Carpeta;Asunto`.
3. Pulsa **Importar clasificación de Claude** y elige ese archivo.

**B) Palabras clave (automática, sin salir de tu PC)**
1. En la hoja **Palabras clave** (el botón **Cargar palabras clave propuestas** pone las 23 reglas que preparó Claude a partir de los correos de 2026, que aciertan en 612 de 624) cada fila es una regla: *carpeta*, *dónde buscar* (Asunto, Remitente, Cuerpo o Asunto o cuerpo), *palabras separadas por `;`* y *Activa*.
   Las reglas se revisan **de arriba abajo** y gana la **primera** que se cumple. No distinguen mayúsculas ni tildes.
2. Pulsa **Clasificar con palabras clave** y elige qué correos:
   - **Sí** → los que hay ahora en la hoja **Bandeja** (lo último que leíste o buscaste). Así puedes clasificar una búsqueda por fechas (por ejemplo, todo 2026), por remitente o una carpeta entera.
   - **No** → los últimos N de la **Bandeja de entrada** de Notes (te pregunta cuántos).
3. Los que no cumplan ninguna regla salen como **Sin clasificar** con *Aplicar = No* (se quedan donde están).

**Revisar y aplicar**
- En la hoja **Clasificacion** puedes cambiar la **carpeta propuesta** (columna D) o poner **Aplicar = No** (columna F) en los que no quieras tocar.
- Pulsa **Aplicar clasificación**: te dice cuántos correos y cuántas carpetas nuevas, pide confirmación, **crea las carpetas que falten** y **copia** cada correo a la suya. La columna *Resultado* dice qué ha pasado y *Aplicar* pasa a «Hecho».
- Por defecto **copia**: el correo sigue también en la Bandeja de entrada. Si un día prefieres **mover**, pon `CLASIFICACION_MUEVE = True` en `modConfiguracion`.
- Si no te gusta el resultado, borra la carpeta *Clasificacion nueva* desde la hoja *Carpetas* (sin pasar los correos a la Bandeja, porque siguen allí): no se pierde nada.

### 4.12 Ayuda interactiva

- **Hoja Ayuda**: elige en la celda amarilla qué quieres hacer y pulsa **Ver los pasos de este tema** (salta a su explicación con ejemplo) o **Llévame a la hoja** (te abre la hoja correcta y te recuerda los pasos en una ventana). También puedes hacer clic en cualquier tema del índice.
- **Botón «? Ayuda de esta hoja»** en todas las hojas: explica qué puedes hacer en la hoja en la que estás.
- **Notas en los encabezados**: pasa el ratón por encima de un encabezado (el triangulito rojo de la esquina) y verás qué poner y un ejemplo.
- **Visita guiada por las hojas**: recorre el libro hoja a hoja con una explicación de cada una.
- Todos los textos están en el módulo `modAyuda` (procedimiento `CargarTemas`). Si los cambias, vuelve a ejecutar **PrepararLibro**.

### 4.13 Actualizar desde una versión anterior

Si ya tenías el libro montado con una versión anterior:

1. **Alt + F11**. En el panel izquierdo, clic derecho en **cada módulo** que ya tengas (todos los que empiezan por `mod`) > *Quitar…* > **No** (no hace falta exportarlos).
2. **Archivo > Importar archivo** e importa los **12** `.bas` de la carpeta `vba`.
3. **Depuración > Compilar VBAProject**.
4. **Alt + F8 > PrepararLibro > Ejecutar**. Crea las hojas nuevas, mueve los botones de la hoja *Bandeja* a la columna L y borra el botón antiguo *Desconectar de Notes*. Tus datos no se tocan.

---

## 5. Configuración

Todo lo que puedes ajustar está en el módulo **`modConfiguracion`**:

| Constante | Para qué sirve | Valor por defecto |
|---|---|---|
| `HOJA_BANDEJA`, `HOJA_ENVIAR`, `HOJA_CARPETAS`, `HOJA_BUSCAR`, `HOJA_REGLAS` | Nombres de las hojas | `"Bandeja"`, `"Enviar"`, `"Carpetas"`, `"Buscar"`, `"Reglas"` |
| `MAX_CORREOS` | Nº de correos que se **propone** al leer la Bandeja de entrada | `100` |
| `MAX_CORREOS_CARPETA` | Nº de correos que se **propone** al leer una carpeta o buscar | `500` |
| `CARPETA_CLASIFICACION` | Carpeta de la que cuelga la clasificación por temas | `"Clasificacion nueva"` |
| `CLASIFICACION_MUEVE` | `False` = copiar, `True` = mover (sacar de la Bandeja) | `False` |
| `CARPETA_BANDEJA` / `CARPETA_ENVIADOS` | Cómo se ven en Excel las carpetas del sistema | `"Bandeja de entrada"` / `"Enviados"` |
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

Pide menos correos en la pregunta «¿Cuántos correos?». Leer cientos de correos de un servidor remoto lleva su tiempo.
Al buscar, usa el filtro de fechas o elige una carpeta concreta: buscar texto en el cuerpo de todo el buzón es lo más lento.

### La columna «Mover a» no tiene desplegable / no sale una carpeta nueva

Ve a la hoja *Carpetas* y pulsa **Actualizar carpetas**. El desplegable usa esa lista.
Después vuelve a leer la bandeja (o la carpeta) para que las filas tengan el desplegable.

### «No leídos» sale vacío o «Marcar como leído» da error

Esas funciones necesitan **HCL Notes 8 o superior**. En versiones antiguas el resto de macros funciona igual.

### He eliminado una carpeta y no encuentro sus correos

Si al eliminarla elegiste **No** (no pasarlos a la Bandeja), los correos siguen en tu buzón, en la vista
**Todos los documentos** de Notes. La hoja *Buscar* no los encuentra (solo busca dentro de carpetas):
ábrelos en Notes desde *Todos los documentos* y muévelos a una carpeta.

---

## 7. Cómo funciona (para curiosos)

- `CreateObject("Lotus.NotesSession")` crea la sesión COM de Notes y `Initialize` hace el login.
- `GetEnvironmentString("MailServer"/"MailFile")` lee de `notes.ini` dónde está tu buzón; `GetDatabase` lo abre.
- La Bandeja de entrada es la carpeta `($Inbox)`; se recorre con `GetLastDocument` / `GetPrevDocument` (de más nuevo a más antiguo).
- Cada correo es un documento con campos: `From`, `Subject`, `Body`, `DeliveredDate`…
- Para enviar: `CreateDocument`, se rellenan `Form = "Memo"`, `SendTo`, `CopyTo`, `Subject`, un campo `Body` de texto enriquecido con `EmbedObject` para los adjuntos, y `Send`.

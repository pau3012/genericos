# Prompt: migrar un script de Tourplan NX a multi-entorno + cola de trabajo en Google Sheets

Uso: pegar este prompt (completando la sección "Datos del script a migrar") al
pedirle a Claude que aplique la misma migración a otro script de Tourplan NX
(Selenium). Documenta los cambios estructurales reales que hizo falta hacer
para llevar `relevamiento_vigencias_multientorno.py` (repo `genericos`) de "un
solo Tourplan + Excel local" a "cualquier Tourplan + Google Sheets resumible",
incluyendo los problemas concretos con los que nos topamos y cómo se
diagnosticaron — no es teoría, es lo que realmente pasó.

No incluye cambios de lógica de negocio específicos del script de vigencias
(por ejemplo, exportar todas las filas de Price Code de un mismo período);
esto es sólo la parte de infraestructura, pensada para ser genérica.

## Datos del script a migrar (completar antes de usar este prompt)

- Path del script: ___
- ¿Ya tiene una cola de trabajo en Excel (openpyxl) con columnas
  ESTADO/OBSERVACIONES? ___
- URL(s) de los entornos de Tourplan contra los que debe poder correr: ___
- ¿Ya existe un Google Sheet real para usar como cola (con sus propias
  columnas de un reporte de TP), o hay que crear uno nuevo?: ___
- Nombres de las pestañas de entrada/salida en ese Sheet: ___

---

## Cambio 1 — Config multi-entorno: agrupar todo lo que varía por instalación de Tourplan

**Por qué:** el script original tenía `BASE_URL`, `USERNAME`, `PASSWORD` y
`STYPE_SIDEBAR` como constantes sueltas, escritas para una sola instalación de
Tourplan. Cada Tourplan es una instalación independiente: puede tener otra URL,
otras credenciales y, sobre todo, **otro catálogo de Service Types** (otras
siglas, y otro orden en el sidebar de Product Search). Si se dejan sueltas,
cambiar de entorno implica tocar variables en cuatro lugares distintos del
archivo y es fácil dejar una desactualizada.

**Qué hacer:** juntar todo lo que es específico de una instalación en un único
diccionario, y elegir la instalación activa con una sola variable:

```python
ENTORNOS = {
    "TEST": {
        "BASE_URL": "https://.../TourplanNX_Test/",   # ver Cambio 2 para el "/" final
        "USERNAME": "...",
        "PASSWORD": "...",
        "SERVICE_TYPES": {
            "HT": ("01", ""),   # (posición en el sidebar o None, nombre o "")
            ...
        },
    },
    "NUEVO": {
        "BASE_URL": "https://.../index.html",
        "USERNAME": "...",
        "PASSWORD": "...",
        "SERVICE_TYPES": {
            "AC": (None, "Accommodation"),   # None = sin confirmar (ver Cambio 3)
            ...
        },
    },
}

ENTORNO_ACTUAL = "NUEVO"  # única variable a tocar para cambiar de Tourplan

_cfg = ENTORNOS[ENTORNO_ACTUAL]
BASE_URL = _cfg["BASE_URL"]
USERNAME = _cfg["USERNAME"]
PASSWORD = _cfg["PASSWORD"]
SERVICE_TYPES_CATALOG = {cod: nombre for cod, (_num, nombre) in _cfg["SERVICE_TYPES"].items()}
STYPE_SIDEBAR = {cod: num for cod, (num, _nombre) in _cfg["SERVICE_TYPES"].items() if num}

if BASE_URL.startswith("PONER_"):
    raise ValueError(f"Falta completar BASE_URL del entorno {ENTORNO_ACTUAL!r}.")
```

El resto del script sigue leyendo `BASE_URL`/`USERNAME`/`PASSWORD`/
`STYPE_SIDEBAR` como constantes globales — no hace falta tocar login,
búsqueda de producto, etc.

Si el script guarda algo en disco local (screenshots, un log, un Excel viejo
que se está reemplazando por Sheets — ver Cambio 5), conviene que ese nombre
de archivo incluya `ENTORNO_ACTUAL` para no mezclar datos de dos Tourplans
distintos en el mismo archivo.

---

## Cambio 2 — La trampa del `BASE_URL`: no todas las instalaciones sirven la SPA igual

**Síntoma real que tuvimos:** con la config del Cambio 1 ya andando, el script
se colgaba siempre en el mismo punto (`TimeoutException` esperando el botón de
búsqueda de Product Search), para códigos distintos, en un entorno nuevo. No
era un problema de datos ni de selectores del sidebar — la app de Tourplan
**nunca llegaba a cargar**.

**Causa real:** el script arma cada URL como `f"{BASE_URL}/#/ruta"`. Eso
funciona sólo si el servidor de esa instalación resuelve la ruta "pelada"
(algunas tienen un rewrite que sirve `index.html` para cualquier path). Otras
instalaciones exigen `index.html` explícito en la URL
(`.../TourplanNX/index.html#/home`) y con la barra de más (`index.html/#/home`,
con un `/` entre medio) tampoco cargan.

**Cómo lo confirmamos:** no lo dedujimos de los selectores — le pedimos al
usuario una grabación real de Chrome DevTools Recorder haciendo el flujo a
mano (login → Products → abrir Product Search) en el entorno nuevo, y ahí
apareció la URL real con `index.html#/home` explícito. **Si un script nuevo
se cuelga siempre en el mismo paso apenas cambia de entorno, pedir esa
grabación (o simplemente navegar a mano y mirar la barra de direcciones) antes
de tocar nada de Selenium** — evita diagnosticar a ciegas selectores que en
realidad están bien.

**Qué hacer:** no concatenar la barra en el código; que la incluya (o no)
`BASE_URL` mismo, y armar la URL siempre igual:

```python
driver.get(f"{BASE_URL}#/login")     # en TODOS los driver.get() de rutas hash
```

Y en `ENTORNOS`, cada entorno completa `BASE_URL` con el sufijo que
efectivamente necesita:

```python
"BASE_URL": "https://tourplannx.eurotur.com.ar/TourplanNX_Test/",     # barra al final
"BASE_URL": "https://la-perwel.nx.tourplan.net/TourplanNX/index.html", # index.html, sin barra
```

Dejar un comentario explícito en el diccionario `ENTORNOS` explicando esta
convención — si no, la próxima persona que agregue un entorno va a volver a
pisar el mismo bug.

---

## Cambio 3 — Catálogo de Service Types por entorno, con fallback sin número de sidebar

**Por qué:** el script original usaba `STYPE_SIDEBAR` (sigla → posición
numérica en el sidebar de Product Search) para hacer el click más robusto que
buscar por texto. Esas posiciones son específicas de cada instalación y hay
que confirmarlas por inspección real — no se pueden adivinar. Un entorno nuevo
puede tener siglas totalmente distintas (en este caso, `AC/BT/CR/DE/EN/EX/
FB/FE/FT/GU/OC/PK/TF/TR/TT`, sin relación con las del entorno viejo).

**Qué hacer:** cargar el catálogo de siglas nuevas con posición `None` cuando
todavía no está confirmada, y confiar en el fallback por texto que ya suele
tener `_completar_filtros_busqueda` (busca el `<li>` del sidebar cuyo texto
matchea la sigla como palabra completa). No hace falta escribir código nuevo
para esto si el script ya tenía ese fallback — sólo asegurarse de que
`STYPE_SIDEBAR` quede vacío para esa sigla en vez de forzar un número
inventado. Si más adelante alguien confirma las posiciones reales, se
completan en el mismo diccionario de `ENTORNOS` sin tocar el resto del código.

---

## Cambio 4 — Diagnóstico real de errores "genéricos" de Selenium

**Síntoma real que tuvimos:** un `except Exception as e: ...(f"error: {e}")`
guardaba mensajes vacíos o sólo direcciones de memoria del binario de
chromedriver (`Message: \nStacktrace:\n#0 0x... <unknown>`), completamente
inútiles para diagnosticar. El traceback completo, en cambio, sí mostraba la
línea exacta (terminó siendo el `wait()` del Cambio 2).

**Qué hacer:** en cualquier except "catch-all" de una fila/código individual,
no confiar en `str(e)`. Guardar el traceback completo a un archivo, más
screenshot y HTML de la página en ese momento:

```python
except Exception as e:
    log_path = f"{SS_DIR}/error_{cod[:10]}_{int(time.time())}.txt"
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(traceback.format_exc())
    fallidos.append(f"{cod}: {type(e).__name__} — ver {log_path}")
    ss(driver, nombre_err)     # screenshot
    dump(driver, nombre_err)   # HTML de la página
```

Esto no cambia el comportamiento resumible del batch (la fila igual se marca
ERROR y se sigue con la próxima) — sólo hace que la próxima vez que algo falle
raro, haya con qué diagnosticar sin pedirle otra vuelta de "pegame el
traceback" al usuario.

---

## Cambio 5 — Cola de trabajo: de Excel local a Google Sheets (resumible de verdad)

**Por qué:** una cola en `.xlsx` local vive en el disco del runtime de Colab.
Si Colab se desconecta o se reinicia, se pierde (o hay que acordarse de
descargarlo/subirlo a Drive a mano). Un Google Sheet vive afuera del runtime:
un corte no pierde nada, y la próxima corrida retoma leyendo el mismo Sheet.

**Punto de partida real en esta migración:** no arrancamos de una hoja en
blanco — el usuario ya tenía un reporte real exportado de Tourplan (columnas
`Loc, Serv, Supplier, SupplierName, Code, Description, Comment, Used,
Deleted`) que quería usar como entrada. Esto importa: **no hay que renombrar
ni reordenar las columnas que ya trae el reporte** — sólo mapear las que el
script necesita y agregar, al final, las que falten para que funcione como
cola de trabajo.

### 5.1 — Auth (Colab interactivo, sin manejar credenciales)

```python
from google.colab import auth as _colab_auth
from google.auth import default as _google_auth_default
import gspread

def conectar_sheets():
    _colab_auth.authenticate_user()          # pide permiso una vez por sesión
    creds, _ = _google_auth_default()
    gc = gspread.authorize(creds)
    sh = gc.open_by_url(GOOGLE_SHEET_URL)

    def _hoja(nombre):
        for ws in sh.worksheets():
            if ws.title.strip().lower() == nombre.strip().lower():
                return ws
        raise ValueError(f"No encontré la pestaña {nombre!r}. "
                          f"Disponibles: {[w.title for w in sh.worksheets()]}")

    return _hoja(HOJA_PRODUCTOS), _hoja(HOJA_VIGENCIAS)
```

(Si en vez de Colab se corre desde la app local, este bloque es el que cambia
por autenticación con Service Account — ver sección "Si se mueve a la app
local" más abajo.)

### 5.2 — Mapear las columnas del reporte real, sin tocarlas

```python
COL_LOCATION     = "Loc"
COL_SUPPLIER     = "Supplier"
COL_SERVICE_TYPE = "Serv"
COL_CODIGO       = "Code"

COLS_A_AGREGAR = ["RATE FROM", "RATE TO", "ESTADO", "OBSERVACIONES", "TIMESTAMP"]

def asegurar_columnas_productos(ws):
    headers = ws.row_values(1)
    faltantes = [c for c in COLS_A_AGREGAR if c not in headers]
    if faltantes:
        nuevos_headers = headers + faltantes          # SIEMPRE al final
        ws.update(range_name="A1", values=[nuevos_headers])
        headers = nuevos_headers
    return {h: i + 1 for i, h in enumerate(headers) if h}
```

Y al leer cada fila, traducir los nombres reales a los nombres internos que ya
usa el resto del script (así el resto de la lógica de negocio no se entera de
que las columnas se llaman distinto):

```python
d = dict(zip(headers, fila))
row = {
    "LOCATION": d.get(COL_LOCATION, ""),
    "SUPPLIER": d.get(COL_SUPPLIER, ""),
    "SERVICE TYPE": d.get(COL_SERVICE_TYPE, ""),
    "CODIGO": d.get(COL_CODIGO, ""),
    "RATE FROM": d.get("RATE FROM", ""),
    "RATE TO": d.get("RATE TO", ""),
    "ESTADO": d.get("ESTADO", ""),
}
```

### 5.3 — ESTADO nace vacío, nunca "PENDIENTE" automático

Cuando la columna ESTADO se agrega por primera vez sobre un reporte real (que
puede tener cientos de filas), **no** marcarlas todas PENDIENTE solas. Eso
dispararía sin querer una corrida gigantesca contra Tourplan (que tiene
licencias concurrentes limitadas) apenas alguien corre el script por primera
vez. Dejarla vacía y que el usuario marque a mano qué filas quiere correr en
cada tanda — es el mismo rol de seguridad que cumplía la fila "EJEMPLO" en la
versión Excel.

### 5.4 — Guardado fila a fila, con las llamadas en el orden correcto

Mismo patrón que la cola en Excel (guardar apenas termina cada fila, para que
un corte a mitad de camino dependa lo menos posible), pero con dos llamadas a
la API de Sheets en vez de un `wb.save()`:

```python
for row in pendientes:
    estado, observaciones, filas_vigencias = procesar_fila_producto(driver, row)
    agregar_filas_vigencias(ws_salida, filas_vigencias)      # 1º: el resultado (append-only)
    actualizar_fila_producto(ws_productos, col_idx, row_idx, estado, observaciones)  # 2º: recién ahí, ESTADO
```

Ese orden importa: si el corte pasa entre las dos llamadas, la fila queda en
PENDIENTE y se reprocesa — lo cual puede duplicar el resultado ya escrito en
la hoja de salida (ventana angosta, pero real; documentarla en el propio
script para que quien lo mantenga sepa que existe).

Actualizar ESTADO/OBSERVACIONES/TIMESTAMP en un solo request (batch), no en
tres llamadas sueltas — importa para no pegarle de más a la cuota de la API:

```python
def actualizar_fila_producto(ws_productos, col_idx, row_idx, estado, observaciones):
    ts = datetime.now().isoformat(timespec="seconds")
    updates = [{"range": rowcol_to_a1(row_idx, col_idx[campo]), "values": [[valor]]}
               for campo, valor in (("ESTADO", estado), ("OBSERVACIONES", observaciones), ("TIMESTAMP", ts))
               if campo in col_idx]
    ws_productos.batch_update(updates)
```

### 5.5 — Gotcha de versión de `gspread`: usar siempre kwargs en `update()`

La firma de `Worksheet.update()` cambió el orden de sus parámetros posicionales
entre versiones de `gspread` (v5 vs v6: `range_name, values` pasó a
`values, range_name`). Para que el script no se rompa según qué versión venga
preinstalada en el runtime de Colab, llamarlo siempre con keywords:

```python
ws.update(range_name="A1", values=[headers])   # NO ws.update("A1", [headers])
```

`append_rows()` no tuvo ese cambio de firma, pero no cuesta nada ser
consistente.

### 5.6 — Hoja de salida (append-only)

Si la pestaña de salida puede estar vacía la primera vez, crearle el header:

```python
def asegurar_columnas_vigencias(ws):
    if not ws.row_values(1):
        ws.update(range_name="A1", values=[HEADERS_SALIDA])
```

Y agregar todo lo nuevo con un solo `append_rows()` por fila de entrada
procesada (no una llamada por línea de resultado):

```python
ws_salida.append_rows([[fila.get(h, "") for h in HEADERS_SALIDA] for fila in filas],
                       value_input_option="USER_ENTERED")
```

---

## Checklist para aplicar esto a otro script

1. ¿El script tiene `BASE_URL`/`USERNAME`/`PASSWORD`/`STYPE_SIDEBAR` sueltos?
   → Cambio 1 (agruparlos en `ENTORNOS` + `ENTORNO_ACTUAL`).
2. ¿Se va a correr contra una instalación de Tourplan distinta a la original?
   → Cambio 2. Antes de tocar selectores, confirmar la URL real navegando a
   mano (o con una grabación de DevTools Recorder) si se cuelga siempre en el
   mismo paso apenas cambia de entorno.
3. ¿Esa instalación nueva tiene service types con otras siglas? → Cambio 3.
   No inventar números de sidebar — dejarlos en `None` y confiar en el
   fallback por texto hasta confirmarlos por inspección real.
4. ¿Los `except Exception` de códigos/filas individuales sólo guardan
   `str(e)`? → Cambio 4, especialmente antes de una corrida real contra un
   entorno nuevo (es justamente cuando más aparecen errores no previstos).
5. ¿La cola de trabajo es un `.xlsx` local y se quiere resumible de verdad
   (sobrevivir a un corte de Colab)? → Cambio 5. Si ya existe un reporte real
   para usar como entrada, mapear sus columnas (5.2) en vez de pedir que se
   arme un Sheet desde cero, y dejar ESTADO vacío por default (5.3).
6. Compilar (`python -m py_compile`) después de cada cambio estructural antes
   de pasar al siguiente — son cambios que se detectan al toque con un
   `SyntaxError`/`NameError`, no hace falta esperar a correrlo contra
   Tourplan para pescarlos.

## Si en vez de Colab esto se mueve a la app local ya desarrollada

Todo lo de los Cambios 1 a 4 es independiente del runtime (Colab vs. app
local) — no cambia nada. Lo único que cambia es el Cambio 5.1 (auth): en vez
de `google.colab.auth`, la app local necesita una **Service Account** de
Google Cloud (JSON key) para no depender de que alguien haga click en un
popup interactivo cada corrida:

```python
import gspread
gc = gspread.service_account(filename="ruta/a/la/key.json")
sh = gc.open_by_url(GOOGLE_SHEET_URL)
```

Con la Service Account creada, hay que compartir el Google Sheet con el email
de esa cuenta de servicio (algo como
`nombre@proyecto.iam.gserviceaccount.com`) dándole permiso de Editor — si no,
`open_by_url` falla con un error de permisos. El resto de las funciones de la
sección 5 (`asegurar_columnas_productos`, `cargar_pendientes`,
`actualizar_fila_producto`, `agregar_filas_vigencias`) no cambian nada.

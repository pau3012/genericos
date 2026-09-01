# ============================================================
# RELEVAMIENTO DE VIGENCIAS — lista de períodos de RATES, SIN entrar
# a ninguno. NO es extracción de tarifas: no lee ningún valor de costo,
# sólo la vigencia/estado de cada período (fechas, Price Code, moneda,
# estado, nombre).
#
# VERSIÓN MULTI-ENTORNO: a diferencia del script original (pensado para
# un solo Tourplan), acá URL, credenciales y catálogo de Service Types
# viven agrupados por ENTORNO en el diccionario ENTORNOS de más abajo.
# Cada Tourplan es una instalación distinta y puede tener su propio
# conjunto de Service Types (siglas y sidebar de Product Search
# distintos) — por eso van atados al mismo entorno y no sueltos. Para
# correr contra otro Tourplan: agregar (o completar) una entrada en
# ENTORNOS y apuntar ENTORNO_ACTUAL a esa clave. El resto del script
# (login, búsqueda de producto, lectura de RATES, cola en Sheets) es
# genérico y no cambia entre entornos.
# Google Colab — celda única (mismo formato que los scripts hermanos)
# ------------------------------------------------------------
#   Objetivo: dado un supplier/location/service type (o un código
#   puntual), buscar sus product codes en Tourplan, entrar al módulo
#   RATES de cada uno y exportar las columnas que se ven en la LISTA
#   de períodos (Rate Period, PC, Buy/Sell Currency, Sale Period, Rate
#   Status, Rate Text, Rate Name) — sin abrir ningún período. Sirve
#   para relevar rápido en qué estado está la carga de vigencias de un
#   supplier y detectar cuáles necesitan actualización, sin el costo
#   de leer la grilla de costos por rango de pax de cada período.
#
#   Adaptado de relevamiento_vigencias.py (a su vez de
#   extraccion_tarifas_vigentes.py, generico-vs-especificos): chrome
#   bootstrap, helpers base (ss/jc/set_val/wait/crear_driver), login,
#   hamburger/menu_item, _completar_filtros_busqueda, buscar_producto,
#   _scrapear_pagina_resultados/_hacer_scroll_resultados/
#   listar_codigos_supplier, parsear_fecha/_parse_rate_period/
#   _periodos_en_rango, logout(). Lo único que cambia respecto de esa
#   versión es la config (multi-entorno) descripta arriba — la lógica
#   de scraping es idéntica.
#
#   NO reutiliza _leer_tabla_rates/_extraer_filas_ad/_seleccionar_price_code
#   del script hermano: esos leen la grilla de costos DENTRO de un
#   período ya abierto, que acá no hace falta abrir. La lectura nueva
#   (_leer_vigencias) usa selectores de clase confirmados por
#   inspección real de la grilla de LISTA: td.tpcol-rateperiod,
#   td.tpcol-pricecodecode, td.tpcol-currencycode (aparece 2 veces por
#   fila — Buy y Sell Currency, en ese orden, sin distinción por clase),
#   td.tpcol-saleperiod, td.tpcol-ratestatuses, td.tpcol-ratetexts
#   (suele venir vacía), td.tpcol-ratenames.
#
#   SELECCIÓN DE PERÍODO — sin ordenar nada: Tourplan ya devuelve la
#   lista de períodos del más reciente al más viejo. Por fila de la
#   cola de entrada:
#     - RATE FROM/RATE TO vacíos → se exporta el primer período de la
#       lista tal cual la entrega Tourplan (el "último"/vigencia más
#       nueva) — y TODAS las filas de ese mismo período (un período
#       puede tener más de una fila, una por Price Code, con el mismo
#       Rate Period).
#     - RATE FROM/RATE TO completos → se exportan TODAS las filas de
#       TODOS los períodos que se solapen con ese rango (puede ser más
#       de uno, cada uno con sus filas por Price Code).
#
#   COLA DE TRABAJO EN GOOGLE SHEETS (gspread, ESTADO/OBSERVACIONES,
#   guardado fila a fila, resumible — misma idea que la skill
#   armando-excel-como-cola-de-trabajo pero contra un Sheet en vez de un
#   .xlsx local, para que la entrada/salida sobrevivan a un corte de
#   Colab). GOOGLE_SHEET_URL apunta a un Sheet real del usuario, con dos
#   pestañas:
#     - "productos" (entrada, ya viene de un reporte de TP con columnas
#       propias: Loc, Serv, Supplier, SupplierName, Code, Description,
#       Comment, Used, Deleted — ver COL_LOCATION/COL_SUPPLIER/
#       COL_SERVICE_TYPE/COL_CODIGO más abajo para el mapeo). El script
#       le agrega, la primera vez que corre, las columnas que le faltan
#       para funcionar como cola de trabajo: RATE FROM, RATE TO, ESTADO,
#       OBSERVACIONES, TIMESTAMP (ver COLS_A_AGREGAR) — sin tocar ni
#       reordenar las columnas que ya tenía el reporte. ESTADO nace
#       vacío en todas las filas: no se dispara nada solo, hay que
#       escribir PENDIENTE a mano en las filas que se quieran correr en
#       cada tanda. Ninguno de LOCATION/SUPPLIER/SERVICE TYPE/CODIGO es
#       obligatorio por sí solo — un campo vacío significa "todos" en
#       esa dimensión (ej. CODIGO vacío = releva TODOS los códigos que
#       matcheen los demás campos, vía listar_codigos_supplier) — pero
#       deben venir completos AL MENOS 2 de estos 4, en cualquier
#       combinación (ej. SERVICE TYPE+LOCATION o SUPPLIER+CODIGO), para
#       no disparar una búsqueda sin acotar. El valor de SERVICE TYPE
#       debe ser una sigla del catálogo del ENTORNO_ACTUAL (ver ENTORNOS
#       más abajo) — el script imprime el catálogo vigente al arrancar.
#     - "vigencias" (salida, append-only): una fila por período
#       exportado — headers creados solos si la pestaña está vacía.
#   Autenticación: auth interactivo de Colab (google.colab.auth) contra
#   la cuenta de Google de quien corre la celda — necesita permiso de
#   edición sobre el Sheet, no requiere manejar ningún archivo de
#   credenciales.
# ============================================================

import os, sys, subprocess, importlib, shutil, time, re, traceback
from datetime import datetime

print("🔧 Verificando entorno...\n")

# ── Paquetes Python ──────────────────────────────────────────
_PIPS_NEEDED = {
    "selenium":          "selenium",
    "webdriver_manager": "webdriver-manager",
    "gspread":           "gspread",
    "google.auth":       "google-auth",
}
_pips_faltantes = [pkg for mod, pkg in _PIPS_NEEDED.items()
                   if importlib.util.find_spec(mod) is None]
if _pips_faltantes:
    subprocess.run([sys.executable, "-m", "pip", "install", "-q"] + _pips_faltantes, check=True)

# ── Google Chrome (REUTILIZADO tal cual de extraccion_tarifas_vigentes.py
# — Colab NO trae Chrome preinstalado, hay que detectarlo/instalarlo cada
# corrida porque Colab resetea el entorno cuando pierde actividad) ──────

def _chrome_version(path):
    try:
        r = subprocess.run([path, "--version"], capture_output=True, text=True, timeout=8)
        v = (r.stdout or r.stderr).strip()
        return v if v and any(c.isdigit() for c in v) else ""
    except Exception:
        return ""


def _find_chrome():
    for name in ["google-chrome-stable", "google-chrome", "chromium-browser", "chromium"]:
        p = shutil.which(name)
        if p:
            v = _chrome_version(p)
            if v:
                return p, v
    for path in ["/usr/bin/google-chrome-stable", "/usr/bin/google-chrome"]:
        if os.path.exists(path):
            v = _chrome_version(path)
            if v:
                return path, v
    return None, ""


CHROMIUM_BIN, ver_chrome = _find_chrome()


def _instalar_chrome():
    """Instala google-chrome-stable en Colab. Estrategia 1: repo oficial
    de Google. Estrategia 2: descarga directa del .deb (fallback)."""
    DEVNULL = subprocess.DEVNULL

    def _apt_update():
        subprocess.run(["apt-get", "update", "-qq"], stdout=DEVNULL, stderr=DEVNULL)

    e1 = None
    try:
        print("  ⏳ Agregando repo Google Chrome y ejecutando apt...")
        subprocess.run(
            "wget -qO- https://dl.google.com/linux/linux_signing_key.pub "
            "| gpg --dearmor -o /usr/share/keyrings/google-chrome.gpg",
            shell=True, check=True, stdout=DEVNULL, stderr=DEVNULL)
        with open("/etc/apt/sources.list.d/google-chrome.list", "w") as f:
            f.write("deb [arch=amd64 signed-by=/usr/share/keyrings/google-chrome.gpg] "
                    "https://dl.google.com/linux/chrome/deb/ stable main\n")
        _apt_update()
        subprocess.run(["apt-get", "install", "-y", "-qq", "google-chrome-stable"],
                       check=True, stdout=DEVNULL, stderr=DEVNULL)
        if _find_chrome()[0]:
            return
    except Exception as e:
        e1 = e
        print(f"  ⚠ Repo Google falló ({e1}), probando descarga directa...")

    try:
        deb = "/tmp/google-chrome-stable.deb"
        url = "https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb"
        print(f"  ⏳ Descargando {url} ...")
        subprocess.run(["wget", "-q", "-O", deb, url], check=True)
        print("  ⏳ Instalando .deb ...")
        subprocess.run(["dpkg", "-i", deb], stdout=DEVNULL, stderr=DEVNULL)
        _apt_update()
        subprocess.run(["apt-get", "install", "-f", "-y", "-qq"],
                       check=True, stdout=DEVNULL, stderr=DEVNULL)
        print("  ✅ Chrome instalado via .deb")
    except Exception as e2:
        raise EnvironmentError(
            f"No se pudo instalar google-chrome-stable.\n"
            f"  Repo: {e1 if e1 else 'n/a'}\n"
            f"  .deb:  {e2}\n"
            "Intentá manualmente en Colab:\n"
            "  !wget -q https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb\n"
            "  !dpkg -i google-chrome-stable_current_amd64.deb\n"
            "  !apt-get install -f -y")


if not CHROMIUM_BIN:
    print("  ⏳ Chrome no encontrado, instalando...")
    _instalar_chrome()
    CHROMIUM_BIN, ver_chrome = _find_chrome()
    if CHROMIUM_BIN:
        print(f"  ✅ Chrome instalado: {ver_chrome}")
    else:
        raise EnvironmentError(
            "Chrome instalado pero no encontrado. "
            "Reiniciá el runtime de Colab y volvé a correr la celda.")
else:
    print(f"  ✅ Chrome OK: {CHROMIUM_BIN}  [{ver_chrome}]")

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from gspread.utils import rowcol_to_a1
from google.colab import auth as _colab_auth
from google.auth import default as _google_auth_default

# ── Config multi-entorno ─────────────────────────────────────────
# Cada entorno es una instalación de Tourplan distinta: URL, usuario/
# clave y catálogo de Service Types (sigla → nombre, y opcionalmente
# su posición numérica en el sidebar de Product Search) van juntos acá
# porque no se pueden mezclar entre entornos.
#
# BASE_URL: el script arma cada URL como f"{BASE_URL}#/login" (sin
# agregar ninguna "/" en el medio) — así que BASE_URL tiene que
# terminar en algo que, seguido de "#/...", dé una URL válida. Esto
# varía según cómo esté servida la app en cada instalación: algunas
# resuelven la ruta "pelada" (ahí BASE_URL termina en "/", ej. TEST más
# abajo); otras exigen "index.html" explícito en la URL (ahí BASE_URL
# termina en "/index.html", sin barra final, ej. NUEVO). Si al cambiar
# de entorno la app queda en blanco / nunca aparece el resto de la UI
# (timeout esperando cualquier selector), lo primero a revisar es esto
# — confirmar navegando a mano en el browser qué URL carga la app.
#
# STYPE_SIDEBAR (compatibilidad con _completar_filtros_busqueda): si un
# service type tiene posición numérica confirmada por inspección real
# del sidebar de ESE entorno, se arma automáticamente más abajo a
# partir de "num". Si "num" es None, el service type igual funciona:
# _completar_filtros_busqueda ya tiene un fallback que lo busca por
# texto (más lento/frágil, pero funcional) — así queda el entorno
# "NUEVO" hasta que se confirmen los números.
ENTORNOS = {
    # Entorno original (STYPE_SIDEBAR confirmado por inspección real —
    # ver script hermano extraccion_tarifas_vigentes.py).
    "TEST": {
        "BASE_URL": "https://tourplannx.eurotur.com.ar/TourplanNX_Test/",
        "USERNAME": "poner minusculas",
        "PASSWORD": "password",
        "SERVICE_TYPES": {
            "HT": ("01", ""), "HX": ("02", ""), "TF": ("03", ""),
            "EX": ("04", ""), "ML": ("05", ""), "RT": ("06", ""),
            "CR": ("07", ""), "FT": ("08", ""), "OC": ("09", ""),
            "LN": ("10", ""), "LP": ("11", ""), "MS": ("12", ""),
            "TA": ("13", ""), "PR": ("14", ""),
        },
    },
    # Entorno nuevo — completar BASE_URL/USERNAME/PASSWORD. Catálogo de
    # Service Types tal cual la tabla CODE/NAME provista para este
    # Tourplan; posición de sidebar aún no confirmada (None = fallback
    # por texto). Si más adelante se confirman los números, completarlos
    # acá mismo (ej. "AC": ("01", "Accommodation")) para que la
    # selección en el sidebar sea más robusta.
    "NUEVO": {
        "BASE_URL": "https://la-perwel.nx.tourplan.net/TourplanNX/index.html",
        "USERNAME": "poner minusculas",
        "PASSWORD": "password",
        "SERVICE_TYPES": {
            "AC": (None, "Accommodation"),
            "BT": (None, "Bus Ticket"),
            "CR": (None, "Cruises"),
            "DE": (None, "Description"),
            "EN": (None, "Entrance Fees"),
            "EX": (None, "Experience"),
            "FB": (None, "Food & Beverage"),
            "FE": (None, "Cancellation Fee"),
            "FT": (None, "Flight Tickets"),
            "GU": (None, "Guide Services"),
            "OC": (None, "Operational Cost"),
            "PK": (None, "Package (Multidays)"),
            "TF": (None, "Transfers"),
            "TR": (None, "Transport"),
            "TT": (None, "Train Ticket"),
        },
    },
}

ENTORNO_ACTUAL = "NUEVO"  # ← cambiar acá según la corrida: "TEST" o "NUEVO"

_cfg_entorno = ENTORNOS[ENTORNO_ACTUAL]
BASE_URL = _cfg_entorno["BASE_URL"]
USERNAME = _cfg_entorno["USERNAME"]
PASSWORD = _cfg_entorno["PASSWORD"]
SERVICE_TYPES_CATALOG = {cod: nombre for cod, (_num, nombre) in _cfg_entorno["SERVICE_TYPES"].items()}
STYPE_SIDEBAR = {cod: num for cod, (num, _nombre) in _cfg_entorno["SERVICE_TYPES"].items() if num}

if BASE_URL.startswith("PONER_"):
    raise ValueError(
        f"Falta completar BASE_URL del entorno {ENTORNO_ACTUAL!r} en ENTORNOS "
        f"(y USERNAME/PASSWORD si también son placeholder) antes de correr.")

# ── Google Sheet como cola de trabajo (entrada Y salida) ─────────────
# Un solo Sheet, independiente del ENTORNO_ACTUAL elegido arriba —
# podés relevar cualquier Tourplan contra el mismo Sheet.
GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1-3djHzmcRK0ov1aiHrH58QpJQ8bz7oDqFHK_yBHp9iY/edit"
HOJA_PRODUCTOS   = "productos"
HOJA_VIGENCIAS   = "vigencias"

# La pestaña "productos" es un reporte real exportado de TP, con sus
# propias columnas — el script NO las renombra ni las reordena, sólo
# lee de estos nombres tal cual están en el Sheet.
COL_LOCATION     = "Loc"
COL_SUPPLIER     = "Supplier"
COL_SERVICE_TYPE = "Serv"
COL_CODIGO       = "Code"

# Columnas que el script necesita para funcionar como cola de trabajo
# (RATE FROM/RATE TO de entrada, ESTADO/OBSERVACIONES/TIMESTAMP de
# salida) y que agrega solo al final de las que ya tiene "productos",
# la primera vez que corre contra este Sheet, si todavía no existen.
COLS_A_AGREGAR = ["RATE FROM", "RATE TO", "ESTADO", "OBSERVACIONES", "TIMESTAMP"]

SS_DIR = "screenshots"
os.makedirs(SS_DIR, exist_ok=True)

# Límite de filas PENDIENTE de PRODUCTOS a procesar en esta corrida (0 =
# sin límite). Para la primera prueba contra Tourplan real conviene un
# número chico — misma idea que LIMIT_PRUEBA en el script hermano.
LIMIT_PRUEBA = 0

# Límite de códigos a procesar por fila cuando CODIGO viene vacío (0 =
# sin límite, releva todos los que encuentre listar_codigos_supplier).
# Útil para probar rápido contra un supplier con muchos códigos.
LIMIT_CODIGOS_PRUEBA = 0

VELOCIDAD = 1.0  # multiplicador de todos los time.sleep — subir si la red es lenta

MOSTRAR_CAPTURAS = False
_ss_n = [0]
_avisado_sin_ipython = [False]


# ── Helpers base (REUTILIZADO tal cual de los scripts hermanos) ────

def ss(driver, nombre):
    _ss_n[0] += 1
    p = f"{SS_DIR}/{_ss_n[0]:03d}_{nombre[:40]}_{int(time.time())}.png"
    driver.save_screenshot(p)
    print(f"  📸 {os.path.basename(p)}")
    if MOSTRAR_CAPTURAS:
        try:
            from IPython.display import display, Image as IPyImage
            display(IPyImage(p, width=900))
        except ImportError:
            if not _avisado_sin_ipython[0]:
                print("    ⚠ MOSTRAR_CAPTURAS=True pero no hay IPython disponible "
                      "(sólo se ve inline en Colab/Jupyter) — se sigue guardando en disco.")
                _avisado_sin_ipython[0] = True


def dump(driver, nombre):
    p = f"{SS_DIR}/{nombre}_{int(time.time())}.html"
    with open(p, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print(f"  💾 HTML: {p}")


def jc(driver, el):
    """JavaScript click — único método confiable en Angular."""
    driver.execute_script("arguments[0].click();", el)


def set_val(driver, el, value):
    driver.execute_script("""
        var inp = arguments[0], val = arguments[1];
        var setter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value').set;
        setter.call(inp, val);
        inp.dispatchEvent(new Event('input',  {bubbles:true}));
        inp.dispatchEvent(new Event('change', {bubbles:true}));
    """, el, value)


def wait(driver, css, t=12):
    return WebDriverWait(driver, t).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css)))


def crear_driver():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1366,911")
    opts.add_argument("--disable-gpu")
    if CHROMIUM_BIN:
        opts.binary_location = CHROMIUM_BIN
    drv_path = ChromeDriverManager().install()
    svc = Service(executable_path=drv_path)
    d = webdriver.Chrome(service=svc, options=opts)
    print(f"✅ Driver iniciado (Chrome: {CHROMIUM_BIN or 'default del sistema'})")
    return d


def login(driver):
    print("🔐 Login...")
    driver.get(f"{BASE_URL}#/login")
    time.sleep(6 * VELOCIDAD)
    ss(driver, "login_page")

    def _campos_visibles():
        return driver.execute_script("""
            function vis(e){return !!(e && (e.offsetWidth||e.offsetHeight
                                      ||e.getClientRects().length)
                                      && !e.disabled);}
            var txt = Array.from(document.querySelectorAll(
                "input[type='text'], input:not([type])")).filter(vis);
            var pwd = Array.from(document.querySelectorAll(
                "input[type='password']")).filter(vis);
            return [txt[0]||null, pwd[0]||null];
        """)

    u_el = p_el = None
    for _ in range(15):
        u_el, p_el = _campos_visibles()
        if u_el and p_el:
            break
        time.sleep(2)
    if not (u_el and p_el):
        ss(driver, "login_sin_campos")
        raise Exception("No aparecieron los campos de login (usuario/password)")

    def _set(el, valor):
        try: el.clear()
        except Exception: pass
        try: el.click()
        except Exception: pass
        try: el.send_keys(valor)
        except Exception:
            set_val(driver, el, valor)

    _set(u_el, USERNAME)
    _set(p_el, PASSWORD)
    time.sleep(0.5)

    clic = driver.execute_script("""
        function vis(e){return !!(e && (e.offsetWidth||e.offsetHeight
                                  ||e.getClientRects().length) && !e.disabled);}
        var b = Array.from(document.querySelectorAll(
            "button.login, button[type='submit'], button")).filter(vis)
            .find(function(x){return /log\\s*in|ingresar|entrar|sign\\s*in/i
                                     .test((x.innerText||'')) ||
                                     x.classList.contains('login');});
        if (b){ b.click(); return (b.innerText||'button.login').trim(); }
        return null;
    """)
    if not clic:
        try: p_el.send_keys(Keys.ENTER)
        except Exception: pass
    time.sleep(8 * VELOCIDAD)
    assert "login" not in driver.current_url.lower(), "Login falló"
    ss(driver, "post_login")
    print("✅ Login OK")


def logout(driver):
    """Cierra ventanas secundarias y hace logout real. El logout solo se
    considera OK si después del click aparece de nuevo el formulario de
    login (input password) o la URL vuelve a #/login — encontrar el
    texto "logged in as" NO es prueba de logout, significa que la sesión
    sigue abierta. Portado de tp-nx-app
    (scripts/valorizacion_pkg/tourplan_valorizacion_pkg_v3.py::logout) —
    este script no lo tenía, y Tourplan tiene licencias concurrentes
    limitadas."""
    print("\n🔒 Finalización: cerrando ventanas y haciendo logout...")
    try:
        principal = driver.window_handles[0]
        for h in driver.window_handles[1:]:
            try:
                driver.switch_to.window(h); driver.close()
            except Exception: pass
        driver.switch_to.window(principal)
    except Exception: pass

    def _click_item(regex):
        return driver.execute_script("""
            var rx = new RegExp(arguments[0], 'i');
            var els = Array.from(document.querySelectorAll(
                'li, label, a, button, span, div'));
            var best = null;
            for (var el of els){
                if (!el.offsetParent) continue;
                var t = (el.innerText || '').trim();
                if (!t || t.length > 40 || !rx.test(t)) continue;
                if (!best || t.length <= (best.innerText||'').trim().length)
                    best = el;
            }
            if (!best) return null;
            best.click();
            return (best.innerText || '').trim().slice(0, 40);
        """, regex)

    def _en_login():
        return driver.execute_script("""
            var pwd = document.querySelector("input[type='password']");
            return (pwd && pwd.offsetParent !== null) ||
                   /login/i.test(window.location.href);
        """)

    try:
        driver.get(f"{BASE_URL}#/home")
        time.sleep(4 * VELOCIDAD)
        ss(driver, "logout_home")

        clicked = None
        try:
            btn_panel = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#openUserPanel")))
            jc(driver, btn_panel)
            time.sleep(2 * VELOCIDAD)
            ss(driver, "logout_usermenu")
            btn_logout = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    "div.panelHeader tp-button button, div.panelHeader button")))
            texto_btn = (btn_logout.text or "Logout").strip()
            jc(driver, btn_logout)
            clicked = texto_btn or "Logout"
        except Exception as _e:
            print(f"    ⚠ Flujo #openUserPanel falló ({_e}) — fallback por texto")

        rx_logout = r"^(log\s?out|sign\s?out|cerrar sesi)"
        if not clicked:
            clicked = _click_item(rx_logout)
        if not clicked:
            padre = _click_item(r"logged in as")
            time.sleep(2 * VELOCIDAD)
            ss(driver, "logout_submenu")
            if padre:
                clicked = _click_item(rx_logout)

        if not clicked:
            candidatos = driver.execute_script("""
                var out = [];
                for (var el of document.querySelectorAll('*')){
                    if (!el.offsetParent || el.children.length > 0) continue;
                    var t = (el.innerText || '').trim();
                    if (t && t.length < 40 && /log|user|usuario/i.test(t)) out.push(t);
                }
                return out.slice(0, 25);
            """)
            print(f"      Textos visibles con 'log/user': {candidatos}")

        time.sleep(5 * VELOCIDAD)
        ss(driver, "logout_done")

        if clicked and _en_login():
            print(f"  🔓 Logout OK (click en '{clicked}', volvió al login)")
        elif clicked:
            print(f"  ⚠ Click en '{clicked}' pero NO volvió al login — logout no confirmado")
        else:
            print("  ⚠ No encontré la opción LOG OUT en el menú — logout no realizado")
    except Exception as e:
        print(f"  ⚠ Error en logout: {e}")


def hamburger(driver):
    img = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "nav img")))
    jc(driver, img)
    time.sleep(2.5 * VELOCIDAD)


def menu_item(driver, n_or_text):
    if isinstance(n_or_text, int):
        xpath = f"(//nav//ul/li)[{n_or_text}]/div/div"
    else:
        texto = n_or_text.upper()
        xpath = (f"//nav//ul/li[.//*[contains("
                 f"translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
                 f"'{texto}')]]/div/div")
    el = WebDriverWait(driver, 8).until(EC.presence_of_element_located((By.XPATH, xpath)))
    jc(driver, el)
    time.sleep(2 * VELOCIDAD)


# ── Búsqueda de producto (REUTILIZADO tal cual de
# extraccion_tarifas_vigentes.py — ver ahí para el detalle de origen).
# STYPE_SIDEBAR ya no es una constante fija: sale del entorno elegido
# arriba (ENTORNOS[ENTORNO_ACTUAL]) y puede venir vacío para algunos
# service types — en ese caso el fallback por texto de más abajo se
# encarga de encontrarlos igual. ──

class ProductoNoEncontrado(Exception):
    pass


def _completar_filtros_busqueda(driver, location, supplier, codigo, service_type):
    st_upper = (service_type or "").strip().upper()

    driver.get(f"{BASE_URL}#/home")
    time.sleep(2 * VELOCIDAD)
    driver.get(f"{BASE_URL}#/product")
    time.sleep(5 * VELOCIDAD)

    lupa = wait(driver, "#searchWrapper li:nth-of-type(2) button")
    jc(driver, lupa)
    time.sleep(3 * VELOCIDAD)

    try:
        WebDriverWait(driver, 4).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.parameters1 input")))
    except Exception:
        driver.execute_script("""
            var els = Array.from(document.querySelectorAll('li,button,a,div,span'));
            for (var el of els){
                if (!el.offsetParent) continue;
                var t = (el.innerText || '').trim().toUpperCase();
                if (t === 'SELECTION'){ el.click(); return true; }
            }
            return false;
        """)
        time.sleep(2 * VELOCIDAD)

    if st_upper:
        try:
            stype_num = STYPE_SIDEBAR.get(st_upper, "")
            _ok_st = False
            if stype_num:
                _ok_st = driver.execute_script("""
                    var num = arguments[0];
                    for (var li of document.querySelectorAll('li')){
                        if (!li.offsetParent) continue;
                        var t = (li.textContent || '').trim();
                        if (t.indexOf(num) >= 0){ li.click(); return t.slice(0,40); }
                    }
                    return null;
                """, stype_num)
                if _ok_st:
                    time.sleep(1.5 * VELOCIDAD)
            if not _ok_st:
                patron_js = r"(^|[\s\-–—/(])" + re.escape(st_upper) + r"([\s\-–—/)]|$)"
                resultado = None
                for _ in range(3):
                    resultado = driver.execute_script("""
                        var rx = new RegExp(arguments[0]);
                        var elegido = null;
                        for (var li of document.querySelectorAll('li')){
                            if (!li.offsetParent) continue;
                            var t = (li.textContent || '').trim();
                            if (!t || t.length > 80) continue;
                            if (!elegido && rx.test(t.toUpperCase())) elegido = li;
                        }
                        if (elegido){ elegido.click(); return true; }
                        return false;
                    """, patron_js)
                    if resultado:
                        break
                    time.sleep(1.5 * VELOCIDAD)
        except Exception as e:
            print(f"    ⚠ Error seleccionando service type: {e}")

    if location:
        try:
            inp_loc = wait(driver, "div.parameters1 li:nth-of-type(1) input")
            inp_loc.click(); time.sleep(0.3)
            set_val(driver, inp_loc, location)
            time.sleep(1.5 * VELOCIDAD)
            row = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "tr.selectedRow > td.description")))
            jc(driver, row); time.sleep(1)
        except Exception:
            pass

    if supplier:
        try:
            inp_sup = wait(driver, "div.parameters1 li:nth-of-type(2) input")
            inp_sup.click(); time.sleep(0.3)
            set_val(driver, inp_sup, supplier)
            time.sleep(1.5 * VELOCIDAD)
            row_sup = WebDriverWait(driver, 5).until(EC.presence_of_element_located(
                (By.XPATH, f"//div[contains(@class,'parameters1')]//td[normalize-space(text())='{supplier.upper()}']")))
            jc(driver, row_sup); time.sleep(1)
        except Exception:
            try:
                inp_sup2 = driver.find_element(By.CSS_SELECTOR, "div.parameters1 li:nth-of-type(2) input")
                inp_sup2.send_keys(Keys.TAB); time.sleep(1)
            except Exception:
                pass

    if codigo:
        inp_cod = wait(driver, "div.parameters1 li:nth-of-type(3) input")
        inp_cod.click(); time.sleep(0.3)
        set_val(driver, inp_cod, codigo)
        time.sleep(0.5)

    btn_search = wait(driver, "#productSearchFilter li:nth-of-type(4) button")
    jc(driver, btn_search)
    time.sleep(6 * VELOCIDAD)


def buscar_producto(driver, location, supplier, codigo, service_type=None):
    """Busca Location/Supplier/Code(/ServiceType) en Product Search,
    abre el resultado y confirma que quedó en contexto de producto
    (menú con RATES/UTILITIES/etc.)."""
    codigo = str(codigo).strip() if codigo not in (None, "") else ""
    location = str(location).strip() if location not in (None, "") else ""
    supplier = str(supplier).strip() if supplier not in (None, "") else ""
    st_upper = (service_type or "").strip().upper()
    print(f"\n  📦 Buscando: {location}/{supplier}/{codigo}"
          f"{('/' + st_upper) if st_upper else ''}")

    _completar_filtros_busqueda(driver, location, supplier, codigo, service_type)

    def _en_contexto_producto():
        try:
            img = WebDriverWait(driver, 4).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "nav img")))
            jc(driver, img)
            time.sleep(2 * VELOCIDAD)
            items = driver.execute_script("""
                return Array.from(document.querySelectorAll('nav ul > li')).map(function(li){
                    return (li.querySelector('div div')||li).innerText.trim().split('\\n')[0].toUpperCase();
                });
            """) or []
            ok = "RATES" in items
            jc(driver, img)
            time.sleep(1)
            return ok, items
        except Exception:
            return False, []

    clicked = driver.execute_script("""
        var cod = (arguments[0] || '').toUpperCase();
        var st  = (arguments[1] || '').toUpperCase();
        var rows = Array.from(document.querySelectorAll('table tbody tr'));
        function matchRow(tr){
            var tds = Array.from(tr.querySelectorAll('td'));
            if (cod){
                var hasCod = tds.some(function(td){
                    return td.children.length===0 && td.innerText.trim().toUpperCase()===cod;
                });
                if(!hasCod) return false;
            }
            if(!st) return true;
            return tds.some(function(td){
                return td.children.length===0 && td.innerText.trim().toUpperCase()===st;
            });
        }
        var match = null;
        for (var tr of rows){ if (matchRow(tr)){ match = tr; break; } }
        if (!match && cod){
            for (var tr of rows){
                if ((tr.innerText||'').toUpperCase().includes(cod)){ match = tr; break; }
            }
        }
        if (!match && !cod && rows.length){ match = rows[0]; }
        if (!match) return null;
        var tds = match.querySelectorAll('td');
        var target = tds.length > 4 ? tds[4] : (tds.length > 0 ? tds[tds.length-1] : match);
        target.click();
        return target.innerText.trim().slice(0,50) || 'clicked';
    """, codigo or "", st_upper)

    if not clicked:
        ss(driver, f"ps_sin_resultado_{(codigo or 'sin_codigo')[:10]}")
        raise ProductoNoEncontrado(
            f"Producto no encontrado (location={location!r} supplier={supplier!r} "
            f"codigo={codigo!r} service_type={service_type!r})")

    time.sleep(6 * VELOCIDAD)
    ok_ctx, menu_items = _en_contexto_producto()
    if not ok_ctx:
        ss(driver, f"ps_contexto_incorrecto_{(codigo or 'sin_codigo')[:10]}")
        raise ProductoNoEncontrado(
            f"Click en resultado OK pero no quedó en contexto de producto "
            f"(menú visto: {menu_items}) — codigo={codigo!r}")


def _scrapear_pagina_resultados(driver):
    return driver.execute_script(r"""
        var tables = Array.from(document.querySelectorAll('table'));
        var target = null, headers = [];
        for (var t of tables){
            var ths = Array.from(t.querySelectorAll('th')).map(h => h.innerText.trim());
            if (ths.some(h => /CODE/i.test(h))){ target = t; headers = ths; break; }
        }
        if (!target){
            for (var t of tables){
                if (t.querySelectorAll('tbody tr').length){ target = t; break; }
            }
        }
        if (!target) return {items: [], paginado: false, total_filas: 0, headers: []};

        var idxCode = headers.findIndex(h => /^CODE$/i.test(h));
        if (idxCode < 0) idxCode = headers.findIndex(h => /CODE/i.test(h) && !/SERVICE/i.test(h));
        var idxLoc  = headers.findIndex(h => /LOCATION/i.test(h));
        var idxDesc = headers.findIndex(h => /DESCRIPTION/i.test(h));

        var rows = Array.from(target.querySelectorAll('tbody tr'));
        var out = rows.map(function(tr){
            var tds = Array.from(tr.querySelectorAll('td'));
            var codigo, descripcion;
            if (idxCode >= 0 && tds[idxCode]){
                codigo = tds[idxCode].innerText.trim();
            } else {
                var textos = tds.map(function(td){ return td.innerText.trim(); }).filter(Boolean);
                codigo = textos.find(function(t, i){
                    return /^[A-Z0-9]{2,10}$/.test(t) && i !== idxLoc;
                }) || '';
            }
            if (idxDesc >= 0 && tds[idxDesc]){
                descripcion = tds[idxDesc].innerText.trim();
            } else {
                var textos2 = tds.map(function(td){ return td.innerText.trim(); }).filter(Boolean);
                descripcion = textos2.reduce(function(a, b){ return b.length > a.length ? b : a; }, '');
            }
            return {codigo: codigo, descripcion: descripcion};
        }).filter(function(r){ return r.codigo; });

        var paginado = !!document.querySelector("[class*='pagin' i], [class*='pager' i]");
        return {items: out, paginado: paginado, total_filas: rows.length, headers: headers};
    """)


def _hacer_scroll_resultados(driver):
    return bool(driver.execute_script("""
        function contenedorScroll(){
            var fila = document.querySelector('table tbody tr');
            if (!fila) return null;
            var cur = fila.closest('table');
            while (cur && cur !== document.body){
                if (cur.scrollHeight > cur.clientHeight + 5) return cur;
                cur = cur.parentElement;
            }
            return document.scrollingElement || document.body;
        }
        var c = contenedorScroll();
        if (!c) return false;
        var antes = c.scrollTop;
        c.scrollTop = c.scrollTop + c.clientHeight;
        return c.scrollTop > antes;
    """))


def listar_codigos_supplier(driver, location, supplier, service_type=None):
    """Busca Location/Supplier(/ServiceType) en Product Search SIN
    código (deja ese campo vacío) y devuelve TODOS los resultados —
    scrolleando la grilla (virtual scroll, sin botón de paginación)
    hasta que dejan de aparecer códigos nuevos — como lista de
    {codigo, descripcion}."""
    location = str(location).strip() if location not in (None, "") else ""
    supplier = str(supplier).strip() if supplier not in (None, "") else ""
    print(f"\n  📋 Listando códigos: {location or '(todas)'}/{supplier or '(todos)'}"
          f"/{service_type.upper() if service_type else '(todos)'}")

    _completar_filtros_busqueda(driver, location, supplier, "", service_type)
    ss(driver, f"listado_{supplier[:15]}")

    items_por_codigo = {}
    headers_vistos = []
    MAX_INTENTOS = 300
    UMBRAL_SIN_NUEVOS = 5
    intentos_sin_nuevos = 0
    for intento in range(1, MAX_INTENTOS + 1):
        resultado = _scrapear_pagina_resultados(driver)
        if not resultado:
            break
        headers_vistos = resultado.get("headers") or headers_vistos
        nuevos = 0
        for item in resultado.get("items", []):
            if item["codigo"] not in items_por_codigo:
                items_por_codigo[item["codigo"]] = item
                nuevos += 1
        if nuevos:
            print(f"    scroll {intento}: {nuevos} códigos nuevos "
                  f"({len(items_por_codigo)} acumulados)")
            intentos_sin_nuevos = 0
        else:
            intentos_sin_nuevos += 1

        if intentos_sin_nuevos >= UMBRAL_SIN_NUEVOS:
            break
        if not _hacer_scroll_resultados(driver):
            if intentos_sin_nuevos >= 2:
                break
        time.sleep(1.5 * VELOCIDAD)
    else:
        print(f"    ⚠ Llegué al tope de {MAX_INTENTOS} scrolls para "
              f"{location}/{supplier} — puede haber más códigos sin leer.")

    driver.execute_script("""
        var fila = document.querySelector('table tbody tr');
        if (!fila) return;
        var cur = fila.closest('table');
        while (cur && cur !== document.body){
            if (cur.scrollHeight > cur.clientHeight + 5){ cur.scrollTop = cur.scrollHeight; return; }
            cur = cur.parentElement;
        }
    """)
    time.sleep(2 * VELOCIDAD)
    resultado_final = _scrapear_pagina_resultados(driver)
    if resultado_final:
        headers_vistos = resultado_final.get("headers") or headers_vistos
        nuevos_final = 0
        for item in resultado_final.get("items", []):
            if item["codigo"] not in items_por_codigo:
                items_por_codigo[item["codigo"]] = item
                nuevos_final += 1
        if nuevos_final:
            print(f"    ⚠ Pasada final (scroll directo al fondo) encontró "
                  f"{nuevos_final} códigos que el loop incremental se había "
                  f"salteado — revisar VELOCIDAD/UMBRAL_SIN_NUEVOS si esto "
                  f"se repite seguido.")

    if not items_por_codigo:
        dump(driver, f"listado_vacio_{supplier[:15]}")
        print(f"    ⚠ No encontré resultados para {location}/{supplier} "
              f"(headers vistos: {headers_vistos})")
        return []
    print(f"    → {len(items_por_codigo)} códigos encontrados en total "
          f"(headers: {headers_vistos})")
    return list(items_por_codigo.values())


# ── Fechas de período (REUTILIZADO tal cual de extraccion_tarifas_vigentes.py) ──

MESES_ES = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
            "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}


def parsear_fecha(txt):
    """'01/Apr/2026', '01/04/2026' o '2026-04-01' → datetime. None si
    no matchea ningún formato conocido. Acepta también un datetime ya
    parseado (por si en algún momento la fuente deja de ser texto plano
    de Sheets)."""
    if not txt:
        return None
    if isinstance(txt, datetime):
        return txt
    txt = str(txt).strip()
    if len(txt) >= 10 and txt[4] == "-" and txt[7] == "-":
        try:
            return datetime.fromisoformat(txt[:10])
        except Exception:
            pass
    p = txt.split("/")
    if len(p) == 3:
        try:
            dia = int(p[0])
            mes_raw = p[1]
            mes = MESES_ES.get(mes_raw[:3].capitalize()) or int(mes_raw)
            anio_raw = int(p[2])
            anio = anio_raw + 2000 if anio_raw < 100 else anio_raw
            return datetime(anio, mes, dia)
        except Exception:
            pass
    return None


def _parse_rate_period(texto):
    """Parsea '01/Apr/2026 - 31/Aug/2026' → (datetime, datetime)."""
    m = re.search(r'(\d+/\w+/\d+)\s*[-–]\s*(\d+/\w+/\d+)', texto or "")
    if m:
        return parsear_fecha(m.group(1)), parsear_fecha(m.group(2))
    return None, None


def _periodos_en_rango(periodos, desde_str, hasta_str):
    """Índices de `periodos` (cada uno con 'texto'/'desde'/'hasta') cuyo
    rango de fechas se superpone con [desde_str, hasta_str]. Períodos
    sin fecha parseable se descartan con aviso, no se incluyen "por si
    acaso"."""
    objetivo_desde = parsear_fecha(desde_str)
    objetivo_hasta = parsear_fecha(hasta_str)
    idxs = []
    for i, p in enumerate(periodos):
        if not (p["desde"] and p["hasta"]):
            print(f"    ⚠ Período con fecha no parseable, se descarta: {p['texto']!r}")
            continue
        if p["desde"] <= objetivo_hasta and p["hasta"] >= objetivo_desde:
            idxs.append(i)
    return idxs


# ── Lista de RATES, SIN entrar a ningún período (NUEVO) ─────────────
# Selectores confirmados por inspección real de la grilla de lista (no
# de la grilla de costos de un período abierto, que usa otras clases).

def _abrir_rates(driver, codigo):
    """Producto ya en contexto (ver buscar_producto) → entra al tab
    RATES. Muestra la lista de períodos tal cual la entrega Tourplan,
    sin aplicar ningún filtro de Price Code (acá no hace falta: la
    vista "All Price Codes" sólo distorsiona los VALORES de costo
    dentro de un período abierto, no las columnas de esta lista)."""
    hamburger(driver)
    menu_item(driver, "RATES")
    time.sleep(3 * VELOCIDAD)
    ss(driver, f"rates_lista_{codigo[:10]}")


def _leer_vigencias(driver):
    """Lee TODAS las filas de la lista de RATES actualmente visible tal
    cual las devuelve Tourplan (ya ordenadas del período más reciente al
    más viejo — no hace falta ordenar nada acá), sin entrar a ningún
    período. Selectores confirmados: td.tpcol-rateperiod,
    td.tpcol-pricecodecode, td.tpcol-currencycode (aparece 2 veces por
    fila — Buy Currency y Sell Currency, en ese orden; ambas comparten
    la misma clase, así que se distinguen por posición dentro de la
    fila, no por clase), td.tpcol-saleperiod, td.tpcol-ratestatuses,
    td.tpcol-ratetexts (frecuentemente vacía), td.tpcol-ratenames."""
    filas = driver.execute_script("""
        var out = [];
        document.querySelectorAll('td.tpcol-rateperiod').forEach(function(d){
            var tr = d.closest('tr');
            if (!tr) return;
            function txt(sel, idx){
                var els = tr.querySelectorAll(sel);
                var el = els[idx || 0];
                return el ? (el.innerText || '').trim() : '';
            }
            out.push({
                rate_period:   (d.innerText || '').trim(),
                pc:            txt('td.tpcol-pricecodecode'),
                buy_currency:  txt('td.tpcol-currencycode', 0),
                sell_currency: txt('td.tpcol-currencycode', 1),
                sale_period:   txt('td.tpcol-saleperiod'),
                rate_status:   txt('td.tpcol-ratestatuses'),
                rate_text:     txt('td.tpcol-ratetexts'),
                rate_name:     txt('td.tpcol-ratenames'),
            });
        });
        return out;
    """) or []
    for f in filas:
        f["desde"], f["hasta"] = _parse_rate_period(f["rate_period"])
        f["texto"] = f["rate_period"]  # clave que espera _periodos_en_rango
    return filas


def leer_vigencias_codigo(driver, codigo):
    """Orquesta: entra a RATES del producto ya en contexto y lee la
    lista completa de períodos (sin abrir ninguno)."""
    _abrir_rates(driver, codigo)
    periodos = _leer_vigencias(driver)
    if not periodos:
        dump(driver, f"rates_sin_periodos_{codigo[:10]}")
    return periodos


# ── Google Sheets como cola de trabajo ───────────────────────────────
# Mismo patrón ESTADO/OBSERVACIONES/guardado fila a fila de la skill
# armando-excel-como-cola-de-trabajo, pero contra un Sheet real en vez
# de un .xlsx local — así el progreso sobrevive a un corte de Colab: la
# próxima corrida retoma leyendo el mismo Sheet, no un archivo que se
# perdió con el runtime.

VIGENCIAS_HEADERS = [
    "TIMESTAMP", "LOCATION", "SUPPLIER", "SERVICE TYPE", "CODIGO",
    "RATE PERIOD", "PC", "BUY CURRENCY", "SELL CURRENCY", "SALE PERIOD",
    "RATE STATUS", "RATE TEXT", "RATE NAME",
]

_COMENTARIOS_PRODUCTOS = {
    "RATE FROM": "ENTRADA (opcional, dd/mm/yyyy). Junto con RATE TO define un rango: "
                 "se exportan TODAS las filas (una por Price Code) de TODOS los "
                 "períodos que se solapen con él. Si ambos quedan vacíos, se exporta "
                 "el período más reciente/último tal cual lo entrega Tourplan, con "
                 "todas sus filas por Price Code.",
    "RATE TO": "ENTRADA (opcional, dd/mm/yyyy). Ver RATE FROM.",
    "ESTADO": "SALIDA. PENDIENTE = a procesar. Volver a poner PENDIENTE para "
              f"reprocesar esta fila. {COL_LOCATION}/{COL_SUPPLIER}/{COL_SERVICE_TYPE}/"
              f"{COL_CODIGO} son opcionales (vacío = todos en esa dimensión) pero "
              "deben venir completos al menos 2 de esos 4, en cualquier combinación, "
              "para no disparar una búsqueda sin acotar en Tourplan.",
    "OBSERVACIONES": "SALIDA. Detalle del resultado (cuántos períodos/códigos, o el error).",
    "TIMESTAMP": "SALIDA. Momento en que se procesó esta fila.",
}


def conectar_sheets():
    """Auth interactiva de Colab contra la cuenta de Google de quien
    corre la celda (pide permiso la primera vez de cada sesión) y
    devuelve los dos worksheets (productos, vigencias) del Sheet
    apuntado por GOOGLE_SHEET_URL."""
    print("🔐 Autenticando con Google (puede pedir elegir cuenta / dar permiso)...")
    _colab_auth.authenticate_user()
    creds, _ = _google_auth_default()
    gc = gspread.authorize(creds)

    sh = gc.open_by_url(GOOGLE_SHEET_URL)

    def _hoja(nombre):
        for ws in sh.worksheets():
            if ws.title.strip().lower() == nombre.strip().lower():
                return ws
        raise ValueError(
            f"No encontré la pestaña {nombre!r} en el Sheet {sh.title!r}. "
            f"Pestañas disponibles: {[w.title for w in sh.worksheets()]}")

    ws_productos = _hoja(HOJA_PRODUCTOS)
    ws_vigencias = _hoja(HOJA_VIGENCIAS)
    print(f"✅ Conectado a Sheets: {sh.title!r} "
          f"(pestañas {ws_productos.title!r} / {ws_vigencias.title!r})")
    return ws_productos, ws_vigencias


def asegurar_columnas_productos(ws):
    """Agrega a la pestaña "productos" las columnas que le falten de
    COLS_A_AGREGAR (RATE FROM/RATE TO/ESTADO/OBSERVACIONES/TIMESTAMP),
    siempre al final de las que ya tenía el reporte de TP — nunca
    renombra ni reordena columnas existentes. Devuelve {nombre_columna:
    índice de columna (1-based)} para TODAS las columnas de la hoja."""
    headers = ws.row_values(1)
    faltantes = [c for c in COLS_A_AGREGAR if c not in headers]
    if faltantes:
        nuevos_headers = headers + faltantes
        ws.update(range_name="A1", values=[nuevos_headers])
        for i, nombre in enumerate(faltantes, start=len(headers) + 1):
            texto = _COMENTARIOS_PRODUCTOS.get(nombre)
            if texto:
                try:
                    ws.insert_note(rowcol_to_a1(1, i), texto)
                except Exception:
                    pass  # nota es sólo documentación, no bloquea si falla
        print(f"📄 Columnas agregadas a la pestaña {HOJA_PRODUCTOS!r}: {faltantes} "
              f"(ESTADO queda vacío en todas las filas — marcar PENDIENTE a mano "
              f"las que se quieran correr).")
        headers = nuevos_headers
    return {h: i + 1 for i, h in enumerate(headers) if h}


def asegurar_columnas_vigencias(ws):
    """Crea la fila de headers en "vigencias" si la pestaña está vacía."""
    if not ws.row_values(1):
        ws.update(range_name="A1", values=[VIGENCIAS_HEADERS])
        print(f"📄 Encabezados creados en la pestaña {HOJA_VIGENCIAS!r}.")


def cargar_pendientes(ws_productos):
    """Lee todas las filas de "productos" y devuelve las que tienen
    ESTADO=PENDIENTE, ya traducidas a los nombres internos que espera
    procesar_fila_producto (LOCATION/SUPPLIER/SERVICE TYPE/CODIGO/
    RATE FROM/RATE TO) — el mapeo a las columnas reales del reporte
    (Loc/Supplier/Serv/Code) queda acá, no en el resto del script."""
    valores = ws_productos.get_all_values()
    if not valores:
        return []
    headers = valores[0]
    rows = []
    for row_idx, fila in enumerate(valores[1:], start=2):
        if not any(c.strip() for c in fila):
            continue
        d = dict(zip(headers, fila))
        rows.append({
            "LOCATION": d.get(COL_LOCATION, ""),
            "SUPPLIER": d.get(COL_SUPPLIER, ""),
            "SERVICE TYPE": d.get(COL_SERVICE_TYPE, ""),
            "CODIGO": d.get(COL_CODIGO, ""),
            "RATE FROM": d.get("RATE FROM", ""),
            "RATE TO": d.get("RATE TO", ""),
            "ESTADO": d.get("ESTADO", ""),
            "__row_idx__": row_idx,
        })
    pendientes = [r for r in rows if r["ESTADO"].strip().upper() == "PENDIENTE"]
    print(f"Sheet cargado: {len(rows)} fila(s) en {HOJA_PRODUCTOS!r}, "
          f"{len(pendientes)} PENDIENTE")
    return pendientes


def actualizar_fila_producto(ws_productos, col_idx, row_idx, estado, observaciones):
    ts = datetime.now().isoformat(timespec="seconds")
    updates = []
    for campo, valor in (("ESTADO", estado), ("OBSERVACIONES", observaciones), ("TIMESTAMP", ts)):
        col = col_idx.get(campo)
        if col:
            updates.append({"range": rowcol_to_a1(row_idx, col), "values": [[valor]]})
    if updates:
        ws_productos.batch_update(updates)


def agregar_filas_vigencias(ws_vigencias, filas):
    """Append-only: agrega las filas ya leídas a la pestaña "vigencias".
    No pisa nada de lo ya escrito por corridas/filas anteriores."""
    if not filas:
        return
    valores = [[fila.get(h, "") for h in VIGENCIAS_HEADERS] for fila in filas]
    ws_vigencias.append_rows(valores, value_input_option="USER_ENTERED")


# ── Lógica de negocio por fila de PRODUCTOS ─────────────────────────

def procesar_fila_producto(driver, row):
    """Procesa una fila de PRODUCTOS: resuelve uno o varios códigos,
    entra a RATES de cada uno y exporta el/los período(s) que
    correspondan según RATE FROM/RATE TO. Nunca lanza fuera de esta
    función — cualquier fallo de un código puntual se registra en
    OBSERVACIONES y se sigue con el próximo, sin frenar la fila.
    Devuelve (estado, observaciones, filas_vigencias).

    Ningún campo es obligatorio por sí solo — un campo vacío significa
    "todos" en esa dimensión (ej. SERVICE TYPE vacío = todos los service
    types). Pero para no disparar una búsqueda sin acotar en Tourplan,
    deben venir completos AL MENOS 2 de los 4 campos LOCATION/SUPPLIER/
    SERVICE TYPE/CODIGO, en cualquier combinación."""
    location = str(row.get("LOCATION") or "").strip()
    supplier = str(row.get("SUPPLIER") or "").strip()
    service_type = str(row.get("SERVICE TYPE") or "").strip()
    codigo = str(row.get("CODIGO") or "").strip()
    rate_from = row.get("RATE FROM")
    rate_to = row.get("RATE TO")

    campos = {"LOCATION": location, "SUPPLIER": supplier,
              "SERVICE TYPE": service_type, "CODIGO": codigo}
    completos = [k for k, v in campos.items() if v]
    if len(completos) < 2:
        return ("ERROR",
                 "Hacen falta al menos 2 de LOCATION/SUPPLIER/SERVICE TYPE/CODIGO "
                 "(vinieron completos: " + (", ".join(completos) or "ninguno") + ")",
                 [])

    if codigo:
        items = [{"codigo": codigo, "descripcion": ""}]
    else:
        items = listar_codigos_supplier(driver, location, supplier, service_type)
        if not items:
            return ("ERROR", "no encontré códigos para ese supplier/location/service type", [])
        if LIMIT_CODIGOS_PRUEBA:
            items = items[:LIMIT_CODIGOS_PRUEBA]

    filas_vigencias = []
    fallidos = []
    for item in items:
        cod = item["codigo"]
        try:
            buscar_producto(driver, location, supplier, cod, service_type=service_type)
            periodos = leer_vigencias_codigo(driver, cod)
            if not periodos:
                fallidos.append(f"{cod}: sin períodos de RATES")
                continue

            if rate_from or rate_to:
                idxs = _periodos_en_rango(periodos, rate_from, rate_to)
                if not idxs:
                    fallidos.append(f"{cod}: ningún período en el rango pedido")
                    continue
            else:
                # Sin rango: Tourplan ya lista los períodos del más
                # reciente al más viejo — el primero YA ES el "último".
                # Un mismo período puede tener varias filas (una por
                # Price Code) con el mismo Rate Period — se exportan
                # todas, no sólo la primera.
                idxs = [i for i, p in enumerate(periodos)
                        if p["rate_period"] == periodos[0]["rate_period"]]

            ts = datetime.now().isoformat(timespec="seconds")
            for i in idxs:
                p = periodos[i]
                filas_vigencias.append({
                    "TIMESTAMP": ts, "LOCATION": location, "SUPPLIER": supplier,
                    "SERVICE TYPE": service_type, "CODIGO": cod,
                    "RATE PERIOD": p["rate_period"], "PC": p["pc"],
                    "BUY CURRENCY": p["buy_currency"], "SELL CURRENCY": p["sell_currency"],
                    "SALE PERIOD": p["sale_period"], "RATE STATUS": p["rate_status"],
                    "RATE TEXT": p["rate_text"], "RATE NAME": p["rate_name"],
                })
        except ProductoNoEncontrado:
            fallidos.append(f"{cod}: no encontrado en Tourplan")
        except Exception as e:
            # WebDriverException y similares suelen traer un str(e) vacío o
            # sólo direcciones de memoria del binario de chromedriver (nada
            # útil para diagnosticar) — se guarda el traceback completo a
            # disco junto con screenshot + HTML de la página en ese momento,
            # y en OBSERVACIONES sólo el nombre del archivo para ir a mirarlo.
            ts_err = int(time.time())
            nombre_err = f"error_{cod[:10]}_{ts_err}"
            log_path = f"{SS_DIR}/{nombre_err}.txt"
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
            fallidos.append(f"{cod}: {type(e).__name__} — ver {log_path}")
            ss(driver, nombre_err)
            dump(driver, nombre_err)

    observaciones = f"{len(filas_vigencias)} período(s) exportado(s) de {len(items)} código(s)"
    if fallidos:
        observaciones += " — fallidos: " + "; ".join(fallidos[:5])
        if len(fallidos) > 5:
            observaciones += f" (+{len(fallidos) - 5} más)"

    estado = "OK" if filas_vigencias else "ERROR"
    return estado, observaciones, filas_vigencias


def main():
    print("=" * 60)
    print("  EXTRACCIÓN DE VIGENCIAS (lista de RATES, sin entrar a los períodos)")
    print(f"  Entorno: {ENTORNO_ACTUAL}  —  {BASE_URL}")
    print("=" * 60)
    print("  Service Types disponibles en este entorno:")
    for cod, nombre in SERVICE_TYPES_CATALOG.items():
        num = STYPE_SIDEBAR.get(cod, "")
        print(f"    {cod:<4} {nombre}" + (f"  (sidebar #{num})" if num else ""))
    print("=" * 60)
    t_inicio = time.time()

    ws_productos, ws_vigencias = conectar_sheets()
    col_idx = asegurar_columnas_productos(ws_productos)
    asegurar_columnas_vigencias(ws_vigencias)
    faltan = [c for c in (COL_LOCATION, COL_SUPPLIER, "ESTADO") if c not in col_idx]
    if faltan:
        raise ValueError(f"Faltan columnas obligatorias en la pestaña {HOJA_PRODUCTOS!r}: {faltan}")

    pendientes = cargar_pendientes(ws_productos)
    if not pendientes:
        print(f"\n⛔ Sin filas PENDIENTE en la pestaña {HOJA_PRODUCTOS!r}. "
              f"Completar ESTADO=PENDIENTE en las filas a procesar.")
        return
    if LIMIT_PRUEBA:
        print(f"⚠ LIMIT_PRUEBA={LIMIT_PRUEBA} — procesando sólo las primeras "
              f"{LIMIT_PRUEBA} filas de {len(pendientes)} pendientes.")
        pendientes = pendientes[:LIMIT_PRUEBA]

    driver = crear_driver()
    try:
        login(driver)

        for n, row in enumerate(pendientes, start=1):
            row_idx = row["__row_idx__"]
            print(f"\n{'─' * 60}")
            print(f"[{n}/{len(pendientes)}] Fila {row_idx}: "
                  f"LOCATION={row.get('LOCATION') or '(todas)'} "
                  f"SUPPLIER={row.get('SUPPLIER') or '(todos)'} "
                  f"SERVICE TYPE={row.get('SERVICE TYPE') or '(todos)'} "
                  f"CODIGO={row.get('CODIGO') or '(todos)'}")

            estado, observaciones, filas_vigencias = "ERROR", "Error desconocido", []
            try:
                estado, observaciones, filas_vigencias = procesar_fila_producto(driver, row)
            except Exception:
                # Un error en ESTA fila no frena el resto del batch — se
                # registra y se sigue con la próxima.
                estado = "ERROR"
                observaciones = traceback.format_exc(limit=3)
                ss(driver, f"error_fila_{row_idx}")

            print(f"  Estado: {estado} — {observaciones}")
            # Guardar INMEDIATAMENTE (filas de "vigencias" + estado de la
            # fila de "productos") — así una corrida cortada a mitad de
            # camino deja registro de lo ya procesado, y la próxima
            # corrida retoma leyendo el mismo Sheet.
            agregar_filas_vigencias(ws_vigencias, filas_vigencias)
            actualizar_fila_producto(ws_productos, col_idx, row_idx, estado, observaciones)

    finally:
        logout(driver)
        driver.quit()
        dur = int(time.time() - t_inicio)
        m, s = divmod(dur, 60)
        print(f"\n🏁 Fin. Duración: {m}m {s:02d}s")
        print(f"📄 Sheet: {GOOGLE_SHEET_URL}")


if __name__ == "__main__":
    main()

"""
═══════════════════════════════════════════════════════════════════════════
  TOURPLAN NX — Actualización de Rate Name / Rate Text por período + price code
═══════════════════════════════════════════════════════════════════════════

Script nuevo (Google Colab, celda única). Modifica el "Rate Name" y el
"Rate Text" de una o más tarifas existentes en Tourplan NX, identificando
primero el período (vigencia) y el price code correspondiente dentro de la
grilla de RATES del producto.

Basado en (solo como referencia de LECTURA, no se modifica):
  tourplan_valorizacion_pkg_v3.py — de ahí se reutiliza la lógica de
  identificación de período/price code dentro de la grilla de rates
  (_parse_rate_period, lectura de td.tpcol-rateperiod / td.tpcol-pricecodecode,
  matching por fecha con tolerancia ±2 días) y el patrón de login/logout,
  búsqueda de producto y escritura Angular con blur.

Flujo (según grabación Puppeteer/Chrome DevTools adjunta):
  1. Product Setup del producto ya localizado (#/product).
  2. Grilla de RATES → click en la fila del período de vigencia
     (td.tpcol-rateperiod).
  3. Click en tab "Rate Set" (#tptablabel-tabs-rateset).
  4. En div.top-panel: completar Rate Name (#rateName input) y/o
     Rate Text (#rateText input).
  5. Click SAVE (tp-button.save > button habilitado dentro del contexto activo).

Reglas de negocio:
  - Rate Name y Rate Text son independientes: campo vacío en el Excel =
    no tocar ese campo en Tourplan.
  - Si ambos vienen vacíos: no se entra a editar nada, fila queda OK sin cambios.
  - Un mismo período puede repetirse con distinto Price Code → se identifica
    por (fecha, price code), no por fecha sola.
  - Price Code vacío en el Excel = aplicar el cambio a TODOS los price codes
    de ese período (se recorren todas las filas que matchean la fecha).
  - Si no se encuentra el período, o el price code indicado no existe dentro
    de ese período: la fila queda ERROR: <detalle> y el batch sigue.
"""

# ── CAMBIOS ──
# v1.0 (2026-09-14): primera versión.
# v2.0 (2026-09-14): multi-entorno + Google Sheets + diagnóstico reforzado.

VERSION       = "2.0"
VERSION_FECHA = "2026-09-14"

# Multiplicador de tiempos de espera.
# 1.0 = entorno de prueba (Test)  |  1.5 = producción (respuestas más lentas)
VELOCIDAD = 1.0

# MODO del script (vocabulario único, ver skill armando-excel-como-cola-de-trabajo):
#   "lectura" → dry-run: identifica período/price code y muestra qué escribiría,
#               NO hace click en SAVE, no persiste nada. Útil para validar el
#               Excel antes de aplicar cambios reales.
#   "aplicar" → ejecuta los cambios de verdad.
MODO = "aplicar"

# ── PASO 0: Entorno ──
import importlib.util
import subprocess
import sys

_PIPS_NEEDED = {
    "selenium": "selenium",
    "webdriver_manager": "webdriver-manager",
    "gspread": "gspread",
    "google.auth": "google-auth",
}


def asegurar_paquetes():
    print("🔧 Verificando entorno...\n")
    faltantes = [pkg for mod, pkg in _PIPS_NEEDED.items()
                 if importlib.util.find_spec(mod) is None]
    if faltantes:
        print(f"  ⏳ pip install {' '.join(faltantes)} ...")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "-q"] + faltantes,
            check=True,
        )
        print("  ✅ Paquetes Python OK")
    else:
        print("  ✅ Paquetes Python ya instalados")


asegurar_paquetes()

import os
import platform
import re
import shutil
import tempfile
import time
import traceback
from datetime import datetime

import gspread
from google.auth import default as _google_auth_default
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


# ── Detección/instalación de Chrome (multiplataforma) ──────────────────────
# Copia tal cual del skill arrancando-un-script-de-tourplan/scripts/chrome_bootstrap.py

def _chrome_version(path):
    try:
        r = subprocess.run([path, "--version"], capture_output=True, text=True, timeout=8)
        v = (r.stdout or r.stderr).strip()
        return v if v and any(c.isdigit() for c in v) else ""
    except Exception:
        return ""


def _find_chrome_windows():
    candidates = []
    for env_var in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        base = os.environ.get(env_var)
        if base:
            candidates.append(os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"))
    candidates += [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path, _chrome_version(path) or "versión no detectada"
    p = shutil.which("chrome.exe") or shutil.which("chrome")
    if p:
        return p, _chrome_version(p) or "versión no detectada"
    return None, ""


def _find_chrome_mac():
    path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    if os.path.exists(path):
        return path, _chrome_version(path) or "versión no detectada"
    return None, ""


def _find_chrome_linux():
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


def find_chrome():
    system = platform.system()
    if system == "Windows":
        return _find_chrome_windows()
    if system == "Darwin":
        return _find_chrome_mac()
    return _find_chrome_linux()


def _install_chrome_linux():
    DEVNULL = subprocess.DEVNULL
    e1 = e2 = None
    try:
        subprocess.run(
            "wget -qO- https://dl.google.com/linux/linux_signing_key.pub "
            "| gpg --dearmor -o /usr/share/keyrings/google-chrome.gpg",
            shell=True, check=True, stdout=DEVNULL, stderr=DEVNULL)
        with open("/etc/apt/sources.list.d/google-chrome.list", "w") as f:
            f.write("deb [arch=amd64 signed-by=/usr/share/keyrings/google-chrome.gpg] "
                    "https://dl.google.com/linux/chrome/deb/ stable main\n")
        subprocess.run(["apt-get", "update", "-qq"], stdout=DEVNULL, stderr=DEVNULL)
        subprocess.run(["apt-get", "install", "-y", "-qq", "google-chrome-stable"],
                        check=True, stdout=DEVNULL, stderr=DEVNULL)
        if find_chrome()[0]:
            return
    except Exception as exc:
        e1 = exc
    try:
        deb = "/tmp/google-chrome-stable.deb"
        url = "https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb"
        subprocess.run(["wget", "-q", "-O", deb, url], check=True)
        subprocess.run(["dpkg", "-i", deb], stdout=DEVNULL, stderr=DEVNULL)
        subprocess.run(["apt-get", "update", "-qq"], stdout=DEVNULL, stderr=DEVNULL)
        subprocess.run(["apt-get", "install", "-f", "-y", "-qq"],
                        check=True, stdout=DEVNULL, stderr=DEVNULL)
    except Exception as exc:
        e2 = exc
        raise EnvironmentError(f"No se pudo instalar Chrome automaticamente. Repo: {e1}  .deb: {e2}")


def find_or_prepare_chrome():
    path, version = find_chrome()
    if path:
        print(f"  Chrome OK: {path}  [{version}]")
        return path, version

    system = platform.system()
    if system == "Windows":
        raise EnvironmentError(
            "No se encontro Google Chrome instalado en esta PC.\n"
            "Instalalo desde https://www.google.com/chrome/ y volve a ejecutar.\n"
            "(Esta app no lo instala automaticamente porque requiere permisos de administrador.)"
        )
    if system == "Darwin":
        raise EnvironmentError(
            "No se encontro Google Chrome instalado en esta Mac.\n"
            "Instalalo desde https://www.google.com/chrome/ y volve a ejecutar."
        )

    print("  Chrome no encontrado, instalando (Linux)...")
    _install_chrome_linux()
    path, version = find_chrome()
    if not path:
        raise EnvironmentError("Chrome instalado pero no encontrado. Reintenta.")
    print(f"  Chrome instalado: {version}")
    return path, version


CHROMIUM_BIN, ver_chrome = find_or_prepare_chrome()


def crear_driver():
    opts = Options()
    # Colab no tiene servidor X ($DISPLAY) — a diferencia de la app de
    # escritorio (tp-nx-app, donde este mismo bloque corre SIN --headless
    # porque la persona ve la ventana de Chrome), acá Chrome no puede
    # levantar sin este flag: falla con "Missing X server or $DISPLAY" /
    # "session not created: Chrome instance exited".
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1704,1012")
    opts.add_argument("--disable-gpu")

    _profile_dir = tempfile.mkdtemp(prefix="tourplan_chrome_profile_")
    opts.add_argument(f"--user-data-dir={_profile_dir}")
    opts.add_argument("--no-first-run")
    opts.add_argument("--no-default-browser-check")

    if CHROMIUM_BIN:
        opts.binary_location = CHROMIUM_BIN
        print(f"Chrome binary: {CHROMIUM_BIN}  ({ver_chrome})")
    else:
        raise RuntimeError("No se encontró Chrome funcional.")

    log_path = os.path.join(tempfile.gettempdir(), "chromedriver.log")
    try:
        drv_path = ChromeDriverManager().install()
        print(f"Chromedriver: {drv_path}")
        svc = Service(executable_path=drv_path, log_output=log_path)
        d = webdriver.Chrome(service=svc, options=opts)
        print("✅ Driver iniciado")
        return d
    except Exception as e:
        if os.path.exists(log_path):
            with open(log_path) as f:
                print(f"\n--- ChromeDriver log ---\n{f.read()[-3000:]}\n---")
        raise RuntimeError(f"No se pudo iniciar Chrome.\nError: {e}")


# ── CONFIG: entornos Tourplan ───────────────────────────────────────────────
# Convención IMPORTANTE: BASE_URL ya trae el sufijo exacto que necesita la
# instalación (barra final si corresponde, o "index.html" sin barra).
# Todas las rutas hash se construyen como f"{BASE_URL}#/ruta".
#
# Para el entorno NUEVO, las posiciones del sidebar de Service Types todavía
# NO están confirmadas. None significa "usar fallback por texto (sigla)".
ENTORNOS = {
    "NUEVO": {
        "BASE_URL": "https://la-perwel.nx.tourplan.net/TourplanNX/index.html",
        "USERNAME": "PONER_USERNAME",
        "PASSWORD": "PONER_PASSWORD",
        "SERVICE_TYPES": {
            "AC": (None, ""),
            "BT": (None, ""),
            "CR": (None, ""),
            "DE": (None, ""),
            "EN": (None, ""),
            "EX": (None, ""),
            "FB": (None, ""),
            "FE": (None, ""),
            "FT": (None, ""),
            "GU": (None, ""),
            "OC": (None, ""),
            "PK": (None, ""),
            "TF": (None, ""),
            "TR": (None, ""),
            "TT": (None, ""),
        },
    },

    # Agregar otros Tourplan acá cuando se conozcan sus URLs/credenciales.
    # Ejemplo:
    # "TEST": {
    #     "BASE_URL": "https://.../TourplanNX_Test/",
    #     "USERNAME": "...",
    #     "PASSWORD": "...",
    #     "SERVICE_TYPES": {"HT": ("01", "Accommodation"), ...},
    # },
}

ENTORNO_ACTUAL = "NUEVO"  # ÚNICA variable a tocar para cambiar de Tourplan

_cfg = ENTORNOS[ENTORNO_ACTUAL]
BASE_URL = _cfg["BASE_URL"]
USERNAME = _cfg["USERNAME"]
PASSWORD = _cfg["PASSWORD"]
SERVICE_TYPES_CATALOG = {
    cod: nombre
    for cod, (_num, nombre) in _cfg["SERVICE_TYPES"].items()
}
STYPE_SIDEBAR = {
    cod: num
    for cod, (num, _nombre) in _cfg["SERVICE_TYPES"].items()
    if num
}

if BASE_URL.startswith("PONER_"):
    raise ValueError(f"Falta completar BASE_URL del entorno {ENTORNO_ACTUAL!r}.")
if USERNAME.startswith("PONER_") or PASSWORD.startswith("PONER_"):
    raise ValueError(f"Faltan credenciales del entorno {ENTORNO_ACTUAL!r}.")

# ── Google Sheets: cola persistente ────────────────────────────────────────
GOOGLE_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1BkayElPCluGVPYSeTjMgnRq8vyHpT_0q/edit"
)
HOJA_PRODUCTOS = "Datos"

# El reporte real de Tourplan se conserva tal cual: no se renombran ni
# reordenan sus columnas originales. El script mapea sus nombres a la lógica
# interna y agrega al FINAL solamente las columnas que falten.
COL_LOCATION     = "Loc"
COL_SUPPLIER     = "Supplier"
COL_SERVICE_TYPE = "Serv"
COL_CODIGO       = "Code"

COL_FECHA_DESDE = "RATE FROM"
COL_FECHA_HASTA = "RATE TO"
COL_PRICE_CODE = "Price Code"
COL_NUEVO_RATE_NAME = "Nuevo Rate Name"
COL_NUEVO_RATE_TEXT = "Nuevo Rate Text"

COLS_A_AGREGAR = [
    COL_FECHA_DESDE,
    COL_FECHA_HASTA,
    COL_PRICE_CODE,
    COL_NUEVO_RATE_NAME,
    COL_NUEVO_RATE_TEXT,
    "ESTADO",
    "OBSERVACIONES",
    "TIMESTAMP",
]

# No hay una pestaña de salida separada para este script: la operación solo
# modifica Tourplan y registra su resultado en la misma cola "Datos".
SS_DIR = f"screenshots_{ENTORNO_ACTUAL}"

# ── Helpers mínimos de DOM ───────────────────────────────────────────────

def jc(driver, el):
    """Click vía JavaScript — único método confiable en Angular."""
    driver.execute_script("arguments[0].click();", el)


def wait(driver, css, t=12):
    return WebDriverWait(driver, t).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css)))


def waitx(driver, xpath, t=12):
    return WebDriverWait(driver, t).until(
        EC.presence_of_element_located((By.XPATH, xpath)))


_ss_n = [0]


def ss(driver, nombre, ss_dir=SS_DIR):
    _ss_n[0] += 1
    nombre_seguro = re.sub(r"[^A-Za-z0-9_-]+", "_", nombre)[:60]
    os.makedirs(ss_dir, exist_ok=True)
    p = f"{ss_dir}/{_ss_n[0]:03d}_{nombre_seguro}_{int(time.time())}.png"
    try:
        ok = driver.save_screenshot(p)
        if not ok:
            print(f"  ⚠️ No se pudo guardar captura: {os.path.basename(p)}")
            return None
        print(f"  📸 {os.path.basename(p)}")
        return p
    except Exception as e:
        print(f"  ⚠️ Error guardando captura ({e})")
        return None


def dump(driver, nombre, ss_dir=SS_DIR):
    os.makedirs(ss_dir, exist_ok=True)
    p = f"{ss_dir}/{nombre}_{int(time.time())}.html"
    with open(p, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print(f"  💾 HTML: {p}")
    return p


def cerrar_nav_backdrop(driver, timeout=3, velocidad=1.0):
    """Limpia el backdrop (.tpnavbackdrop) que a veces deja colgado el menú
    hamburguesa tras cerrarlo. Ver skill buscando-productos-en-tourplan."""
    fin = time.time() + timeout * velocidad
    while time.time() < fin:
        if not driver.find_elements(By.CSS_SELECTOR, ".tpnavbackdrop"):
            return
        time.sleep(0.2)
    driver.execute_script("""
        document.querySelectorAll('.tpnavbackdrop').forEach(function(e){ e.remove(); });
    """)
    print("    ⚠ .tpnavbackdrop seguía presente tras cerrar el menú — removido a mano")


def esperar_fin_carga(driver, timeout=15, velocidad=1.0):
    """Tourplan muestra un <dialog> nativo 'PLEASE WAIT...' mientras termina
    de procesar la fila anterior — esperar a que cierre antes de clickear la
    lupa de búsqueda de la fila siguiente. Ver skill buscando-productos-en-tourplan."""
    fin = time.time() + timeout * velocidad
    while time.time() < fin:
        if not driver.find_elements(By.CSS_SELECTOR, "dialog[open]"):
            return
        time.sleep(0.3)
    print("    ⚠ El dialog de carga ('PLEASE WAIT...') seguía abierto tras esperar")


# ── Fechas / formatos Tourplan ───────────────────────────────────────────

MESES_ES = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
            "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}


def parsear_fecha(txt):
    """Acepta ISO (YYYY-MM-DD, típico de openpyxl) o dd/Mon/yyyy /
    dd/mm/yyyy / dd/mm/yy (típico de Tourplan/Excel)."""
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
            mes = MESES_ES.get(mes_raw[:3].capitalize(), None) or int(mes_raw)
            anio_raw = int(p[2])
            anio = anio_raw + 2000 if anio_raw < 100 else anio_raw
            return datetime(anio, mes, dia)
        except Exception:
            pass
    return None


def fmt_tp(dt):
    """dd/Mon/yyyy — formato que esperan los inputs de fecha de Tourplan."""
    meses = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
             7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}
    return f"{dt.day:02d}/{meses[dt.month]}/{dt.year}"


def _parse_rate_period(texto):
    """Parsea '01/Apr/2026 - 31/Aug/2026' → (datetime, datetime).
    (Tomado de tourplan_valorizacion_pkg_v3.py — misma lógica de matching
    por fecha con tolerancia, no se reinventa.)"""
    m = re.search(r'(\d+/\w+/\d+)\s*[-–]\s*(\d+/\w+/\d+)', texto)
    if m:
        return parsear_fecha(m.group(1)), parsear_fecha(m.group(2))
    return None, None


# ── Escritura Angular (set_val / set_val_con_blur) ───────────────────────

def set_val(driver, el, value):
    """Versión corta — inputs de filtro/búsqueda (no requieren blur)."""
    driver.execute_script("""
        var inp = arguments[0], val = arguments[1];
        var setter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value').set;
        setter.call(inp, val);
        inp.dispatchEvent(new Event('input',  {bubbles:true}));
        inp.dispatchEvent(new Event('change', {bubbles:true}));
    """, el, value)


def set_val_con_blur(driver, el, value):
    """Versión completa — necesaria para inputs tp-validator/tp-number que
    solo confirman el valor al perder foco (blur). Replica el flujo humano:
    focus → setter nativo → eventos → blur."""
    driver.execute_script("""
        var inp = arguments[0], val = arguments[1];
        inp.focus();
        inp.dispatchEvent(new Event('focus', {bubbles:true}));
        var setter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value').set;
        setter.call(inp, val);
        inp.dispatchEvent(new Event('input',  {bubbles:true}));
        inp.dispatchEvent(new Event('change', {bubbles:true}));
        inp.dispatchEvent(new KeyboardEvent('keyup', {bubbles:true}));
        inp.blur();
        inp.dispatchEvent(new Event('blur',     {bubbles:true}));
        inp.dispatchEvent(new Event('focusout', {bubbles:true}));
    """, el, value)


SAVE_POLL_INTERVAL = 0.4
SAVE_POLL_TIMEOUT = 15


def wait_save_habilitado(driver, save_selector, timeout=SAVE_POLL_TIMEOUT, interval=SAVE_POLL_INTERVAL):
    """Polling activo del botón SAVE (no un sleep fijo) — puede tardar en
    habilitarse tras escribir."""
    fin = time.time() + timeout
    encontrado_alguna_vez = False
    while time.time() < fin:
        elems = driver.find_elements(By.CSS_SELECTOR, save_selector)
        if elems:
            encontrado_alguna_vez = True
            try:
                if elems[0].is_displayed() and elems[0].is_enabled():
                    return elems[0], None
            except Exception:
                pass
        time.sleep(interval)
    if not encontrado_alguna_vez:
        return None, "el botón SAVE no se encontró en la página"
    return None, "el botón SAVE nunca se habilitó (timeout)"


def guardar_cambios(driver):
    """Clickea el tp-button.save > button ENABLED, priorizando el que vive
    dentro de un diálogo activo si lo hubiera (puede haber más de un
    tp-button.save en la página — el de nivel producto suele estar siempre
    disabled). En el flujo de Rate Name/Rate Text (confirmado por la
    grabación) el editor del período NO abre un tp-dialog — el SAVE vive en
    #productcostsview, dentro de la misma página — así que el criterio cae
    directo a "el que está enabled"; el chequeo de diálogo se deja igual
    como desempate defensivo por si en algún caso sí apareciera uno.
    (Mismo criterio de ranking que _escribir_rates() en
    tourplan_valorizacion_pkg_v3.py.)"""
    try:
        botones = driver.find_elements(By.CSS_SELECTOR, "tp-button.save > button")

        def _rank(b):
            try:
                en = b.is_enabled() and b.is_displayed()
                in_dlg = driver.execute_script(
                    "return !!arguments[0].closest('tp-dialog, tp-modal, [role=\"dialog\"]');", b)
            except Exception:
                en, in_dlg = False, False
            return (2 if (en and in_dlg) else 1 if en else 0)

        botones = sorted(botones, key=_rank, reverse=True)
        for b in botones:
            try:
                if b.is_displayed() and b.is_enabled():
                    in_dlg = driver.execute_script(
                        "return !!arguments[0].closest('tp-dialog, tp-modal, [role=\"dialog\"]');", b)
                    jc(driver, b)
                    return f"tp-button.save (click real{', dialog' if in_dlg else ''})", None
            except Exception:
                continue
    except Exception:
        pass

    guardado = driver.execute_script("""
        function vis(e){ return !!(e.offsetWidth || e.offsetHeight
                                   || e.getClientRects().length); }
        var saves = Array.from(document.querySelectorAll('tp-button.save > button'))
            .filter(function(b){ return vis(b) && !b.disabled; });
        saves.sort(function(a,b){
            var da = a.closest('tp-dialog,tp-modal,[role="dialog"]') ? 1 : 0;
            var db = b.closest('tp-dialog,tp-modal,[role="dialog"]') ? 1 : 0;
            return db - da;
        });
        if (saves.length){ saves[0].click(); return 'tp-button.save enabled (JS)'; }
        var bs = Array.from(document.querySelectorAll('button')).filter(vis);
        for (var b of bs){
            var t = (b.innerText || '').trim().toUpperCase();
            if (t === 'SAVE' && !b.disabled){ b.click(); return t + ' (JS)'; }
        }
        return null;
    """)
    if guardado:
        return guardado, None
    return None, "el botón SAVE no se encontró habilitado en ningún contexto"


# ── Login / logout / navegación ──────────────────────────────────────────

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
    for intento in range(15):
        u_el, p_el = _campos_visibles()
        if u_el and p_el:
            break
        time.sleep(2 * VELOCIDAD)
    if not (u_el and p_el):
        ss(driver, "login_sin_campos")
        raise Exception("No aparecieron los campos de login (usuario/password)")

    def _set(el, valor):
        try:
            el.clear()
        except Exception:
            pass
        try:
            el.click()
        except Exception:
            pass
        try:
            el.send_keys(valor)
        except Exception:
            driver.execute_script("""
                var el=arguments[0], v=arguments[1];
                var s=Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype,'value').set;
                s.call(el,v);
                el.dispatchEvent(new Event('input',{bubbles:true}));
                el.dispatchEvent(new Event('change',{bubbles:true}));
            """, el, valor)

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
        try:
            p_el.send_keys(Keys.ENTER)
        except Exception:
            pass
    time.sleep(8 * VELOCIDAD)
    assert "login" not in driver.current_url.lower(), "Login falló"
    ss(driver, "post_login")
    print("✅ Login OK")


def logout(driver):
    """Cierra ventanas secundarias y hace logout real. Ver skill
    arrancando-un-script-de-tourplan/references/licencias-y-sesiones.md:
    Tourplan tiene licencias concurrentes limitadas, una sesión colgada
    puede bloquear a otra persona real."""
    print("\n🔒 Finalización: cerrando ventanas y haciendo logout...")
    try:
        principal = driver.window_handles[0]
        for h in driver.window_handles[1:]:
            try:
                driver.switch_to.window(h)
                driver.close()
            except Exception:
                pass
        driver.switch_to.window(principal)
    except Exception:
        pass

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
            if padre:
                clicked = _click_item(rx_logout)

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
    """nav img = ícono hamburger en cualquier vista."""
    img = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "nav img")))
    jc(driver, img)
    time.sleep(2.5 * VELOCIDAD)
    try:
        items = driver.execute_script("""
            return Array.from(document.querySelectorAll('nav ul > li')).map(function(li,i){
                var txt = (li.querySelector('div div') || li).innerText.trim().split('\\n')[0];
                return (i+1) + ': ' + txt.slice(0,30);
            });
        """)
        print(f"    Menu items: {items}")
    except Exception:
        pass


def menu_item(driver, n_or_text):
    """Clickea un ítem del menú lateral, por posición (int) o texto (str)."""
    if isinstance(n_or_text, int):
        xpath = f"(//nav//ul/li)[{n_or_text}]/div/div"
    else:
        texto = n_or_text.upper()
        xpath = (f"//nav//ul/li[.//*[contains("
                 f"translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),"
                 f"'{texto}')]]/div/div")
    el = WebDriverWait(driver, 8).until(
        EC.presence_of_element_located((By.XPATH, xpath)))
    jc(driver, el)
    time.sleep(2 * VELOCIDAD)


# ── Búsqueda de producto (entra de lleno al registro — necesitamos RATES) ──

class ProductoNoEncontrado(Exception):
    pass


def buscar_producto(driver, location, supplier, codigo, service_type=None):
    """Localiza el producto por (location + supplier + código + service_type)
    y verifica contexto (menú con UTILITIES/RATES). Tomado de
    tourplan_valorizacion_pkg_v3.py / skill buscando-productos-en-tourplan —
    se usa esta variante (no buscar_y_abrir_option) porque necesitamos ENTRAR
    de lleno al registro para llegar a RATES."""
    st_upper = (service_type or "").strip().upper()
    st_str = f"/{st_upper}" if st_upper else ""
    print(f"\n  📦 Buscando: {location}/{supplier}/{codigo}{st_str}")

    driver.get(f"{BASE_URL}#/home")
    time.sleep(2 * VELOCIDAD)
    driver.get(f"{BASE_URL}#/product")
    time.sleep(5 * VELOCIDAD)
    ss(driver, f"ps_inicial_{codigo[:10]}")

    esperar_fin_carga(driver, velocidad=VELOCIDAD)

    lupa = wait(driver, "#searchWrapper li:nth-of-type(2) button")
    jc(driver, lupa)
    time.sleep(3 * VELOCIDAD)
    ss(driver, f"ps_modal_{codigo[:10]}")

    try:
        WebDriverWait(driver, 4).until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div.parameters1 input")))
    except Exception:
        tab_sel = driver.execute_script("""
            var els = Array.from(document.querySelectorAll('li,button,a,div,span'));
            for (var el of els){
                if (!el.offsetParent) continue;
                var t = (el.innerText || '').trim().toUpperCase();
                if (t === 'SELECTION'){ el.click(); return true; }
            }
            return false;
        """)
        print(f"    Tab SELECTION re-activado: {tab_sel}")
        time.sleep(2 * VELOCIDAD)

    if st_upper:
        try:
            st_buscar = STYPE_SIDEBAR.get(st_upper, st_upper)
            patron_js = (r"(^|[\s\-–—/(])" + re.escape(st_buscar) + r"([\s\-–—/)]|$)")
            resultado = None
            for intento in range(3):
                resultado = driver.execute_script("""
                    var rx = new RegExp(arguments[0]);
                    var textos = [], elegido = null;
                    for (var li of document.querySelectorAll('li')){
                        if (!li.offsetParent) continue;
                        var t = (li.textContent || '').trim();
                        if (!t || t.length > 80) continue;
                        textos.push(t);
                        if (!elegido && rx.test(t.toUpperCase())) elegido = li;
                    }
                    if (elegido){
                        elegido.click();
                        return {ok: true, texto: (elegido.textContent||'').trim().slice(0,40)};
                    }
                    return {ok: false, textos: textos.slice(0,40)};
                """, patron_js)
                if resultado and resultado.get("ok"):
                    break
                time.sleep(1.5)
            if resultado and resultado.get("ok"):
                time.sleep(1.5)
                print(f"    Service type '{st_upper}' seleccionado: {resultado.get('texto')}")
            else:
                print(f"    ⚠ Service type '{st_upper}' no encontrado en sidebar — buscando sin filtro")
                print(f"      Ítems visibles del sidebar: {(resultado or {}).get('textos', [])}")
        except Exception as e:
            print(f"    ⚠ Error seleccionando service type: {e}")

    if location:
        try:
            inp_loc = wait(driver, "div.parameters1 li:nth-of-type(1) input")
            inp_loc.click()
            time.sleep(0.3)
            set_val(driver, inp_loc, location)
            time.sleep(1.5)
            row = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "tr.selectedRow > td.description")))
            jc(driver, row)
            time.sleep(1)
            print(f"    Location: {location} OK")
        except Exception:
            try:
                rows = driver.find_elements(By.CSS_SELECTOR,
                    "div.parameters1 li:nth-of-type(1) table tbody tr")
                if rows:
                    jc(driver, rows[0])
                    time.sleep(1)
            except Exception:
                pass

    if supplier:
        try:
            inp_sup = wait(driver, "div.parameters1 li:nth-of-type(2) input")
            inp_sup.click()
            time.sleep(0.3)
            set_val(driver, inp_sup, supplier)
            time.sleep(1.5)
            row_sup = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located(
                    (By.XPATH,
                     f"//div[contains(@class,'parameters1')]//td[normalize-space(text())='{supplier.upper()}']")))
            jc(driver, row_sup)
            time.sleep(1)
            print(f"    Supplier: {supplier} OK")
        except Exception:
            try:
                inp_sup2 = driver.find_element(By.CSS_SELECTOR,
                    "div.parameters1 li:nth-of-type(2) input")
                inp_sup2.send_keys(Keys.TAB)
                time.sleep(1)
                print(f"    Supplier: {supplier} (TAB)")
            except Exception:
                pass

    inp_cod = wait(driver, "div.parameters1 li:nth-of-type(3) input")
    inp_cod.click()
    time.sleep(0.3)
    set_val(driver, inp_cod, codigo)
    time.sleep(0.5)
    print(f"    Code: {codigo}")
    ss(driver, f"ps_modal_lleno_{codigo[:10]}")

    btn_search = wait(driver, "#productSearchFilter li:nth-of-type(4) button")
    jc(driver, btn_search)
    time.sleep(6 * VELOCIDAD)
    ss(driver, f"ps_resultados_{codigo[:10]}")

    def _en_contexto_producto():
        try:
            img = WebDriverWait(driver, 4).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "nav img")))
            driver.execute_script("arguments[0].click();", img)
            time.sleep(2 * VELOCIDAD)
            items = driver.execute_script("""
                return Array.from(document.querySelectorAll('nav ul > li')).map(function(li){
                    return (li.querySelector('div div')||li).innerText.trim().split('\\n')[0].toUpperCase();
                });
            """) or []
            print(f"    Menu: {items}")
            PROD_ITEMS = {'UTILITIES', 'RATES', 'SEASONALITY', 'OPERATION', 'CONTENT', 'PRODUCT DETAILS'}
            ok = bool(set(items) & PROD_ITEMS)
            driver.execute_script("arguments[0].click();", img)
            time.sleep(1)
            cerrar_nav_backdrop(driver, velocidad=VELOCIDAD)
            return ok, items
        except Exception:
            return False, []

    clicked = driver.execute_script(f"""
        var cod = '{codigo}'.toUpperCase();
        var st  = '{st_upper}';
        var rows = Array.from(document.querySelectorAll('table tbody tr'));

        function matchRow(tr){{
            var tds = Array.from(tr.querySelectorAll('td'));
            var hasCod = tds.some(function(td){{
                return td.children.length===0 && td.innerText.trim().toUpperCase()===cod;
            }});
            if(!hasCod) return false;
            if(!st) return true;
            return tds.some(function(td){{
                return td.children.length===0 && td.innerText.trim().toUpperCase()===st;
            }});
        }}

        var match = null;
        for(var tr of rows){{ if(matchRow(tr)){{ match=tr; break; }} }}
        if(!match){{
            for(var tr of rows){{
                if((tr.innerText||'').toUpperCase().includes(cod)){{ match=tr; break; }}
            }}
        }}
        if(!match) return null;

        var tds = match.querySelectorAll('td');
        var target = tds.length > 4 ? tds[4] : (tds.length > 0 ? tds[tds.length-1] : match);
        target.click();
        return target.innerText.trim().slice(0,50) || 'clicked';
    """)

    if clicked:
        time.sleep(6 * VELOCIDAD)
        ss(driver, f"ps_cargado_{codigo[:10]}")
        ok_ctx, menu_items = _en_contexto_producto()
        if ok_ctx:
            print(f"  ✅ Producto {codigo} en contexto correcto (menú: {menu_items})")
            return
        print(f"    ⚠ Click ok ({clicked}) pero menú={menu_items}")
        return

    page_info = driver.execute_script(f"""
        return {{
            url: window.location.href,
            inputs: Array.from(document.querySelectorAll('input')).map(i=>i.value).filter(Boolean).slice(0,5),
            tables: Array.from(document.querySelectorAll('table tbody tr')).slice(0,5).map(r=>r.innerText.trim().slice(0,60))
        }};
    """)
    print(f"    Sin resultado: {page_info}")
    dump(driver, f"ps_sin_resultado_{codigo[:10]}")
    ss(driver, f"ps_sin_resultado_final_{codigo[:10]}")
    raise ProductoNoEncontrado(f"Producto '{codigo}' no encontrado en resultados")


# ── Localización de período + price code en la grilla de RATES ──────────

def _leer_periodos(driver):
    """Cada fila de la lista de rates trae la FECHA (td.tpcol-rateperiod) y su
    PRICE CODE (td.tpcol-pricecodecode). El orden coincide con
    find_elements('td.tpcol-rateperiod'). (Idéntico a
    tourplan_valorizacion_pkg_v3.py — misma lógica de lectura, no se reinventa.)"""
    return driver.execute_script("""
        var out = [];
        document.querySelectorAll('td.tpcol-rateperiod').forEach(function(d){
            var tr = d.closest('tr');
            var p = tr ? tr.querySelector('td.tpcol-pricecodecode') : null;
            out.push({date:(d.innerText||'').trim(),
                      pc:(p?(p.innerText||'').trim():'')});
        });
        return out;
    """)


def _periodos_matching(periodos, rf, rt, price_code):
    """Índices de 'periodos' cuya fecha coincide (±2 días, comparación por
    fecha parseada — nunca por texto exacto) con rf/rt.
    - price_code vacío  → TODOS los que matchean fecha (todas las variantes
      de Price Code de ese período — regla de negocio "PC vacío = todos").
    - price_code con valor → solo la fila de fecha+PC exactos."""
    pc_obj = (price_code or "").strip().upper()
    idxs = []
    for i, r in enumerate(periodos):
        pf, pt = _parse_rate_period(r["date"])
        if not (pf and pt and abs((pf - rf).days) <= 2 and abs((pt - rt).days) <= 2):
            continue
        if pc_obj:
            pc_row = (r["pc"] or "").strip().upper()
            if pc_row != pc_obj:
                continue
        idxs.append(i)
    return idxs


# ── Tab "Rate Set" (Rate Name / Rate Text) ───────────────────────────────

def _el_visible_unico(driver, css, t=10):
    """Devuelve el elemento VISIBLE que matchea `css`, no el primero que
    aparece en el DOM. Confirmado en corrida real (screenshot + HTML): al
    abrir varios períodos seguidos en la misma sesión, el diálogo de un
    período ya cerrado (ej. TR) NO se destruye del DOM, solo queda oculto
    — id="rateName"/"tptablabel-tabs-rateset" quedan duplicados, y
    `find_element`/`presence_of_element_located` (que devuelven el PRIMERO
    en orden de documento) agarraban el nodo viejo y oculto de TR en vez
    del diálogo realmente abierto del período nuevo. Por eso ND/EM leían y
    escribían sobre los inputs fantasma de TR."""
    fin = time.time() + t
    while time.time() < fin:
        el = driver.execute_script("""
            function vis(e){ return !!(e && (e.offsetWidth || e.offsetHeight
                                       || e.getClientRects().length)); }
            var els = Array.from(document.querySelectorAll(arguments[0])).filter(vis);
            return els.length ? els[els.length - 1] : null;
        """, css)
        if el is not None:
            return el
        time.sleep(0.3)
    raise Exception(f"No encontré ningún elemento VISIBLE para '{css}'")


def ir_a_tab_rate_set(driver):
    """Click en el tab Rate Set del período abierto — trae el panel superior
    (div.top-panel) con los campos Rate Name / Rate Text. Usa
    `_el_visible_unico` para no clickear un tab fantasma de un período
    previo ya cerrado (ver docstring de esa función)."""
    tab = _el_visible_unico(driver, "#tptablabel-tabs-rateset")
    jc(driver, tab)
    time.sleep(2 * VELOCIDAD)
    _el_visible_unico(driver, "div.top-panel")


def leer_rate_name_text(driver):
    """Lee el valor actual de #rateName / #rateText (div.top-panel) del
    diálogo VISIBLE, o None si no está presente."""
    def _val(field_id):
        try:
            el = _el_visible_unico(driver, f"#{field_id} input", t=5)
            return (el.get_attribute("value") or "").strip()
        except Exception:
            return None
    return _val("rateName"), _val("rateText")


def escribir_rate_name_text(driver, nuevo_name, nuevo_text):
    """Escribe SOLO los campos no vacíos (regla de negocio: campo vacío en
    el Excel = no tocar ese campo en Tourplan), sobre el input VISIBLE del
    diálogo actual. Devuelve la lista de campos que sí se escribieron."""
    escrito = []
    if nuevo_name:
        inp = _el_visible_unico(driver, "#rateName input")
        set_val_con_blur(driver, inp, nuevo_name)
        escrito.append("Rate Name")
        time.sleep(0.3 * VELOCIDAD)
    if nuevo_text:
        inp = _el_visible_unico(driver, "#rateText input")
        set_val_con_blur(driver, inp, nuevo_text)
        escrito.append("Rate Text")
        time.sleep(0.3 * VELOCIDAD)
    return escrito


def asegurar_lista_rates(driver, timeout=3):
    """Confirma que la grilla de RATES esté visible. Confirmado por
    grabación real: al hacer SAVE en el editor de un período, este se
    CIERRA SOLO y vuelve a la lista — no hace falta pasar por el menú
    hamburguesa para abrir el siguiente período, alcanza con clickear
    directamente la fila siguiente. Esta función es solo una red de
    seguridad para el caso borde en que el período anterior haya quedado
    abierto porque no hubo SAVE que lo cerrara (ej. 'sin cambios'): ahí sí
    hace falta forzar la vuelta a la lista por el menú."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "td.tpcol-rateperiod")))
        return
    except Exception:
        pass
    hamburger(driver)
    menu_item(driver, "RATES")
    time.sleep(4 * VELOCIDAD)


def procesar_periodo(driver, rf, rt, pc_real, nuevo_name, nuevo_text, codigo):
    """Abre el período (rf, rt, pc_real), escribe Rate Name/Text, guarda y
    verifica releyendo. NO navega por el menú entre períodos (ver
    `asegurar_lista_rates`) — el SAVE ya deja la grilla de RATES visible
    para poder clickear directamente la fila del próximo período. Devuelve
    (ok: bool, detalle: str)."""
    etiqueta_pc = pc_real or "Unassigned"

    asegurar_lista_rates(driver)
    periodos = _leer_periodos(driver)
    idxs = _periodos_matching(periodos, rf, rt, pc_real)
    if not idxs:
        return False, f"período {fmt_tp(rf)}–{fmt_tp(rt)} [{etiqueta_pc}] no aparece en la lista de rates"
    idx = idxs[0]
    filas_td = driver.find_elements(By.CSS_SELECTOR, "td.tpcol-rateperiod")
    jc(driver, filas_td[idx])
    time.sleep(5 * VELOCIDAD)
    ss(driver, f"periodo_abierto_{codigo[:10]}")

    ir_a_tab_rate_set(driver)
    viejo_name, viejo_text = leer_rate_name_text(driver)
    print(f"    [{etiqueta_pc}] abrí idx={idx} fecha='{periodos[idx]['date']}' "
          f"pc_en_grilla='{periodos[idx]['pc']}' — actual: "
          f"Rate Name='{viejo_name}' Rate Text='{viejo_text}'")

    escrito = escribir_rate_name_text(driver, nuevo_name, nuevo_text)
    if not escrito:
        return True, f"[{etiqueta_pc}] sin cambios (nada para escribir)"

    target_name = nuevo_name or viejo_name
    target_text = nuevo_text or viejo_text

    if MODO == "lectura":
        return True, (f"[{etiqueta_pc}] [LECTURA] se escribiría {'/'.join(escrito)}: "
                       f"Rate Name '{viejo_name}'→'{target_name}', "
                       f"Rate Text '{viejo_text}'→'{target_text}'")

    guardado, save_err = guardar_cambios(driver)
    if not guardado:
        if viejo_name == target_name and viejo_text == target_text:
            return True, f"[{etiqueta_pc}] sin cambios (ya tenía estos valores)"
        ss(driver, f"save_error_{codigo[:10]}")
        dump(driver, f"save_error_{codigo[:10]}")
        return False, f"[{etiqueta_pc}] no encontré botón SAVE habilitado ({save_err})"
    print(f"    SAVE clickeado ('{guardado}')")
    time.sleep(4 * VELOCIDAD)

    # ── Verificación post-SAVE: el SAVE ya cerró el período y volvió a la
    # grilla (confirmado por grabación) — reabrir el mismo período desde
    # ahí y releer, sin pasar por el menú.
    asegurar_lista_rates(driver)
    periodos_v = _leer_periodos(driver)
    idxs_v = _periodos_matching(periodos_v, rf, rt, pc_real)
    if not idxs_v:
        return False, f"[{etiqueta_pc}] verificación post-SAVE: el período no aparece al recargar RATES"
    filas_td_v = driver.find_elements(By.CSS_SELECTOR, "td.tpcol-rateperiod")
    jc(driver, filas_td_v[idxs_v[0]])
    time.sleep(5 * VELOCIDAD)
    ir_a_tab_rate_set(driver)
    name_final, text_final = leer_rate_name_text(driver)

    difs = []
    if nuevo_name and name_final != nuevo_name:
        difs.append(f"Rate Name esperado '{nuevo_name}', leído '{name_final}'")
    if nuevo_text and text_final != nuevo_text:
        difs.append(f"Rate Text esperado '{nuevo_text}', leído '{text_final}'")

    if difs:
        ss(driver, f"verif_error_{codigo[:10]}")
        dump(driver, f"verif_error_{codigo[:10]}")
        return False, f"[{etiqueta_pc}] " + "; ".join(difs)

    print(f"    ✔ Verificación post-SAVE OK [{etiqueta_pc}]")
    return True, f"[{etiqueta_pc}] OK — {'/'.join(escrito)} actualizado y verificado"


# ── Lógica de negocio por fila del Excel ─────────────────────────────────

def process_row(driver, row, col_idx):
    """Devuelve (estado, observaciones). No debe lanzar excepciones no
    capturadas hacia el loop principal — se resuelven acá en un ERROR."""
    location = str(row.get("_LOCATION") or "").strip()
    supplier = str(row.get("_SUPPLIER") or "").strip()
    service_type = str(row.get("_SERVICE_TYPE") or "").strip()
    codigo = str(row.get("_CODIGO") or "").strip()
    price_code = str(row.get("_PRICE_CODE") or "").strip()
    nuevo_name = str(row.get("_NUEVO_RATE_NAME") or "").strip()
    nuevo_text = str(row.get("_NUEVO_RATE_TEXT") or "").strip()

    # Regla de negocio: ambos campos vacíos → no se toca nada.
    if not nuevo_name and not nuevo_text:
        return "OK", "Rate Name y Rate Text vacíos — no se modifica nada"

    rf = parsear_fecha(row.get("_FECHA_DESDE"))
    rt = parsear_fecha(row.get("_FECHA_HASTA"))
    if not (rf and rt):
        return ("ERROR: fechas inválidas",
                f"RATE FROM/RATE TO inválidos ('{row.get('_FECHA_DESDE')}' / '{row.get('_FECHA_HASTA')}')")

    if not (location and supplier and service_type and codigo):
        return ("ERROR: faltan datos",
                "Faltan datos obligatorios (Location/Supplier/Service Type/Codigo)")

    try:
        buscar_producto(driver, location, supplier, codigo, service_type=service_type)
    except ProductoNoEncontrado as e:
        return f"ERROR: producto no encontrado", str(e)

    hamburger(driver)
    menu_item(driver, "RATES")
    time.sleep(4 * VELOCIDAD)
    ss(driver, f"rates_lista_{codigo[:10]}")

    periodos = _leer_periodos(driver)
    print(f"    Períodos leídos en {codigo}: "
          + ", ".join(f"{p['date']}[{p['pc'] or 'Unassigned'}]" for p in periodos[:15])
          + (f" … (+{len(periodos) - 15})" if len(periodos) > 15 else ""))

    idxs = _periodos_matching(periodos, rf, rt, price_code)

    if not idxs:
        idxs_sin_pc = _periodos_matching(periodos, rf, rt, "")
        if idxs_sin_pc and price_code:
            pcs_existentes = sorted({(periodos[i]["pc"] or "Unassigned") for i in idxs_sin_pc})
            return ("ERROR: price code no existe",
                    f"Price Code '{price_code}' no existe en el período "
                    f"{fmt_tp(rf)}–{fmt_tp(rt)} (existen: {', '.join(pcs_existentes)})")
        return ("ERROR: período no encontrado",
                f"No encontré el período {fmt_tp(rf)}–{fmt_tp(rt)} en {codigo}")

    # Price Code vacío = aplicar a TODOS los price codes de ese período
    # (regla de negocio): _periodos_matching ya devuelve, en ese caso,
    # TODOS los índices que matchean la fecha sin filtrar por PC — hay que
    # procesarlos todos, no solo el primero.
    pcs_a_procesar = [(periodos[i]["pc"] or "") for i in idxs]

    # No hace falta navegar entre price codes: confirmado por grabación real
    # que el SAVE cierra el período y deja la grilla de RATES visible para
    # clickear directamente la fila siguiente (ver `asegurar_lista_rates`,
    # que además cubre el caso borde de que el período anterior haya
    # quedado abierto sin SAVE).
    resultados = []
    fails = 0
    for pc_real in pcs_a_procesar:
        try:
            ok, detalle = procesar_periodo(
                driver, rf, rt, pc_real, nuevo_name, nuevo_text, codigo
            )
        except Exception:
            # También instrumentar el catch-all por price code: es una fila
            # lógica individual dentro del batch y debe dejar evidencia completa.
            tb = traceback.format_exc()
            try:
                log_path, ss_path, html_path = guardar_diagnostico_fila(
                    driver, codigo, row.get("__row_idx__", "?")
                )
                detalle = (
                    f"[{pc_real or 'Unassigned'}] excepción inesperada — "
                    f"traceback: {log_path} | "
                    f"screenshot: {ss_path or 'no disponible'} | "
                    f"HTML: {html_path or 'no disponible'}"
                )
            except Exception:
                detalle = f"[{pc_real or 'Unassigned'}] excepción inesperada:\n{tb}"
            ok = False
        if not ok:
            fails += 1
        resultados.append(detalle)

    observaciones = " | ".join(resultados)
    if fails == 0:
        return "OK", observaciones
    return f"ERROR: {fails}/{len(pcs_a_procesar)} price code(s) con error", observaciones


# ── Google Sheets como cola de trabajo ─────────────────────────────────────
def conectar_sheets():
    """Autentica en Colab y devuelve la pestaña de trabajo 'Datos'."""
    from google.colab import auth as _colab_auth

    print("🔐 Autenticando Google Sheets...")
    _colab_auth.authenticate_user()
    creds, _ = _google_auth_default()
    gc = gspread.authorize(creds)
    sh = gc.open_by_url(GOOGLE_SHEET_URL)

    for ws in sh.worksheets():
        if ws.title.strip().lower() == HOJA_PRODUCTOS.strip().lower():
            print(f"📊 Google Sheet conectado: '{ws.title}'")
            return ws

    raise ValueError(
        f"No encontré la pestaña {HOJA_PRODUCTOS!r}. "
        f"Disponibles: {[w.title for w in sh.worksheets()]}"
    )


def asegurar_columnas_productos(ws):
    """Agrega al FINAL las columnas que necesita la cola, sin tocar las
    columnas originales del reporte de Tourplan."""
    headers = ws.row_values(1)
    if not headers:
        raise ValueError("La pestaña 'Datos' está vacía: falta el header del reporte.")

    faltantes = [c for c in COLS_A_AGREGAR if c not in headers]
    if faltantes:
        nuevos_headers = headers + faltantes
        # Siempre kwargs: compatible con gspread v5/v6.
        ws.update(range_name="A1", values=[nuevos_headers])
        headers = nuevos_headers
        print(f"  ➕ Columnas agregadas al final: {', '.join(faltantes)}")
    else:
        print("  ✅ Columnas de cola ya presentes")

    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}

    # Validar las columnas del reporte que la lógica de Tourplan necesita.
    obligatorias = [COL_LOCATION, COL_SUPPLIER, COL_SERVICE_TYPE, COL_CODIGO]
    faltan_reporte = [c for c in obligatorias if c not in col_idx]
    if faltan_reporte:
        raise ValueError(
            "El reporte de Tourplan no contiene las columnas obligatorias: "
            + ", ".join(faltan_reporte)
        )

    return headers, col_idx


def cargar_pendientes(ws, headers):
    """Lee la cola y traduce los nombres del reporte real a nombres internos."""
    values = ws.get_all_values()
    if not values:
        return []

    rows = []
    for row_idx, fila in enumerate(values[1:], start=2):
        # Pad para que zip no pierda columnas al final.
        fila = list(fila) + [""] * max(0, len(headers) - len(fila))
        if not any(str(v).strip() for v in fila):
            continue

        d = dict(zip(headers, fila))

        # ESTADO vacío NO equivale a PENDIENTE: evita disparar por accidente
        # cientos de filas nuevas del reporte.
        estado = str(d.get("ESTADO") or "").strip().upper()
        if estado != "PENDIENTE":
            continue

        d["_LOCATION"] = d.get(COL_LOCATION, "")
        d["_SUPPLIER"] = d.get(COL_SUPPLIER, "")
        d["_SERVICE_TYPE"] = d.get(COL_SERVICE_TYPE, "")
        d["_CODIGO"] = d.get(COL_CODIGO, "")
        d["_FECHA_DESDE"] = d.get(COL_FECHA_DESDE, "")
        d["_FECHA_HASTA"] = d.get(COL_FECHA_HASTA, "")
        d["_PRICE_CODE"] = d.get(COL_PRICE_CODE, "")
        d["_NUEVO_RATE_NAME"] = d.get(COL_NUEVO_RATE_NAME, "")
        d["_NUEVO_RATE_TEXT"] = d.get(COL_NUEVO_RATE_TEXT, "")
        d["__row_idx__"] = row_idx
        rows.append(d)

    return rows


def actualizar_fila_producto(ws, col_idx, row_idx, estado, observaciones):
    """Persiste ESTADO/OBSERVACIONES/TIMESTAMP en un solo request."""
    ts = datetime.now().isoformat(timespec="seconds")
    updates = []

    for campo, valor in (
        ("ESTADO", estado),
        ("OBSERVACIONES", observaciones),
        ("TIMESTAMP", ts),
    ):
        if campo in col_idx:
            updates.append({
                "range": _a1(row_idx, col_idx[campo]),
                "values": [[valor]],
            })

    if updates:
        ws.batch_update(updates)


def _a1(row, col):
    """Conversión simple fila/columna → A1, sin depender de openpyxl."""
    letras = ""
    n = col
    while n:
        n, rem = divmod(n - 1, 26)
        letras = chr(65 + rem) + letras
    return f"{letras}{row}"


def guardar_diagnostico_fila(driver, codigo, row_idx):
    """Guarda traceback + screenshot + HTML de una excepción inesperada."""
    os.makedirs(SS_DIR, exist_ok=True)
    seguro = re.sub(r"[^A-Za-z0-9_-]+", "_", str(codigo))[:10] or "sin_codigo"
    stamp = int(time.time())
    log_path = os.path.join(
        SS_DIR, f"error_{seguro}_fila{row_idx}_{stamp}.txt"
    )
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(
            f"ENTORNO={ENTORNO_ACTUAL}\n"
            f"BASE_URL={BASE_URL}\n"
            f"CODIGO={codigo}\n"
            f"FILA_SHEET={row_idx}\n\n"
        )
        f.write(traceback.format_exc())

    ss_path = ss(driver, f"error_{seguro}_fila{row_idx}", ss_dir=SS_DIR)
    html_path = dump(driver, f"error_{seguro}_fila{row_idx}", ss_dir=SS_DIR)

    return log_path, ss_path, html_path


# ── MAIN ──────────────────────────────────────────────────────────────────
def main():
    print("=" * 70)
    print(
        f"  📌 VERSION {VERSION} ({VERSION_FECHA}) · "
        f"MODO={MODO} · VELOCIDAD={VELOCIDAD}"
    )
    print(f"  Entorno: {ENTORNO_ACTUAL}")
    print(f"  Base URL: {BASE_URL}")
    print("  Actualización de Rate Name / Rate Text por período + price code")
    print("=" * 70)

    t_inicio = time.time()
    ws = conectar_sheets()
    headers, col_idx = asegurar_columnas_productos(ws)
    pendientes = cargar_pendientes(ws, headers)

    print(f"Filas PENDIENTE: {len(pendientes)}")
    if not pendientes:
        print(
            "\n⛔ Sin filas PENDIENTE. "
            "Las filas nuevas quedan vacías hasta que se marquen manualmente."
        )
        return

    driver = crear_driver()
    try:
        login(driver)

        for row in pendientes:
            row_idx = row["__row_idx__"]
            codigo = str(row.get("_CODIGO") or "").strip()
            print(f"\n{'─' * 70}")
            print(
                f"Fila {row_idx}: {codigo} "
                f"{row.get('_FECHA_DESDE')}–{row.get('_FECHA_HASTA')} "
                f"PC={row.get('_PRICE_CODE') or 'todos'}"
            )

            estado, observaciones = "ERROR", "Error desconocido"
            try:
                estado, observaciones = process_row(driver, row, col_idx)
            except Exception:
                estado = "ERROR"
                tb = traceback.format_exc()
                try:
                    log_path, ss_path, html_path = guardar_diagnostico_fila(
                        driver, codigo, row_idx
                    )
                    observaciones = (
                        f"{tb.strip()} | "
                        f"traceback: {log_path} | "
                        f"screenshot: {ss_path or 'no disponible'} | "
                        f"HTML: {html_path or 'no disponible'}"
                    )
                except Exception:
                    # El diagnóstico nunca debe impedir que la fila se marque ERROR.
                    observaciones = tb

            print(
                f"  Estado: {estado}"
                + (f" — {observaciones}" if observaciones else "")
            )

            # No hay resultado append-only separado en este script.
            # Primero se considera terminado el trabajo de Tourplan; recién
            # después se confirma ESTADO en Sheets. Si hay un corte entre
            # ambas cosas, la fila puede reprocesarse al quedar PENDIENTE.
            actualizar_fila_producto(
                ws, col_idx, row_idx, estado, observaciones
            )

    finally:
        try:
            logout(driver)
        finally:
            driver.quit()
            dur = int(time.time() - t_inicio)
            m, s = divmod(dur, 60)
            print(f"\n🏁 Fin. Duración: {m}m {s:02d}s")
            print(f"📊 Google Sheet: pestaña '{HOJA_PRODUCTOS}'")


if __name__ == "__main__":
    main()

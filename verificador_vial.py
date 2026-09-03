import os
import re
import sys
import json
import time
import shutil
import threading
import unicodedata
from datetime import datetime, date
from difflib import SequenceMatcher
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, WebDriverException

VERSION = "9.0"
URL_LOGIN = "https://sistemas.seguridad.mendoza.gov.ar/vialcaminera//servlet/com.ktksuitelr.mdlsgt.hlogin2"
URL_CONSULTA_HINT = "wpconsultaantecedentes"

# Estructura de la planilla Playa San Ignacio de Loyola
HEADER_ROW = 3
COL_FECHA = 2          # B
COL_TIPO = 3           # C
COL_MARCA_MODELO = 4   # D
COL_COLOR = 5          # E
COL_DOMINIO = 6        # F
COL_INTERVIENE = 10    # J
COL_SUMARIO = 11       # K
COL_OBS = 13           # M

# Por decisión del usuario NO se comparan Motor ni Chasis.
CAMPOS_COMPARADOS = ("TIPO VEHICULO", "MARCA/MODELO", "COLOR", "DOMINIO", "INTERVIENE", "DOMINIO↔ACTA")



# Diccionario editable de normalización Marca/Modelo.
DICCIONARIO_DEFAULT = {
    "MOTO MEL": "MOTOMEL",
    "MOTOMEL": "MOTOMEL",
    "BUSSINES": "BUSINESS",
    "BUSINES": "BUSINESS",
    "BUSSINESS": "BUSINESS",
    "ZANELLA SOL BUSSINES": "ZANELLA SOL BUSINESS",
    "ZANELLA SOL BUSINES": "ZANELLA SOL BUSINESS",
    "HONDA WAVE 110": "HONDA WAVE",
    "HONDA WAVE110": "HONDA WAVE",
    "CORVEN ENERGY 110": "CORVEN ENERGY",
    "GILERA SMASH 110": "GILERA SMASH",
    "YAMAHA YBR 125": "YAMAHA YBR"
}

def directorio_app():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def cargar_diccionario_modelos():
    ruta = os.path.join(directorio_app(), "diccionario_vehiculos.json")
    data = dict(DICCIONARIO_DEFAULT)
    try:
        if os.path.exists(ruta):
            with open(ruta, "r", encoding="utf-8") as f:
                extra = json.load(f)
            if isinstance(extra, dict):
                data.update({norm(k): norm(v) for k, v in extra.items()})
        else:
            with open(ruta, "w", encoding="utf-8") as f:
                json.dump(DICCIONARIO_DEFAULT, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return {norm(k): norm(v) for k, v in data.items()}

DICCIONARIO_MODELOS = None

def aplicar_diccionario_modelo(v):
    global DICCIONARIO_MODELOS
    if DICCIONARIO_MODELOS is None:
        DICCIONARIO_MODELOS = cargar_diccionario_modelos()
    n = normalizar_modelo(v)
    # Sustituciones de frase completa y luego parciales.
    if n in DICCIONARIO_MODELOS:
        n = DICCIONARIO_MODELOS[n]
    for k, val in sorted(DICCIONARIO_MODELOS.items(), key=lambda kv: len(kv[0]), reverse=True):
        if k and k in n:
            n = n.replace(k, val)
    return norm(n)

def combinar_marca_modelo(datos):
    marca = str(datos.get("Marca", "") or "").strip()
    modelo = str(datos.get("Modelo", "") or "").strip()
    # Muchos registros GeneXus usan 0 como modelo vacío.
    if norm(modelo) in {"", "0", "S D", "SD", "SIN MODELO"}:
        modelo = ""
    nm = norm(marca)
    nmod = norm(modelo)
    if nmod and nmod not in nm:
        return (marca + " " + modelo).strip()
    return marca or modelo

def norm(v):
    if v is None:
        return ""
    s = str(v).upper().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def compact(v):
    return re.sub(r"[^A-Z0-9]", "", norm(v))


def xpath_literal(s):
    s = str(s)
    if "'" not in s:
        return f"'{s}'"
    if '"' not in s:
        return f'"{s}"'
    parts = s.split("'")
    return "concat(" + ", \"'\", ".join(f"'{p}'" for p in parts) + ")"


def visible(el):
    try:
        return el.is_displayed()
    except Exception:
        return False


def fecha_excel(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    s = norm(v)
    m = re.search(r"(\d{1,2})[ /.-](\d{1,2})[ /.-](\d{2,4})", s)
    if m:
        d, mo, y = m.groups()
        if len(y) == 2:
            y = "20" + y
        return f"{int(d):02d}/{int(mo):02d}/{y}"
    return str(v or "").strip()


def normalizar_tipo(v):
    n = norm(v)
    aliases = {
        "MOTO": "MOTOCICLETA",
        "MOTOCICLETA": "MOTOCICLETA",
        "MOTOVEHICULO": "MOTOCICLETA",
        "AUTO": "AUTOMOVIL",
        "AUTOMOVIL": "AUTOMOVIL",
        "AUTOMOTOR": "AUTOMOVIL",
        "CAMIONETA": "CAMIONETA",
        "CAMION": "CAMION",
        "BICIMOTO": "BICIMOTO",
    }
    return aliases.get(n, n)


def normalizar_color(v):
    n = norm(v)
    cambios = {
        "NEGRA": "NEGRO", "NEGRO": "NEGRO",
        "BLANCA": "BLANCO", "BLANCO": "BLANCO",
        "ROJA": "ROJO", "ROJO": "ROJO",
        "VERDE": "VERDE", "GRIS": "GRIS",
        "AZUL": "AZUL", "AMARILLA": "AMARILLO", "AMARILLO": "AMARILLO",
    }
    toks = [cambios.get(t, t) for t in n.split()]
    return " ".join(toks)


def dominio_sin_chapa(v):
    n = norm(v)
    return n in {
        "", "S D", "SD", "S CHAPA", "SIN CHAPA", "SIN DOMINIO",
        "NO POSEE", "NO POSEE DOMINIO", "NO SE DIVISA", "NO SE OBSERVA"
    }


def normalizar_modelo(v):
    n = norm(v)
    reemplazos = {
        "BUSSINES": "BUSINESS",
        "BUSINES": "BUSINESS",
        "BUSSINESS": "BUSINESS",
        "MOTOMEL": "MOTOMEL",
    }
    toks = []
    for t in n.split():
        t = reemplazos.get(t, t)
        # La cilindrada no debe provocar una diferencia si el sistema no la muestra.
        if re.fullmatch(r"\d{2,4}CC", t) or t == "CC":
            continue
        toks.append(t)
    return " ".join(toks)


def tokens_significativos(v):
    stop = {"DE", "DEL", "LA", "EL", "LOS", "LAS", "Y", "N", "NRO", "NUMERO", "CALLE", "AV", "AVENIDA"}
    return [t for t in norm(v).split() if len(t) >= 2 and t not in stop]


def coincide_fecha(excel, web):
    f = fecha_excel(excel)
    return bool(f) and f in str(web or "")


def coincide_tipo(excel, web):
    return normalizar_tipo(excel) == normalizar_tipo(web) or compact(normalizar_tipo(excel)) in compact(normalizar_tipo(web))


def coincide_color(excel, web):
    a = normalizar_color(excel)
    b = normalizar_color(web)
    if not a:
        return True
    ta = set(a.split())
    tb = set(b.split())
    return bool(ta) and ta.issubset(tb)


def coincide_dominio(excel, web):
    if dominio_sin_chapa(excel):
        return dominio_sin_chapa(web)
    a = compact(excel)
    b = compact(web)
    return bool(a) and (a == b or a in b or b in a)


def coincide_marca_modelo(excel, web):
    a = aplicar_diccionario_modelo(excel)
    b = aplicar_diccionario_modelo(web)
    if not a:
        return True
    if not b:
        return False
    ca, cb = compact(a), compact(b)
    if ca == cb or ca in cb or cb in ca:
        return True
    ta = tokens_significativos(a)
    tb = tokens_significativos(b)
    if not ta:
        return True
    # Coincidencia por variables/palabras significativas.
    similares = 0
    for t in ta:
        if t in tb or any(SequenceMatcher(None, t, w).ratio() >= 0.78 for w in tb):
            similares += 1
    # La marca suele ser la primera variable y debe coincidir.
    marca_ok = (ta[0] in tb) or any(SequenceMatcher(None, ta[0], w).ratio() >= 0.86 for w in tb)
    return marca_ok and similares / len(ta) >= 0.60

def acronimo(v):
    toks = tokens_significativos(v)
    return "".join(t[0] for t in toks if t)


def coincide_interviene(excel, web):
    a = norm(excel)
    b = norm(web)
    if not a:
        return True
    ca, cb = compact(a), compact(b)
    if ca and ca in cb:
        return True

    # U.R.V.G.A. equivale a U.R.V. GENERAL ALVEAR.
    # Se compara también por iniciales de las palabras de la dependencia.
    ac_web = acronimo(b)
    ac_excel = acronimo(a)
    if ca and (ca in ac_web or ac_web in ca):
        return True
    if ac_excel and (ac_excel in ac_web or ac_web in ac_excel):
        return True

    ta = tokens_significativos(a)
    tb = tokens_significativos(b)
    if not ta:
        return True
    presentes = sum(1 for t in ta if t in tb)
    return presentes / len(ta) >= 0.60


def find_input_acta(driver, log=None):
    """Localiza el input REAL de Nro Acta usando proximidad visual al rótulo.

    En este sistema GeneXus hay contenedores cuyo texto completo también contiene
    "Nro Acta". La v3 podía tomar un input equivocado. Esta versión busca primero
    elementos hoja con el rótulo y elige el campo de texto visible más cercano,
    a la derecha y en la misma fila.
    """
    def _log(msg):
        if log:
            try:
                log(msg)
            except Exception:
                pass

    # 1) Reunir inputs de texto visibles y habilitados.
    inputs = []
    for e in driver.find_elements(By.CSS_SELECTOR, "input"):
        try:
            tipo = (e.get_attribute("type") or "text").lower()
            if tipo not in {"text", "search", "tel", "number", ""}:
                continue
            if not visible(e) or not e.is_enabled():
                continue
            r = e.rect
            if r.get("width", 0) <= 10 or r.get("height", 0) <= 5:
                continue
            inputs.append(e)
        except Exception:
            pass

    # 2) Buscar rótulos hoja cuyo texto sea exactamente Nro Acta (con o sin ':').
    labels = []
    xpath_labels = (
        "//*[not(*) and string-length(normalize-space(.)) > 0 "
        "and contains(translate(normalize-space(.),"
        "'abcdefghijklmnopqrstuvwxyzáéíóúñ','ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÑ'),'NRO ACTA')]"
    )
    for et in driver.find_elements(By.XPATH, xpath_labels):
        try:
            txt = norm(et.text)
            if txt in {"NRO ACTA", "NRO ACTA N", "N ACTA", "NUMERO ACTA"} or txt.startswith("NRO ACTA"):
                if visible(et):
                    labels.append(et)
        except Exception:
            pass

    # 3) Seleccionar por geometría: misma altura y a la derecha del rótulo.
    mejores = []
    for lab in labels:
        try:
            lr = lab.rect
            ly = lr.get("y", 0) + lr.get("height", 0) / 2
            lx_right = lr.get("x", 0) + lr.get("width", 0)
            for inp in inputs:
                ir = inp.rect
                iy = ir.get("y", 0) + ir.get("height", 0) / 2
                ix = ir.get("x", 0)
                dy = abs(iy - ly)
                dx = ix - lx_right
                # Debe estar aproximadamente en la misma fila y no muy lejos a la izquierda.
                if dy <= 32 and dx >= -20:
                    score = dy * 10 + max(dx, 0)
                    mejores.append((score, inp, lab))
        except Exception:
            pass

    if mejores:
        mejores.sort(key=lambda t: t[0])
        inp = mejores[0][1]
        try:
            _log("  Campo detectado: id='{}' name='{}' x={} y={}".format(
                inp.get_attribute("id") or "",
                inp.get_attribute("name") or "",
                round(inp.rect.get("x", 0), 1),
                round(inp.rect.get("y", 0), 1),
            ))
        except Exception:
            pass
        return inp

    # 4) Fallback por atributos GeneXus / nombre del control.
    candidatos = []
    for e in inputs:
        try:
            attrs = " ".join([
                e.get_attribute("id") or "",
                e.get_attribute("name") or "",
                e.get_attribute("placeholder") or "",
                e.get_attribute("title") or "",
            ])
            na = norm(attrs)
            if "ACTA" in na and "DOCUMENT" not in na:
                candidatos.append(e)
        except Exception:
            pass
    if candidatos:
        return candidatos[0]

    # 5) Último fallback: en la pantalla conocida, Nro Acta es el input de texto
    # de la columna izquierda que está debajo de Nro Documento. Elegimos por Y.
    if len(inputs) >= 2:
        try:
            # Agrupar por mitad izquierda de la pantalla y ordenar verticalmente.
            ancho = driver.execute_script("return window.innerWidth || document.documentElement.clientWidth") or 1400
            izquierdos = [e for e in inputs if e.rect.get("x", 0) < ancho * 0.55]
            izquierdos.sort(key=lambda e: e.rect.get("y", 0))
            if len(izquierdos) >= 2:
                return izquierdos[-1]
        except Exception:
            pass
    return None

def find_input_dominio(driver, log=None):
    def _log(msg):
        if log:
            try: log(msg)
            except Exception: pass
    inputs=[]
    for e in driver.find_elements(By.CSS_SELECTOR, "input"):
        try:
            tipo=(e.get_attribute("type") or "text").lower()
            if tipo not in {"text","search","tel","number",""} or not visible(e) or not e.is_enabled():
                continue
            if e.rect.get("width",0)>10 and e.rect.get("height",0)>5:
                inputs.append(e)
        except Exception: pass
    labels=[]
    xp="//*[not(*) and contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyzáéíóúñ','ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÑ'),'DOMINIO')]"
    for lab in driver.find_elements(By.XPATH,xp):
        try:
            if visible(lab) and norm(lab.text) in {"DOMINIO","NRO DOMINIO","N DOMINIO"}: labels.append(lab)
        except Exception: pass
    cand=[]
    for lab in labels:
        lr=lab.rect; ly=lr.get('y',0)+lr.get('height',0)/2; rx=lr.get('x',0)+lr.get('width',0)
        for inp in inputs:
            ir=inp.rect; iy=ir.get('y',0)+ir.get('height',0)/2; dx=ir.get('x',0)-rx; dy=abs(iy-ly)
            if dy<=35 and dx>=-20: cand.append((dy*10+max(dx,0),inp))
    if cand:
        cand.sort(key=lambda x:x[0]); e=cand[0][1]
        _log("  Campo Dominio detectado: id='{}' name='{}'".format(e.get_attribute('id') or '',e.get_attribute('name') or ''))
        return e
    for e in inputs:
        attrs=norm(' '.join([e.get_attribute('id') or '',e.get_attribute('name') or '',e.get_attribute('placeholder') or '']))
        if 'DOMINIO' in attrs: return e
    return None

def cargar_input_geneXus(driver, inp, valor):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
        inp.click(); inp.send_keys(Keys.CONTROL,'a'); inp.send_keys(Keys.BACKSPACE)
        for ch in str(valor):
            inp.send_keys(ch); time.sleep(0.035)
        inp.send_keys(Keys.TAB); time.sleep(0.25)
        actual=(inp.get_attribute('value') or '').strip()
        if compact(actual)==compact(valor): return True
        driver.execute_script("""const el=arguments[0],val=arguments[1]; const p=Object.getPrototypeOf(el); const d=Object.getOwnPropertyDescriptor(p,'value'); if(d&&d.set)d.set.call(el,val); else el.value=val; ['input','change','keyup','blur'].forEach(t=>el.dispatchEvent(new Event(t,{bubbles:true})));""", inp, str(valor))
        time.sleep(0.25)
        return compact(inp.get_attribute('value') or '')==compact(valor)
    except Exception:
        return False

def click_buscar(driver):
    xpaths = [
        "//button[contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'BUSCAR')]",
        "//input[(@type='button' or @type='submit') and contains(translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'BUSCAR')]",
        "//a[contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'BUSCAR')]",
    ]
    for xp in xpaths:
        for e in driver.find_elements(By.XPATH, xp):
            if visible(e) and e.is_enabled():
                driver.execute_script("arguments[0].click();", e)
                return True
    return False


def esperar_resultado_acta(driver, acta, timeout=10):
    lit = xpath_literal(acta)
    xp = f"//*[normalize-space(text())={lit}]"
    try:
        WebDriverWait(driver, timeout).until(lambda d: any(visible(e) for e in d.find_elements(By.XPATH, xp)))
        elems = [e for e in driver.find_elements(By.XPATH, xp) if visible(e)]
        return elems[-1] if elems else None
    except TimeoutException:
        return None


def clickable_from_element(el):
    try:
        tag = (el.tag_name or "").lower()
        if tag in {"a", "button", "input"}:
            return el
        if tag == "img":
            anc = el.find_elements(By.XPATH, "./ancestor::*[self::a or self::button][1]")
            return anc[0] if anc else el
    except Exception:
        pass
    return el


def click_tres_puntos(driver, acta_el):
    """Abre Información de Objeto desde la fila del resultado."""
    try:
        row = acta_el.find_elements(By.XPATH, "./ancestor::tr[1]")
        row = row[0] if row else None
    except Exception:
        row = None

    scope = row if row is not None else driver
    candidatos = []
    try:
        elems = scope.find_elements(By.XPATH, ".//*[self::a or self::button or self::input or self::img]" if row is not None else "//*[self::a or self::button or self::input or self::img]")
        seen = set()
        for raw in elems:
            e = clickable_from_element(raw)
            try:
                if not visible(e) or not e.is_enabled():
                    continue
                key = e.id
                if key in seen:
                    continue
                seen.add(key)
                r = e.rect
                if r.get("width", 0) <= 0 or r.get("height", 0) <= 0:
                    continue
                meta = " ".join([
                    e.text or "", e.get_attribute("title") or "", e.get_attribute("alt") or "",
                    e.get_attribute("aria-label") or "", e.get_attribute("class") or "",
                    e.get_attribute("src") or ""
                ])
                candidatos.append((e, norm(meta), r.get("x", 0)))
            except Exception:
                pass
    except Exception:
        pass

    # Preferencias semánticas.
    for palabra in ("DETALLE", "INFORMACION", "OBJETO", "ACCION", "OPCION", "MENU", "PUNTOS"):
        for e, meta, _ in candidatos:
            if palabra in meta:
                driver.execute_script("arguments[0].click();", e)
                return True
    for e, meta, _ in candidatos:
        if "..." in (e.text or "") or "•••" in (e.text or ""):
            driver.execute_script("arguments[0].click();", e)
            return True

    # En este sistema hay dos íconos al extremo derecho de la fila:
    # tres puntos y luego otro ícono. El botón de tres puntos es el segundo desde la derecha.
    if row is not None and len(candidatos) >= 2:
        candidatos.sort(key=lambda x: x[2])
        target = candidatos[-2][0]
        driver.execute_script("arguments[0].click();", target)
        return True
    return False


def _texto_contexto_actual(driver):
    try:
        return driver.find_element(By.TAG_NAME, "body").text or ""
    except Exception:
        return ""


def _buscar_contexto_con_texto(driver, frases, profundidad=0, max_profundidad=3):
    """Busca texto visible en documento principal o iframes y deja activo el contexto encontrado."""
    texto = norm(_texto_contexto_actual(driver))
    if all(norm(f) in texto for f in frases if f):
        return True

    if profundidad >= max_profundidad:
        return False

    try:
        frames = driver.find_elements(By.CSS_SELECTOR, "iframe, frame")
    except Exception:
        frames = []

    for fr in frames:
        try:
            if not visible(fr):
                continue
            driver.switch_to.frame(fr)
            if _buscar_contexto_con_texto(driver, frases, profundidad + 1, max_profundidad):
                return True
            driver.switch_to.parent_frame()
        except Exception:
            try:
                driver.switch_to.parent_frame()
            except Exception:
                pass
    return False


def _activar_contexto_info_objeto(driver):
    """Ubica el popup de Información de Objeto, incluso cuando GeneXus lo abre dentro de un iframe."""
    try:
        driver.switch_to.default_content()
    except Exception:
        pass

    # La captura del sistema muestra ambos textos; cualquiera de ellos identifica el popup.
    for frases in (["INFORMACION DE OBJETO"], ["ADMINISTRACION DE OBJETOS"], ["DATOS DESCRIPTIVOS", "INFRACCIONES"]):
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        if _buscar_contexto_con_texto(driver, frases):
            return True
    return False


def esperar_info_objeto(driver, timeout=10):
    fin = time.time() + timeout
    while time.time() < fin:
        if _activar_contexto_info_objeto(driver):
            return True
        time.sleep(0.25)
    return False


def click_lupa(driver, acta=None, log=None):
    """Abre Objetos Datos Descriptivos y confirma que la ventana realmente apareció.

    v6: el ícono amarillo de lupa del sistema GeneXus no siempre expone href/onclick
    en el IMG. Por eso se localiza la fila de datos y se prueban tanto el control
    real como puntos físicos dentro de la primera celda (donde está la lupa).
    """
    def _log(msg):
        if log:
            try:
                log(msg)
            except Exception:
                pass

    def _probar_y_confirmar(el, descripcion):
        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center',inline:'center'});", el)
            time.sleep(0.15)
            # click nativo primero; si GeneXus sólo escucha eventos del mouse, usar MouseEvent.
            try:
                el.click()
            except Exception:
                driver.execute_script("arguments[0].click();", el)
            if acta:
                texto = esperar_datos_descriptivos(driver, acta, timeout=2.2)
                if texto:
                    _log(f"  -> DATOS DESCRIPTIVOS abiertos ({descripcion})")
                    return texto
            else:
                return True
        except Exception:
            pass
        return None

    # Asegurar contexto del popup Información de Objeto.
    if not _activar_contexto_info_objeto(driver):
        return None if acta else False

    # 1) Localizar la tabla de Modelo / Marca / Color y su primera fila de datos.
    filas = []
    try:
        headers = driver.find_elements(
            By.XPATH,
            "//*[normalize-space(text())='Modelo' or normalize-space(text())='Marca' or normalize-space(text())='Color']"
        )
        for h in headers:
            if not visible(h):
                continue
            tables = h.find_elements(By.XPATH, "./ancestor::table[1]")
            for table in tables:
                rows = [r for r in table.find_elements(By.XPATH, ".//tr") if visible(r)]
                for row in rows[1:]:
                    txt = norm(row.text)
                    if txt and not ("MODELO" in txt and "MARCA" in txt and "COLOR" in txt):
                        filas.append(row)
                if filas:
                    break
            if filas:
                break
    except Exception:
        pass

    if filas:
        row = filas[0]
        _log("  -> Fila del vehículo detectada; buscando lupa real")

        # Primero: controles/elementos con onclick dentro de la primera celda.
        try:
            celdas = row.find_elements(By.XPATH, "./td")
            celda = celdas[0] if celdas else row
            candidatos = celda.find_elements(
                By.XPATH,
                ".//*[@onclick or @href or self::a or self::button or self::img or self::input or self::span]"
            )
            # incluir la propia celda por si el onclick vive allí
            candidatos.append(celda)
            vistos = set()
            for raw in candidatos:
                try:
                    if raw.id in vistos or not visible(raw):
                        continue
                    vistos.add(raw.id)
                    texto = _probar_y_confirmar(raw, "control primera celda")
                    if texto:
                        return texto
                    # tras un click que no abrió el detalle, volver a ubicar el popup
                    _activar_contexto_info_objeto(driver)
                except Exception:
                    pass
        except Exception:
            pass

        # Segundo: click físico dentro de la primera celda/fila.
        # En la captura la lupa está aproximadamente 15-25 px desde el borde izquierdo.
        try:
            target = (row.find_elements(By.XPATH, "./td") or [row])[0]
            ancho = max(1, int(target.rect.get("width", 40)))
            alto = max(1, int(target.rect.get("height", 25)))
            offsets = [8, 14, 20, 26, 32]
            for x in offsets:
                x = min(x, max(1, ancho - 2))
                try:
                    ActionChains(driver).move_to_element_with_offset(target, x - ancho/2, 0).click().perform()
                except Exception:
                    # alternativa con coordenadas absolutas del centro del elemento
                    try:
                        r = target.rect
                        cx = r.get("x", 0) + x
                        cy = r.get("y", 0) + alto/2
                        driver.execute_script(
                            "const e=document.elementFromPoint(arguments[0]-window.scrollX,arguments[1]-window.scrollY);"
                            "if(e){e.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));"
                            "e.dispatchEvent(new MouseEvent('mouseup',{bubbles:true}));e.click();}",
                            cx, cy
                        )
                    except Exception:
                        continue
                if acta:
                    texto = esperar_datos_descriptivos(driver, acta, timeout=2.2)
                    if texto:
                        _log(f"  -> DATOS DESCRIPTIVOS abiertos (click físico x={x})")
                        return texto
                else:
                    return True
                _activar_contexto_info_objeto(driver)
        except Exception:
            pass

    # 2) Fallback por geometría: imágenes/controles pequeños a la izquierda del modal.
    try:
        _activar_contexto_info_objeto(driver)
        elems = driver.find_elements(By.XPATH, "//*[self::img or self::a or self::button or self::span or @onclick]")
        candidatos = []
        for e in elems:
            if not visible(e):
                continue
            r = e.rect
            w, h = r.get("width", 0), r.get("height", 0)
            if 0 < w <= 60 and 0 < h <= 60:
                candidatos.append((e, r.get("x", 0), r.get("y", 0)))
        candidatos.sort(key=lambda z: (z[1], z[2]))
        for e, x, y in candidatos:
            meta = norm(" ".join([e.get_attribute("src") or "", e.get_attribute("title") or "", e.get_attribute("class") or ""]))
            if "CERR" in meta or "CLOSE" in meta:
                continue
            texto = _probar_y_confirmar(e, f"fallback x={round(x)} y={round(y)}")
            if texto:
                return texto
            _activar_contexto_info_objeto(driver)
    except Exception:
        pass

    return None if acta else False

def _recortar_bloque_datos_descriptivos(texto, acta=""):
    """Devuelve SOLO el bloque del modal final.

    GeneXus deja visible en el DOM el formulario de Consulta de Antecedentes detrás
    del modal. En v8 se leía el body completo y por eso podían capturarse valores
    como Tipo Vehículo=Todos (filtro de búsqueda) para todas las actas.
    """
    texto = str(texto or "")
    lineas = [x.rstrip() for x in texto.splitlines()]
    # Usar la ÚLTIMA aparición del título: el modal final suele estar al final del DOM.
    indices = [i for i, x in enumerate(lineas) if "OBJETOS DATOS DESCRIPTIVOS" in norm(x)]
    if indices:
        lineas = lineas[indices[-1]:]
    bloque = "\n".join(lineas).strip()

    # Si hubiera más de un bloque por ventanas anteriores, conservar el que contiene el acta actual.
    if acta and compact(acta) not in compact(bloque):
        # Buscar desde la última aparición del acta hacia atrás hasta el título más cercano.
        inds_acta = [i for i, x in enumerate(texto.splitlines()) if compact(acta) in compact(x)]
        if inds_acta:
            all_lines = texto.splitlines()
            ia = inds_acta[-1]
            inicio = 0
            for i in range(ia, -1, -1):
                if "OBJETOS DATOS DESCRIPTIVOS" in norm(all_lines[i]):
                    inicio = i
                    break
            bloque = "\n".join(all_lines[inicio:]).strip()
    return bloque


def esperar_datos_descriptivos(driver, acta, timeout=10):
    """Espera el modal final y retorna solamente SU contenido, no el body completo."""
    fin = time.time() + timeout
    while time.time() < fin:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass

        encontrado = _buscar_contexto_con_texto(driver, ["OBJETOS DATOS DESCRIPTIVOS"])
        if encontrado:
            texto_completo = _texto_contexto_actual(driver)
            bloque = _recortar_bloque_datos_descriptivos(texto_completo, acta)
            nt = norm(bloque)
            if compact(acta) in compact(bloque) and "TIPO VEHICULO" in nt and "DOMINIO" in nt:
                return bloque
        time.sleep(0.25)
    return None


def extraer_campo(texto, etiqueta, siguientes):
    """Extrae la ÚLTIMA ocurrencia válida de una etiqueta dentro del modal final."""
    lines = [re.sub(r"\s+", " ", x).strip() for x in str(texto or "").splitlines() if x.strip()]
    et = norm(etiqueta)
    sigs = {norm(s) for s in siguientes}
    invalidos = {"TODOS", "TODAS", "SELECCIONE", "SELECCIONAR", "--", "-"}
    candidatos = []

    for i, line in enumerate(lines):
        nl = norm(line)
        if nl == et or nl.startswith(et + " "):
            # etiqueta + valor en la misma línea
            m = re.match(rf"(?i)^\s*{re.escape(etiqueta)}\s*[:\-]?\s*(.+)$", line)
            if m:
                val = m.group(1).strip()
                if val and norm(val) not in invalidos and norm(val) != et:
                    candidatos.append(val)
                    continue
            # etiqueta y valor en línea siguiente; saltar líneas vacías/etiquetas
            for j in range(i + 1, min(len(lines), i + 6)):
                nv = norm(lines[j])
                if nv in sigs:
                    break
                if not nv or nv in invalidos:
                    continue
                candidatos.append(lines[j])
                break

    return candidatos[-1] if candidatos else ""


def extraer_datos(texto):
    etiquetas = [
        "Acta", "Fecha de Labrado", "Retiene Licencia", "Número de Licencia",
        "Categoria Licencia", "Procedencia Licencia", "Vto. de Licencia",
        "Tipo Vehículo", "Dominio", "Marca", "Color", "Modelo", "Juzgado",
        "COMISARIA", "Usuario", "Lugar"
    ]
    datos = {}
    for e in etiquetas:
        datos[e] = extraer_campo(texto, e, etiquetas)

    # Fallback regex multiline específico para los campos que necesitamos.
    patrones = {
        "Acta": r"(?im)^\s*Acta\s+([^\r\n]+)$",
        "Fecha de Labrado": r"(?im)^\s*Fecha\s+de\s+Labrado\s+([^\r\n]+)$",
        "Tipo Vehículo": r"(?im)^\s*Tipo\s+Veh[ií]culo\s+([^\r\n]+)$",
        "Dominio": r"(?im)^\s*Dominio\s+([^\r\n]+)$",
        "Marca": r"(?im)^\s*Marca\s+([^\r\n]+)$",
        "Color": r"(?im)^\s*Color\s+([^\r\n]+)$",
        "Modelo": r"(?im)^\s*Modelo\s+([^\r\n]+)$",
        "Juzgado": r"(?im)^\s*Juzgado\s+([^\r\n]+)$",
        "COMISARIA": r"(?im)^\s*COMISARIA\s+([^\r\n]+)$",
    }
    for k, pat in patrones.items():
        if not datos.get(k):
            ms = list(re.finditer(pat, texto or ""))
            if ms:
                datos[k] = ms[-1].group(1).strip()

    # Nunca aceptar valores propios de los filtros de la pantalla principal.
    for k in ("Tipo Vehículo", "Marca", "Color", "Modelo", "Dominio"):
        if norm(datos.get(k, "")) in {"TODOS", "TODAS", "SELECCIONE", "SELECCIONAR"}:
            datos[k] = ""
    return datos


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Verificador Vial Caminera v{VERSION}")
        self.root.geometry("920x690")
        self.root.minsize(820, 600)
        self.driver = None
        self.consulta_url = None
        self.stop_flag = False
        self.archivo = tk.StringVar()
        self.solo_pendientes = tk.BooleanVar(value=True)
        self.limite = tk.StringVar(value="1")
        self._ui()

    def _ui(self):
        tk.Label(self.root, text="VERIFICADOR VIAL CAMINERA", font=("Segoe UI", 18, "bold")).pack(pady=(15, 3))
        tk.Label(self.root, text=f"Versión {VERSION} · Verifica TIPO, MARCA/MODELO, COLOR, DOMINIO, INTERVIENE y cruce DOMINIO↔ACTA. Fecha, Motor y Chasis se omiten.", font=("Segoe UI", 10)).pack(pady=(0, 4))
        tk.Label(self.root, text="No guarda usuario ni contraseña. La sesión se inicia manualmente en Chrome.", font=("Segoe UI", 9)).pack(pady=(0, 12))

        f = tk.Frame(self.root)
        f.pack(fill="x", padx=20)
        tk.Entry(f, textvariable=self.archivo, font=("Segoe UI", 10)).pack(side="left", fill="x", expand=True)
        tk.Button(f, text="Elegir Excel", command=self.elegir).pack(side="left", padx=(8, 0))

        opts = tk.Frame(self.root)
        opts.pack(fill="x", padx=20, pady=10)
        tk.Checkbutton(opts, text="Saltar filas ya VERIFICADAS", variable=self.solo_pendientes).pack(side="left")
        tk.Label(opts, text="Máx. filas (0 = todas):").pack(side="left", padx=(25, 0))
        tk.Entry(opts, textvariable=self.limite, width=5).pack(side="left", padx=5)
        tk.Label(opts, text="Para la primera prueba deje 1.", font=("Segoe UI", 9, "italic")).pack(side="left", padx=8)

        botones = tk.Frame(self.root)
        botones.pack(fill="x", padx=20, pady=(0, 10))
        tk.Button(botones, text="1. ABRIR SISTEMA / INICIAR SESIÓN", command=self.abrir_sistema, height=2).pack(side="left", fill="x", expand=True)
        tk.Button(botones, text="2. INICIAR VERIFICACIÓN", command=self.iniciar, height=2).pack(side="left", fill="x", expand=True, padx=8)
        tk.Button(botones, text="DETENER", command=self.detener, height=2).pack(side="left")

        self.estado = tk.Label(self.root, text="Listo.", anchor="w", font=("Segoe UI", 10, "bold"))
        self.estado.pack(fill="x", padx=20, pady=(2, 5))
        self.logbox = ScrolledText(self.root, height=26, font=("Consolas", 9))
        self.logbox.pack(fill="both", expand=True, padx=20, pady=(0, 15))

    def log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.root.after(0, lambda: (self.logbox.insert("end", f"[{ts}] {msg}\n"), self.logbox.see("end")))

    def set_estado(self, msg):
        self.root.after(0, lambda: self.estado.config(text=msg))

    def elegir(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p:
            self.archivo.set(p)
            self.log(f"Excel seleccionado: {p}")

    def abrir_sistema(self):
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
        try:
            opts = webdriver.ChromeOptions()
            opts.add_argument("--start-maximized")
            opts.add_experimental_option("detach", True)
            self.driver = webdriver.Chrome(options=opts)
            self.driver.get(URL_LOGIN)
            self.consulta_url = None
            self.log("Chrome abierto. Inicie sesión normalmente y entre a CONSULTA DE ANTECEDENTES.")
            self.set_estado("Esperando inicio de sesión...")
            messagebox.showinfo("Paso 1", "Inicie sesión en Chrome y abra 'CONSULTA DE ANTECEDENTES'.\n\nDespués vuelva al programa y pulse 'INICIAR VERIFICACIÓN'.")
        except WebDriverException as e:
            messagebox.showerror("Chrome/Selenium", f"No pude abrir Chrome.\n\n{e}")

    def detener(self):
        self.stop_flag = True
        self.log("Se solicitó detener el proceso.")

    def iniciar(self):
        if not self.archivo.get() or not os.path.exists(self.archivo.get()):
            messagebox.showwarning("Excel", "Seleccione primero el archivo Excel.")
            return
        if not self.driver:
            messagebox.showwarning("Sistema", "Primero pulse 'ABRIR SISTEMA / INICIAR SESIÓN'.")
            return
        try:
            if URL_CONSULTA_HINT not in (self.driver.current_url or "").lower():
                messagebox.showwarning("Sistema", "Chrome todavía no está en 'Consulta de Antecedentes'.\n\nAbra esa opción en el sistema y vuelva a intentar.")
                return
            self.consulta_url = self.driver.current_url
        except Exception:
            pass
        self.stop_flag = False
        threading.Thread(target=self.procesar, daemon=True).start()

    def volver_consulta(self):
        if not self.consulta_url:
            return
        self.driver.get(self.consulta_url)
        WebDriverWait(self.driver, 8).until(lambda d: find_input_acta(d) is not None)

    def consultar_acta(self, acta):
        self.volver_consulta()
        inp = find_input_acta(self.driver, self.log)
        if not inp:
            raise RuntimeError("No encontré el campo Nro Acta.")
        # v3: escritura real + eventos del formulario. El sistema GeneXus no siempre
        # acepta un valor si sólo se asigna mediante JavaScript/Selenium.
        cargado = False
        for intento in range(1, 4):
            inp = find_input_acta(self.driver, self.log)
            if not inp:
                return None, "NO SE ENCONTRO CAMPO NRO ACTA"
            try:
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
                inp.click()
                inp.send_keys(Keys.CONTROL, "a")
                inp.send_keys(Keys.BACKSPACE)
                # escribir carácter por carácter para simular teclado humano
                for ch in str(acta):
                    inp.send_keys(ch)
                    time.sleep(0.04)
                inp.send_keys(Keys.TAB)
                time.sleep(0.35)
                valor = (inp.get_attribute("value") or "").strip()
                self.log(f"  Campo Nro Acta intento {intento}: '{valor}'")

                # Si send_keys no dejó el valor, usar el setter nativo del input y
                # disparar los eventos que GeneXus suele escuchar.
                if compact(valor) != compact(acta):
                    try:
                        self.driver.execute_script(
                            """
                            const el = arguments[0], val = arguments[1];
                            const proto = Object.getPrototypeOf(el);
                            const desc = Object.getOwnPropertyDescriptor(proto, 'value');
                            if (desc && desc.set) desc.set.call(el, val); else el.value = val;
                            ['input','change','keyup','blur'].forEach(t =>
                                el.dispatchEvent(new Event(t, {bubbles:true})));
                            """, inp, str(acta)
                        )
                        time.sleep(0.25)
                        valor = (inp.get_attribute("value") or "").strip()
                        self.log(f"  Campo Nro Acta tras eventos JS: '{valor}'")
                    except Exception as ex:
                        self.log(f"  Fallback JS no pudo cargar el acta: {ex}")

                if compact(valor) == compact(acta):
                    cargado = True
                    break
            except Exception as ex:
                self.log(f"  Reintento de carga del acta: {ex}")
            time.sleep(0.5)
        if not cargado:
            try:
                self.log("  Diagnóstico de inputs visibles:")
                for i, e in enumerate(self.driver.find_elements(By.CSS_SELECTOR, "input"), 1):
                    if visible(e):
                        self.log("    #{} type={} id='{}' name='{}' value='{}' x={} y={}".format(
                            i, e.get_attribute("type") or "", e.get_attribute("id") or "",
                            e.get_attribute("name") or "", e.get_attribute("value") or "",
                            round(e.rect.get("x", 0), 1), round(e.rect.get("y", 0), 1)))
            except Exception:
                pass
            return None, "EL SISTEMA NO ACEPTO EL NRO ACTA"

        if not click_buscar(self.driver):
            inp.send_keys(Keys.ENTER)
        time.sleep(0.5)

        # Si el sistema muestra el aviso de campo vacío, NO marcar como no encontrada.
        texto_pagina = norm(self.driver.find_element(By.TAG_NAME, "body").text)
        if "DEBE CARGAR EL N" in texto_pagina and "ACTA" in texto_pagina and "DOMINIO" in texto_pagina:
            return None, "CONSULTA RECHAZADA: EL SISTEMA TOMO NRO ACTA VACIO"

        acta_el = esperar_resultado_acta(self.driver, acta, timeout=12)
        if not acta_el:
            texto_pagina = norm(self.driver.find_element(By.TAG_NAME, "body").text)
            if "DEBE CARGAR EL N" in texto_pagina and "ACTA" in texto_pagina:
                return None, "CONSULTA RECHAZADA: EL SISTEMA TOMO NRO ACTA VACIO"
            return None, "NO ENCONTRADA"

        if not click_tres_puntos(self.driver, acta_el):
            return None, "NO SE PUDO ABRIR LOS TRES PUNTOS"
        if not esperar_info_objeto(self.driver, timeout=10):
            return None, "NO SE DETECTO INFORMACION DE OBJETO"
        self.log("  -> INFORMACION DE OBJETO detectada")

        self.log("  -> Abriendo lupa y confirmando DATOS DESCRIPTIVOS")
        texto = click_lupa(self.driver, acta, self.log)
        if not texto:
            return None, "NO SE PUDO ABRIR DATOS DESCRIPTIVOS DESDE LA LUPA"
        return extraer_datos(texto), None

    def consultar_dominio_contiene_acta(self, dominio, acta):
        if dominio_sin_chapa(dominio):
            return None, None  # no aplica
        self.volver_consulta()
        inp = find_input_dominio(self.driver, self.log)
        if not inp:
            return False, "NO SE ENCONTRO CAMPO DOMINIO"
        if not cargar_input_geneXus(self.driver, inp, dominio):
            return False, "EL SISTEMA NO ACEPTO EL DOMINIO"
        self.log(f"  Verificación cruzada: Dominio {dominio} -> Acta {acta}")
        if not click_buscar(self.driver):
            inp.send_keys(Keys.ENTER)
        time.sleep(0.7)
        try:
            WebDriverWait(self.driver, 10).until(lambda d: compact(acta) in compact(d.find_element(By.TAG_NAME,'body').text) or 'NO SE ENCONTR' in norm(d.find_element(By.TAG_NAME,'body').text))
        except Exception:
            pass
        texto=self.driver.find_element(By.TAG_NAME,'body').text
        return compact(acta) in compact(texto), None

    def procesar(self):
        try:
            limite = int(self.limite.get() or "0")
        except ValueError:
            limite = 0

        p = self.archivo.get()
        carpeta, nombre = os.path.split(p)
        base, ext = os.path.splitext(nombre)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        respaldo = os.path.join(carpeta, f"{base}_RESPALDO_{stamp}{ext}")
        salida = os.path.join(carpeta, f"{base}_VERIFICADO{ext}")
        shutil.copy2(p, respaldo)
        self.log(f"Respaldo creado: {os.path.basename(respaldo)}")

        wb = load_workbook(p)
        ws = wb.active
        total = verificados = diferencias = no_encontrados = errores = 0

        for fila in range(HEADER_ROW + 1, ws.max_row + 1):
            if self.stop_flag:
                break
            acta = ws.cell(fila, COL_SUMARIO).value
            obs = str(ws.cell(fila, COL_OBS).value or "").strip().upper()
            if not acta:
                continue
            if self.solo_pendientes.get() and "VERIFICADO" in obs:
                continue
            if limite and total >= limite:
                break

            total += 1
            acta = str(acta).strip()
            self.set_estado(f"Fila {fila} · Acta {acta}")
            self.log(f"Consultando fila {fila} - Acta {acta}...")

            try:
                datos, problema = self.consultar_acta(acta)
                if problema:
                    if problema == "NO ENCONTRADA":
                        ws.cell(fila, COL_OBS).value = "NO ENCONTRADA EN SISTEMA"
                        no_encontrados += 1
                        self.log("  -> NO ENCONTRADA")
                    else:
                        ws.cell(fila, COL_OBS).value = f"ERROR DE NAVEGACION: {problema}"
                        errores += 1
                        self.log(f"  -> {problema}")
                    wb.save(salida)
                    # Ante un problema de navegación, detener la prueba para no crear falsos resultados.
                    if problema != "NO ENCONTRADA":
                        break
                    continue

                self.log(
                    "  Sistema (modal final): "
                    f"Fecha={datos.get('Fecha de Labrado','')} | "
                    f"Tipo={datos.get('Tipo Vehículo','')} | "
                    f"Dominio={datos.get('Dominio','')} | "
                    f"Marca={datos.get('Marca','')} | Modelo={datos.get('Modelo','')} | "
                    f"Color={datos.get('Color','')} | "
                    f"Juzgado={datos.get('Juzgado','')} | Comisaría={datos.get('COMISARIA','')}"
                )

                web_interviene = " ".join(filter(None, [datos.get("Juzgado", ""), datos.get("COMISARIA", "")]))
                self.log(
                    "  Excel: "
                    f"Tipo={ws.cell(fila, COL_TIPO).value or ''} | "
                    f"Marca/Modelo={ws.cell(fila, COL_MARCA_MODELO).value or ''} | "
                    f"Color={ws.cell(fila, COL_COLOR).value or ''} | "
                    f"Dominio={ws.cell(fila, COL_DOMINIO).value or ''} | "
                    f"Interviene={ws.cell(fila, COL_INTERVIENE).value or ''}"
                )
                fallas = []
                if not coincide_tipo(ws.cell(fila, COL_TIPO).value, datos.get("Tipo Vehículo", "")):
                    fallas.append("TIPO VEHICULO")
                web_marca_modelo = combinar_marca_modelo(datos)
                if not coincide_marca_modelo(ws.cell(fila, COL_MARCA_MODELO).value, web_marca_modelo):
                    fallas.append("MARCA/MODELO")
                if not coincide_color(ws.cell(fila, COL_COLOR).value, datos.get("Color", "")):
                    fallas.append("COLOR")
                if not coincide_dominio(ws.cell(fila, COL_DOMINIO).value, datos.get("Dominio", "")):
                    fallas.append("DOMINIO")
                if not coincide_interviene(ws.cell(fila, COL_INTERVIENE).value, web_interviene):
                    fallas.append("INTERVIENE")

                dominio_excel = ws.cell(fila, COL_DOMINIO).value
                cruce_ok = None
                if not dominio_sin_chapa(dominio_excel):
                    cruce_ok, cruce_error = self.consultar_dominio_contiene_acta(str(dominio_excel).strip(), acta)
                    if cruce_error:
                        self.log(f"  -> Cruce dominio-acta no pudo verificarse: {cruce_error}")
                    elif cruce_ok:
                        self.log("  -> DOMINIO ASOCIA EL ACTA")
                    else:
                        fallas.append("DOMINIO NO ASOCIA ACTA")
                        self.log("  -> DOMINIO NO ASOCIA EL ACTA")

                if fallas:
                    ws.cell(fila, COL_OBS).value = "NO COINCIDE: " + ", ".join(fallas)
                    diferencias += 1
                    self.log("  -> DIFERENCIAS: " + ", ".join(fallas))
                else:
                    if cruce_ok is True:
                        ws.cell(fila, COL_OBS).value = "VERIFICADO POR ACTA Y DOMINIO"
                        self.log("  -> VERIFICADO POR ACTA Y DOMINIO")
                    else:
                        ws.cell(fila, COL_OBS).value = "VERIFICADO"
                        self.log("  -> VERIFICADO")
                    verificados += 1

                wb.save(salida)
                # Reinicia la pantalla antes de la siguiente acta.
                try:
                    self.volver_consulta()
                except Exception:
                    pass

            except Exception as e:
                errores += 1
                ws.cell(fila, COL_OBS).value = f"ERROR DE CONSULTA: {str(e)[:120]}"
                wb.save(salida)
                self.log(f"  -> ERROR: {e}")
                break

        wb.save(salida)
        self.set_estado("Proceso finalizado.")
        resumen = (
            f"Finalizado. Procesadas: {total} | Verificadas: {verificados} | "
            f"Con diferencias: {diferencias} | No encontradas: {no_encontrados} | Errores: {errores}"
        )
        self.log(resumen)
        self.log(f"Archivo de salida: {salida}")
        self.root.after(0, lambda: messagebox.showinfo("Verificación finalizada", resumen + f"\n\nSalida:\n{salida}"))


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()

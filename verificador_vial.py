import os
import re
import time
import shutil
import threading
import unicodedata
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

URL_LOGIN = "https://sistemas.seguridad.mendoza.gov.ar/vialcaminera//servlet/com.ktksuitelr.mdlsgt.hlogin2"
URL_CONSULTA_HINT = "wpconsultaantecedentes"

# Columnas del Excel actual
COL_SUMARIO = 11      # K
COL_OBS = 13          # M
HEADER_ROW = 3

# Campos a contrastar. Se comparan solo si el valor de Excel es utilizable.
CAMPOS = {
    3: "TIPO VEHICULO",   # C
    4: "MARCA/MODELO",    # D
    5: "COLOR",           # E
    6: "DOMINIO",         # F
    7: "MOTOR",           # G
    8: "CHASIS",          # H
    9: "MOTIVO",          # I
    10: "INTERVIENE",     # J
}

IGNORAR_EXCEL = {
    "", "S/D", "SD", "SIN DATO", "SIN DATOS", "NO SE DIVISA",
    "NO SE OBSERVA", "NO POSEE", "SIN DOMINIO", "S/CHAPA", "S/ CHAPA"
}


def norm(s):
    if s is None:
        return ""
    s = str(s).upper().strip()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    # Normalización fuerte para números alfanuméricos, dominios, motor/chasis.
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def compact(s):
    return re.sub(r"[^A-Z0-9]", "", norm(s))


def valor_ignorable(v):
    n = norm(v)
    if n in IGNORAR_EXCEL:
        return True
    if n.startswith("S CHAPA") and len(compact(v)) <= 6:
        return True
    return False


def palabras_significativas(v):
    stop = {"CC", "C", "DE", "DEL", "LA", "EL", "Y", "LEY", "NRO", "N"}
    return [p for p in norm(v).split() if len(p) >= 3 and p not in stop]


def coincide(valor_excel, texto_web, campo):
    if valor_ignorable(valor_excel):
        return True

    ve = norm(valor_excel)
    wc = compact(texto_web)
    vc = compact(valor_excel)

    # Identificadores exactos: dominio, motor, chasis.
    if campo in {"DOMINIO", "MOTOR", "CHASIS"}:
        # Si el Excel dice S/CHAPA pero contiene un dominio entre paréntesis, extraerlo.
        if campo == "DOMINIO":
            candidatos = re.findall(r"\b[A-Z]{2,3}\d{3}[A-Z]{0,2}\b|\b[A-Z]\d{3}[A-Z]{3}\b|\b\d{3}[A-Z]{3}\b", ve)
            if candidatos:
                return any(compact(x) in wc for x in candidatos)
        return bool(vc) and vc in wc

    # Para marca/modelo, tipo, color, motivo, dependencia: basta que la mayoría
    # de palabras significativas figure en el detalle visible.
    toks = palabras_significativas(valor_excel)
    if not toks:
        return True
    presentes = sum(1 for t in toks if compact(t) in wc)
    umbral = 1 if len(toks) <= 2 else max(2, int(len(toks) * 0.65 + 0.49))
    return presentes >= umbral


def find_input_by_label(driver, label_text):
    # 1) Buscar label y un input cercano.
    xpath = (
        f"//*[contains(translate(normalize-space(.), 'abcdefghijklmnopqrstuvwxyzáéíóúñ', "
        f"'ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÑ'), '{label_text.upper()}')]"
    )
    elems = driver.find_elements(By.XPATH, xpath)
    for el in elems:
        try:
            # input dentro del mismo contenedor o inmediatamente cercano
            for xp in [".//input", "./following::input[1]", "../following-sibling::*//input[1]", "../input[1]"]:
                cand = el.find_elements(By.XPATH, xp)
                if cand and cand[0].is_displayed() and cand[0].is_enabled():
                    return cand[0]
        except Exception:
            pass

    # 2) Heurísticas por atributos.
    for css in [
        "input[name*='ACTA' i]", "input[id*='ACTA' i]",
        "input[name*='NRO' i]", "input[id*='NRO' i]",
    ]:
        try:
            cands = driver.find_elements(By.CSS_SELECTOR, css)
            for c in cands:
                if c.is_displayed() and c.is_enabled():
                    return c
        except Exception:
            pass
    return None


def click_buscar(driver):
    candidatos = driver.find_elements(By.XPATH,
        "//button[contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'BUSCAR')]"
        "|//input[@type='button' or @type='submit'][contains(translate(@value,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'BUSCAR')]"
        "|//a[contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'BUSCAR')]"
    )
    for e in candidatos:
        if e.is_displayed() and e.is_enabled():
            driver.execute_script("arguments[0].click();", e)
            return True
    return False


def abrir_detalle_si_existe(driver):
    # Intenta el botón de tres puntos / detalle visible de la primera fila de resultados.
    xpaths = [
        "//*[contains(@title,'Detalle') or contains(@title,'detalle')][1]",
        "//button[contains(normalize-space(.),'...')][1]",
        "//a[contains(normalize-space(.),'...')][1]",
        "//table//tbody//tr[1]//*[self::button or self::a][last()]",
    ]
    for xp in xpaths:
        try:
            elems = driver.find_elements(By.XPATH, xp)
            for e in elems:
                if e.is_displayed() and e.is_enabled():
                    before = driver.current_url
                    driver.execute_script("arguments[0].click();", e)
                    time.sleep(0.8)
                    return before
        except Exception:
            pass
    return None


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Verificador Vial Caminera")
        self.root.geometry("860x650")
        self.root.minsize(760, 560)
        self.driver = None
        self.stop_flag = False
        self.archivo = tk.StringVar()
        self.solo_pendientes = tk.BooleanVar(value=True)
        self.abrir_detalle = tk.BooleanVar(value=True)
        self.limite = tk.StringVar(value="1")
        self._ui()

    def _ui(self):
        tk.Label(self.root, text="VERIFICADOR VIAL CAMINERA", font=("Segoe UI", 18, "bold")).pack(pady=(15, 4))
        tk.Label(self.root, text="Compara cada fila del Excel con el acta consultada en el sistema. No guarda usuario ni contraseña.", font=("Segoe UI", 10)).pack(pady=(0, 12))

        f = tk.Frame(self.root)
        f.pack(fill="x", padx=20)
        tk.Entry(f, textvariable=self.archivo, font=("Segoe UI", 10)).pack(side="left", fill="x", expand=True)
        tk.Button(f, text="Elegir Excel", command=self.elegir).pack(side="left", padx=(8, 0))

        opts = tk.Frame(self.root)
        opts.pack(fill="x", padx=20, pady=10)
        tk.Checkbutton(opts, text="Saltar filas ya VERIFICADAS", variable=self.solo_pendientes).pack(side="left")
        tk.Checkbutton(opts, text="Intentar abrir detalle del acta", variable=self.abrir_detalle).pack(side="left", padx=15)
        tk.Label(opts, text="Máx. filas (0 = todas):").pack(side="left")
        tk.Entry(opts, textvariable=self.limite, width=5).pack(side="left", padx=5)

        botones = tk.Frame(self.root)
        botones.pack(fill="x", padx=20, pady=(0, 10))
        tk.Button(botones, text="1. ABRIR SISTEMA / INICIAR SESIÓN", command=self.abrir_sistema, height=2).pack(side="left", fill="x", expand=True)
        tk.Button(botones, text="2. INICIAR VERIFICACIÓN", command=self.iniciar, height=2).pack(side="left", fill="x", expand=True, padx=8)
        tk.Button(botones, text="DETENER", command=self.detener, height=2).pack(side="left")

        self.estado = tk.Label(self.root, text="Listo.", anchor="w", font=("Segoe UI", 10, "bold"))
        self.estado.pack(fill="x", padx=20, pady=(2, 5))
        self.logbox = ScrolledText(self.root, height=25, font=("Consolas", 9))
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
            self.log("Chrome abierto. Inicie sesión normalmente y entre a CONSULTA DE ANTECEDENTES.")
            self.set_estado("Esperando que inicie sesión y abra Consulta de Antecedentes...")
            messagebox.showinfo("Paso 1", "Inicie sesión en la ventana de Chrome y abra 'CONSULTA DE ANTECEDENTES'.\n\nDespués vuelva aquí y pulse 'INICIAR VERIFICACIÓN'.")
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
        self.stop_flag = False
        threading.Thread(target=self.procesar, daemon=True).start()

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
        total = 0
        verificados = 0
        diferencias = 0
        no_encontrados = 0
        errores = 0

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
                inp = find_input_by_label(self.driver, "NRO ACTA") or find_input_by_label(self.driver, "NRO. ACTA")
                if not inp:
                    raise RuntimeError("No encontré el campo 'Nro Acta'. Verifique que esté en Consulta de Antecedentes.")

                inp.click()
                inp.send_keys(Keys.CONTROL, "a")
                inp.send_keys(acta)
                if not click_buscar(self.driver):
                    inp.send_keys(Keys.ENTER)
                time.sleep(1.2)

                body = self.driver.find_element(By.TAG_NAME, "body").text
                body_n = compact(body)
                acta_c = compact(acta)
                if acta_c not in body_n:
                    ws.cell(fila, COL_OBS).value = "NO ENCONTRADA EN SISTEMA"
                    no_encontrados += 1
                    self.log("  -> NO ENCONTRADA")
                    wb.save(salida)
                    continue

                url_antes_detalle = None
                if self.abrir_detalle.get():
                    url_antes_detalle = abrir_detalle_si_existe(self.driver)
                    if url_antes_detalle:
                        try:
                            WebDriverWait(self.driver, 4).until(lambda d: len(d.find_element(By.TAG_NAME, "body").text) > 50)
                        except Exception:
                            pass
                        body = self.driver.find_element(By.TAG_NAME, "body").text

                fallas = []
                for col, campo in CAMPOS.items():
                    val = ws.cell(fila, col).value
                    if not coincide(val, body, campo):
                        fallas.append(campo)

                if fallas:
                    ws.cell(fila, COL_OBS).value = "NO COINCIDE: " + ", ".join(fallas)
                    diferencias += 1
                    self.log("  -> DIFERENCIAS: " + ", ".join(fallas))
                else:
                    ws.cell(fila, COL_OBS).value = "VERIFICADO"
                    verificados += 1
                    self.log("  -> VERIFICADO")

                wb.save(salida)  # guardar progreso tras cada acta

                # Si el detalle navegó a otra pantalla, volver a consulta.
                if url_antes_detalle:
                    try:
                        self.driver.back()
                        time.sleep(0.8)
                    except Exception:
                        pass

            except Exception as e:
                errores += 1
                ws.cell(fila, COL_OBS).value = f"ERROR DE CONSULTA: {str(e)[:120]}"
                wb.save(salida)
                self.log(f"  -> ERROR: {e}")
                # Ante un error de pantalla, no seguir en masa para evitar falsos resultados.
                if "Nro Acta" in str(e) or "NRO ACTA" in str(e).upper():
                    break

        wb.save(salida)
        self.set_estado("Proceso finalizado.")
        resumen = (f"Finalizado. Procesadas: {total} | Verificadas: {verificados} | "
                   f"Con diferencias: {diferencias} | No encontradas: {no_encontrados} | Errores: {errores}")
        self.log(resumen)
        self.log(f"Archivo de salida: {salida}")
        self.root.after(0, lambda: messagebox.showinfo("Verificación finalizada", resumen + f"\n\nSalida:\n{salida}"))


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()

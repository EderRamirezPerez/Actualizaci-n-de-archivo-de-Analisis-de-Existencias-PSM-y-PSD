# -*- coding: utf-8 -*-
"""
Descarga diaria de reportes Oracle BI, los mueve a una carpeta local
y limpia duplicados.  Necesita la variable de entorno ORACLE_KEY
con tu contraseña antes de ejecutarse (se define en GitHub Secrets).
"""

import os
import time
import shutil
import csv
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 📂 Carpetas locales (relativas al repo)
DOWNLOAD_DIR = "downloads"                                  # donde Chrome descargará los CSV
DESTINO_DIR  = os.path.join("outputs", "Reporte Existencia")  # destino final

# Crea las carpetas si no existen
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(DESTINO_DIR,  exist_ok=True)

MAX_INTENTOS_DEFAULT     = 3
TIMEOUT_DESCARGA_DEFAULT = 40  # seg

# Configurar Chrome
prefs = {
    "download.default_directory": os.path.abspath(DOWNLOAD_DIR),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
chrome_options = Options()
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver        = webdriver.Chrome(service=Service(), options=chrome_options)
wait_global   = WebDriverWait(driver, 30)

# ──────────────────────────────────────────────────────────────────────────────
def _descarga_exitosa(antes, timeout):
    inicio = time.time()
    while time.time() - inicio < timeout:
        nuevos       = {f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.csv')} - antes
        en_progreso  = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.crdownload')]
        if nuevos and not en_progreso:
            return True
        time.sleep(2)
    return False


def _limpiar_descargas_parciales():
    for f in os.listdir(DOWNLOAD_DIR):
        if f.endswith('.crdownload'):
            try:
                os.remove(os.path.join(DOWNLOAD_DIR, f))
            except OSError:
                pass


def iniciar_sesion():
    driver.get("https://eisq.fa.us6.oraclecloud.com/analytics/")
    time.sleep(5)
    wait_global.until(EC.element_to_be_clickable((By.ID, "Languages"))).click()
    wait_global.until(EC.element_to_be_clickable((By.XPATH, "//option[@value='es-es']"))).click()
    time.sleep(2)
    driver.execute_script("callLanguageChange();")
    time.sleep(5)
    print("✅ Idioma establecido en español")

    driver.find_element(By.NAME, "userid").send_keys("Eder.Ramirez")
    driver.find_element(By.NAME, "password").send_keys(os.environ["ORACLE_KEY"])
    driver.find_element(By.NAME, "password").send_keys(Keys.RETURN)
    time.sleep(10)
    print("✅ Sesión iniciada")


def descargar_multiples_reportes():
    urls = {
        "ASN": "https://eisq.fa.us6.oraclecloud.com/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2FCustom%2FPacksys%2FCompras%2FAnalisis%2FASN",
        "Catalogo de Productos": "https://eisq.fa.us6.oraclecloud.com/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2FCustom%2FPacksys%2FCompras%2FAnalisis%2FCatalogo%20de%20Productos",
        "Detalle de Ordenes de Venta": "https://eisq.fa.us6.oraclecloud.com/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2FCustom%2FPacksys%2FCompras%2FAnalisis%2FDetalle%20de%20Ordenes%20de%20Venta",
        "Existencias Inventario Disponible Localizador": "https://eisq.fa.us6.oraclecloud.com/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2FCustom%2FPacksys%2FCompras%2FAnalisis%2FExistencias%20Inventario%20Disponible%20Localizador",
        "No. Cliente y No. Sitio": "https://eisq.fa.us6.oraclecloud.com/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2FCustom%2FPacksys%2FCompras%2FAnalisis%2FNo.%20Cliente%20y%20No.%20Sitio",
        "Ordenes de Compra Abiertas": "https://eisq.fa.us6.oraclecloud.com/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2FCustom%2FPacksys%2FCompras%2FAnalisis%2FOrdenes%20de%20Compra%20Abiertas",
        "Orden de Transferencia": "https://eisq.fa.us6.oraclecloud.com/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2FCustom%2FPacksys%2FCompras%2FEnvios%2FOrden%20de%20Transferencia",
        "Orden de Transferencia maquinas": "https://eisq.fa.us6.oraclecloud.com/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2FCustom%2FPacksys%2FCompras%2FAnalisis%2FOrden%20de%20Transferencia%20maquinas",
        "Existencia maquinas": "https://eisq.fa.us6.oraclecloud.com/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2FCustom%2FPacksys%2FCompras%2FAnalisis%2FExistencia%20maquinas",
        "Reporte de Transacciones de Venta": "https://eisq.fa.us6.oraclecloud.com/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2FCustom%2FPacksys%2FCompras%2FAnalisis%2FReporte%20de%20Transacciones%20de%20Venta",
    }

    fallidos = []

    for nombre, url in urls.items():
        intentos_max   = 5 if nombre == "Detalle de Ordenes de Venta" else MAX_INTENTOS_DEFAULT
        timeout_archiv = 80 if nombre == "Detalle de Ordenes de Venta" else TIMEOUT_DESCARGA_DEFAULT
        wait_local     = WebDriverWait(driver, 60) if nombre == "Detalle de Ordenes de Venta" else wait_global

        print(f"\n📄 {nombre} → intentos: {intentos_max}, timeout: {timeout_archiv}s")
        exito = False

        for intento in range(1, intentos_max + 1):
            print(f"   ▶ Intento {intento}/{intentos_max}")
            prev_csv = {f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.csv')}
            try:
                driver.execute_script(f"window.open('{url}', '_blank');")
                time.sleep(2)
                driver.switch_to.window(driver.window_handles[-1])

                wait_local.until(EC.presence_of_element_located((By.CLASS_NAME, "ResultLinksCell")))
                ActionChains(driver).move_to_element(driver.find_element(By.CLASS_NAME, "ResultLinksCell")).perform()

                def click(sel):
                    ActionChains(driver).move_to_element(wait_local.until(EC.element_to_be_clickable(sel))).click().perform()

                click((By.LINK_TEXT, "Exportar"))
                click((By.XPATH, "//td[contains(text(),'Datos')]"))
                click((By.XPATH, "//td[contains(text(),'CSV')]"))
                print("      ⏬ descarga solicitada")

                if _descarga_exitosa(prev_csv, timeout_archiv):
                    exito = True
                    break

            except Exception as e:
                print(f"      ⚠️ Error intento {intento}: {e}")
            finally:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                _limpiar_descargas_parciales()

        if exito:
            print(f"✅ {nombre} listo")
        else:
            print(f"❌ {nombre} NO se descargó")
            fallidos.append(nombre)

    if fallidos:
        with open("errores_descarga.csv", "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows([[e] for e in fallidos])
        print("⚠️ Ver 'errores_descarga.csv'")


def mover_a_directorio():
    print("\n🔄 Moviendo CSV a carpeta destino…")
    for f in [x for x in os.listdir(DOWNLOAD_DIR) if x.endswith('.csv')]:
        shutil.move(os.path.join(DOWNLOAD_DIR, f), os.path.join(DESTINO_DIR, f))
        print(f"   📁 {f}")


def limpiar_duplicados():
    print("\n🧹 Limpiando duplicados…")
    bases = [
        "Detalle de Ordenes de Venta",
        "Orden de Transferencia maquinas",
        "Reporte de Transacciones de Venta",
        "Existencias Inventario Disponible Localizador",
        "Existencia maquinas",
        "No. Cliente y No. Sitio",
        "Ordenes de Compra Abiertas",
        "Catalogo de Productos",
        "ASN",
        "Orden de Transferencia",
    ]
    patron = re.compile(r"^(?P<base>.+?) \((?P<num>\d+)\)\.csv$")

    for base in bases:
        original  = f"{base}.csv"
        archivos  = [f for f in os.listdir(DESTINO_DIR) if f.startswith(base) and f.endswith('.csv')]
        if not archivos:
            continue

        if original in archivos:
            for f in archivos:
                if f != original and patron.match(f):
                    os.remove(os.path.join(DESTINO_DIR, f))
                    print(f"   🗑️ {f}")
        else:
            duplicados = sorted(
                [f for f in archivos if patron.match(f)],
                key=lambda x: int(patron.match(x).group("num")),
            )
            conservar = duplicados.pop(0)
            os.rename(os.path.join(DESTINO_DIR, conservar), os.path.join(DESTINO_DIR, original))
            print(f"   🔄 {conservar} ➜ {original}")
            for f in duplicados:
                os.remove(os.path.join(DESTINO_DIR, f))
                print(f"   🗑️ {f}")
    print("✅ Limpieza terminada")


def ejecutar_proceso():
    iniciar_sesion()
    descargar_multiples_reportes()
    mover_a_directorio()
    limpiar_duplicados()


if __name__ == "__main__":
    try:
        ejecutar_proceso()
    finally:
        driver.quit()

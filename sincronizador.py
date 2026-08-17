# sincronizador.py
import time
import os
import glob
import pandas as pd
import gspread
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials

print("🤖 Iniciando Robot Sincronizador Multilínea en la Nube...")

# ==========================================
# CONFIGURACIÓN GENERAL
# ==========================================
ruta_actual = os.getcwd()
ID_CARPETA_PADRE_FOTOS = '1QG5vmof76jSefx9RmMLtLMkaxi_OsL2L'

archivo_oro = os.path.join(ruta_actual, 'lista_productos_oro.xlsx')
archivo_plata = os.path.join(ruta_actual, 'lista_productos_plata.xlsx')

archivos_viejos = glob.glob(os.path.join(ruta_actual, 'lista_productos*.xlsx'))
for archivo in archivos_viejos:
    try:
        if os.path.exists(archivo):
            os.remove(archivo)
    except:
        pass

# ==========================================
# FUNCIÓN SECUNDARIA: MAPEAR FOTOS DE DRIVE
# ==========================================
def mapear_fotos_de_drive():
    print("🔍 Escaneando carpetas de fotos en Google Drive...")
    scope = ['https://www.googleapis.com/auth/drive.readonly']
    creds = Credentials.from_service_account_file('credenciales.json', scopes=scope)
    service = build('drive', 'v3', credentials=creds)

    carpetas = []
    page_token = None

    while True:
        query_carpetas = f"'{ID_CARPETA_PADRE_FOTOS}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        resultados = service.files().list(
            q=query_carpetas,
            fields="nextPageToken, files(id, name)",
            pageSize=1000,
            pageToken=page_token
        ).execute()

        carpetas.extend(resultados.get('files', []))
        page_token = resultados.get('nextPageToken')
        if not page_token:
            break

    diccionario_fotos = {}

    for carp in carpetas:
        id_carpeta = carp['id']
        codigo_referencia = str(carp['name']).strip()

        query_fotos = f"'{id_carpeta}' in parents and mimeType contains 'image/' and trashed = false"
        resultados_fotos = service.files().list(
            q=query_fotos,
            fields="files(id, name)",
            pageSize=1000
        ).execute()

        fotos = resultados_fotos.get('files', [])

        foto_perfil = []
        otras_fotos = []

        for foto in fotos:
            enlace = f"https://drive.google.com/file/d/{foto['id']}/view"
            nombre_archivo = str(foto.get('name', '')).upper()

            if 'PERFIL' in nombre_archivo:
                foto_perfil.append(enlace)
            else:
                otras_fotos.append(enlace)

        enlaces_ordenados = foto_perfil + otras_fotos

        if enlaces_ordenados:
            diccionario_fotos[codigo_referencia] = ";".join(enlaces_ordenados)

    return diccionario_fotos

mapa_imagenes = {}
try:
    mapa_imagenes = mapear_fotos_de_drive()
except Exception as e:
    print(f"⚠️ No se pudieron cargar las fotos: {e}")

# ==========================================
# FASE 2: DESCARGAR EXCELES CON SELENIUM (MODO FANTASMA)
# ==========================================
opciones = Options()
opciones.add_argument('--headless=new')
opciones.add_argument('--no-sandbox')
opciones.add_argument('--disable-dev-shm-usage')
# 🔥 SOLUCIÓN 1: Le damos un monitor gigante al robot para que nada se amontone
opciones.add_argument('--window-size=1920,1080')

prefs = {"download.default_directory": ruta_actual, "download.prompt_for_download": False}
opciones.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opciones)

def esperar_y_renombrar(nombre_destino):
    print("⏳ Procesando descarga...")
    archivo_descargado = None
    while True:
        if glob.glob(os.path.join(ruta_actual, '*.crdownload')):
            time.sleep(1)
            continue
        archivos = [f for f in glob.glob(os.path.join(ruta_actual, 'lista_productos*.xlsx'))
                    if f != archivo_oro and f != archivo_plata]
        if archivos:
            archivo_descargado = archivos[0]
            if os.path.getsize(archivo_descargado) > 0:
                break
        time.sleep(1)

    time.sleep(2)
    if os.path.exists(nombre_destino):
        os.remove(nombre_destino)
    os.rename(archivo_descargado, nombre_destino)

try:
    print("🔐 Iniciando sesión en Belatriz...")
    driver.get('https://belatrizcolombia.com/app/public/login')
    time.sleep(3)

    usuario = os.environ.get('BELATRIZ_USER')
    clave = os.environ.get('BELATRIZ_PASS')

    driver.find_element(By.ID, 'Documento').send_keys(usuario)
    driver.find_element(By.ID, 'password').send_keys(clave)

    # 🔥 SOLUCIÓN 2: Clic forzado inyectado con JS (Ignora si algo estorba en pantalla)
    boton_login = driver.find_element(By.ID, 'login-text')
    driver.execute_script("arguments[0].click();", boton_login)
    time.sleep(5)

    print("📥 Descargando catálogo de Oro Laminado 18K...")
    driver.get('https://belatrizcolombia.com/app/public/biblioteca')
    time.sleep(3)

    # Clic inyectado para el botón exportar oro
    boton_exportar_oro = driver.find_element(By.ID, 'btnExportar')
    driver.execute_script("arguments[0].click();", boton_exportar_oro)
    esperar_y_renombrar(archivo_oro)

    print("🔄 Cambiando a la línea Silver...")
    driver.get('https://belatrizcolombia.com/app/public/home')
    time.sleep(3)
    driver.execute_script("cambioLineaProducto(2);")
    time.sleep(5)

    print("📥 Descargando catálogo de Silver...")
    driver.get('https://belatrizcolombia.com/app/public/biblioteca')
    time.sleep(3)

    # Clic inyectado para el botón exportar plata
    boton_exportar_plata = driver.find_element(By.ID, 'btnExportar')
    driver.execute_script("arguments[0].click();", boton_exportar_plata)
    esperar_y_renombrar(archivo_plata)

finally:
    driver.quit()

# ==========================================
# FASE 3: TRANSFORMAR Y FUSIONAR DATOS
# ==========================================
print("🧠 Fusionando catálogos...")

df_oro = pd.read_excel(archivo_oro, skiprows=4)
df_plata = pd.read_excel(archivo_plata, skiprows=4)

df_proveedor = pd.concat([df_oro, df_plata], ignore_index=True)
df_proveedor.columns = df_proveedor.columns.str.strip()

def limpiar_y_combinar_nombre(row):
    cat_original = str(row['CATEGORIA']).strip().upper()
    nombre_original = str(row['NOMBRE']).strip()
    cat_map = {
        'ARETES': ('Aretes', 'Arete'), 'DIJES': ('Dijes', 'Dije'),
        'CADENAS': ('Cadenas', 'Cadena'), 'PULSERAS': ('Pulseras', 'Pulsera'),
        'ANILLOS': ('Anillos', 'Anillo'), 'HERRAJES': ('Herrajes', 'Herraje'),
        'EXCLUSIVIDADES': ('Exclusividades', 'Exclusividad'),
        'DIFERENCIAL': ('Diferencial', 'Diferencial'),
        'GARGANTILLA': ('Gargantillas', 'Gargantilla'), 'OTROS': ('Otros', 'Otros')
    }
    cat_plural, cat_singular = cat_map.get(cat_original, (cat_original.title(), cat_original.title()))
    if not nombre_original: return cat_singular
    palabras = nombre_original.split()
    if not palabras: return cat_singular
    primer_elemento = palabras[0]
    if len(primer_elemento) == 1 and primer_elemento.isalpha():
        nombre_limpio = " ".join(palabras[1:])
    elif primer_elemento.lower() in [cat_plural.lower(), cat_singular.lower(), cat_original.lower()]:
        nombre_limpio = " ".join(palabras[1:])
    else:
        nombre_limpio = nombre_original
    return f"{cat_singular} {nombre_limpio}".strip()

def extraer_variantes(referencia):
    if '-' not in referencia: return referencia, ''
    partes = referencia.split('-')
    if len(partes) >= 2 and partes[0].isalpha() and len(partes[0]) <= 2:
        codigo_base = f"{partes[0]}-{partes[1]}".strip()
        indice_variante = 2
    else:
        codigo_base = partes[0].strip()
        indice_variante = 1

    variante = partes[indice_variante].strip().upper() if len(partes) > indice_variante else ''
    if not variante: return codigo_base, ''
    if len(variante) == 1: return referencia, ''
    if variante.startswith('T') and variante[1:].replace('.', '').isdigit(): return referencia, ''
    return codigo_base, variante

df_final = pd.DataFrame()
df_final['id'] = df_proveedor.index + 1

referencias_limpias = df_proveedor['REFERENCIA'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
bases_y_colores = referencias_limpias.apply(extraer_variantes)

df_final['code'] = referencias_limpias
df_final['base_code'] = [x[0] for x in bases_y_colores]
df_final['name'] = df_proveedor.apply(limpiar_y_combinar_nombre, axis=1)
df_final['collection'] = ''
df_final['category'] = df_proveedor['CATEGORIA'].str.strip().str.title()
df_final['material'] = df_proveedor['LINEA PRODUCTO']
df_final['color'] = [x[1] for x in bases_y_colores]
df_final['price'] = df_proveedor['PRECIO DETAL'] + 10000
df_final['previousPrice'] = ''
df_final['stock'] = df_proveedor['ESTATUS'].apply(lambda x: 10 if str(x).strip().lower() == 'activo' else 0)
df_final['status'] = df_final['stock'].apply(lambda x: 'Normal' if x > 0 else 'Agotado')
df_final['images'] = df_final['code'].map(mapa_imagenes).fillna(df_final['base_code'].map(mapa_imagenes)).fillna('')
df_final['description'] = 'Hermosa pieza de ' + df_final['material']
df_final['tags'] = ''

df_final = df_final.fillna('')
df_final = df_final[df_final['stock'] > 0]
df_final['id'] = range(1, len(df_final) + 1)

# ==========================================
# FASE 4: SINCRONIZACIÓN CON LA NUBE
# ==========================================
print("☁️ Subiendo catálogo a Google Sheets...")
try:
    gc = gspread.service_account(filename='credenciales.json')
    sh = gc.open_by_key('1jRTJqwPir1CSGQzKazjgWNo3TWr3b8Nz-RkWHOfatxo')
    worksheet = sh.sheet1
    worksheet.clear()
    datos_a_subir = [df_final.columns.values.tolist()] + df_final.values.tolist()
    worksheet.update(values=datos_a_subir, range_name='A1')
    print("🎉 ¡Sincronización completada con éxito!")
except Exception as e:
    print(f"❌ Error al subir a Sheets: {e}")
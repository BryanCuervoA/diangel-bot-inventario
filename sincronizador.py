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
# FASE 3.5: INYECTAR PRODUCTOS PERSONALIZADOS
# ==========================================
print("🔗 Inyectando catálogo de Personalizados...")
try:
    # 🔥 AQUÍ ESTABA EL ERROR: Faltaba esta línea para iniciar sesión en Google
    gc = gspread.service_account(filename='credenciales.json')
    sh = gc.open_by_key('1jRTJqwPir1CSGQzKazjgWNo3TWr3b8Nz-RkWHOfatxo')

    worksheet_pers = sh.worksheet('Personalizados')
    datos_pers = worksheet_pers.get_all_records()

    if datos_pers:
        df_pers = pd.DataFrame(datos_pers)

        # 1. Emparejar columnas exactamente
        for col in df_final.columns:
            if col not in df_pers.columns:
                df_pers[col] = ''

        # 2. El Super Cerebro de Precios
        def limpiar_precio(p):
            if pd.isna(p) or str(p).strip() == '': return 0
            if isinstance(p, (int, float)):
                if 0 < p < 1000: return int(p * 1000)
                return int(p)
            p_str = str(p).replace('$', '').replace('COP', '').replace('.', '').replace(',', '').strip()
            try: return int(p_str)
            except: return 0

        df_pers['price'] = df_pers['price'].apply(limpiar_precio)

        # 3. Rellenar vacíos por si acaso
        df_pers['status'] = df_pers['status'].apply(lambda x: 'Normal' if str(x).strip() == '' else x)
        df_pers['description'] = df_pers['description'].apply(lambda x: 'Hermosa pieza exclusiva' if str(x).strip() == '' else x)

        # 4. Enlaces de imágenes (respeta tu link manual o busca en Drive)
        def asignar_imagen(row):
            img_actual = str(row.get('images', '')).strip()
            if img_actual: return img_actual
            codigo = str(row.get('code', ''))
            base = str(row.get('base_code', ''))
            if codigo in mapa_imagenes: return mapa_imagenes[codigo]
            if base in mapa_imagenes: return mapa_imagenes[base]
            return ''

        df_pers['images'] = df_pers.apply(asignar_imagen, axis=1)

        # 5. Filtrar los que tienen stock
        df_pers['stock'] = pd.to_numeric(df_pers['stock'], errors='coerce').fillna(0)
        df_pers = df_pers[df_pers['stock'] > 0]

        # 6. Unir al catálogo de Belatriz y renumerar IDs ordenadamente
        df_final = pd.concat([df_final, df_pers], ignore_index=True)
        df_final['id'] = range(1, len(df_final) + 1)

        print(f"✅ ¡{len(df_pers)} productos personalizados procesados!")
except Exception as e:
    print(f"⚠️ Omitiendo personalizados (Pestaña vacía o error): {e}")

# ==========================================
# FASE 4: SUBIR A HOJA 1 (PÁGINA WEB)
# ==========================================
print("⬆️ Subiendo catálogo unificado a Hoja 1...")
try:
    worksheet_hoja1 = sh.worksheet('Hoja 1')
    worksheet_hoja1.clear()
    # fillna('') es vital para que Google Sheets no arroje error con celdas vacías
    datos_hoja1 = [df_final.columns.values.tolist()] + df_final.fillna('').values.tolist()
    worksheet_hoja1.update(values=datos_hoja1, range_name='A1')
    print("✅ 'Hoja 1' actualizada con éxito. ¡Lista para la página web!")
except Exception as e:
    print(f"❌ Error al subir a Hoja 1: {e}")

# ==========================================
# FASE 5: GENERAR CATÁLOGO PARA META Y WHATSAPP
# ==========================================
print("🛒 Generando feed estructurado para Meta Business...")
df_meta = pd.DataFrame()

df_meta['id'] = df_final['id']
df_meta['title'] = df_final['name'].astype(str).str[:200]
df_meta['description'] = df_final['description'].astype(str).str[:9999]
df_meta['availability'] = df_final['stock'].apply(lambda x: 'in stock' if x > 0 else 'out of stock')
df_meta['condition'] = 'new'
# Aquí Meta exige los dos ceros al final, por eso aplicamos ".00 COP"
df_meta['price'] = df_final['price'].apply(lambda x: f"{float(x):.2f} COP")
df_meta['link'] = df_final['id'].apply(lambda x: f"https://diangel-catalogo.vercel.app/producto/{x}")

def convert_drive_link_meta(url):
    if not url or str(url).strip() == '': return ""
    first_url = str(url).split(';')[0].strip()
    if 'drive.google.com/file/d/' in first_url:
        file_id = first_url.split('/d/')[1].split('/')[0]
        return f"https://drive.google.com/uc?export=view&id={file_id}"
    return first_url

df_meta['image_link'] = df_final['images'].apply(convert_drive_link_meta)
df_meta['brand'] = 'Diangel Joyería'
df_meta['google_product_category'] = 'Apparel & Accessories > Jewelry'
df_meta['inventory'] = df_final['stock']
df_meta['item_group_id'] = df_final['base_code']
df_meta['color'] = df_final['color']
df_meta['material'] = df_final['material']

# Cerebro clasificador de WhatsApp
def clasificar_whatsapp(row):
    cat = str(row['category']).strip().title()
    nombre = str(row['name']).upper()
    if cat == 'Cadenas' or cat == 'Cadena':
        if '45' in nombre: return 'Cadenas 45 cm'
        elif '65' in nombre or '60' in nombre: return 'Cadenas 65 cm'
        else: return 'Cadenas y Gargantillas'
    return cat

df_meta['custom_label_0'] = df_final.apply(clasificar_whatsapp, axis=1)

# Filtramos los que tengan imagen válida
df_meta = df_meta[df_meta['image_link'] != '']
print(f"📦 Total de productos con foto listos para Meta: {len(df_meta)}")

try:
    try:
        worksheet_meta = sh.worksheet('Meta_Feed')
    except:
        worksheet_meta = sh.add_worksheet(title="Meta_Feed", rows="1000", cols="20")

    worksheet_meta.clear()
    datos_meta = [df_meta.columns.values.tolist()] + df_meta.fillna('').values.tolist()
    worksheet_meta.update(values=datos_meta, range_name='A1')
    print("✅ Pestaña 'Meta_Feed' actualizada con éxito. ¡Lista para Facebook y WhatsApp!")
except Exception as e:
    print(f"❌ Error al subir catálogo a Meta_Feed: {e}")
# sincronizador.py
import time
import os
import glob
import json
import urllib.request
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
# FUNCIÓN SECUNDARIA: MAPEAR FOTOS DE CLOUDINARY
# ==========================================
def mapear_fotos_cloudinary():
    print("⚡ Leyendo base de datos ultrarrápida de Cloudinary...")
    gc = gspread.service_account(filename='credenciales.json')
    sh = gc.open_by_key('1jRTJqwPir1CSGQzKazjgWNo3TWr3b8Nz-RkWHOfatxo')

    try:
        ws_fotos = sh.worksheet('Base_Fotos_Cloudinary')
        registros = ws_fotos.get_all_records()

        diccionario_fotos = {}
        for row in registros:
            if 'codigo' in row and 'enlaces_cloudinary' in row:
                diccionario_fotos[str(row['codigo']).strip()] = str(row['enlaces_cloudinary']).strip()

        return diccionario_fotos
    except Exception as e:
        print(f"⚠️ Error leyendo la pestaña Base_Fotos_Cloudinary: {e}")
        return {}

mapa_imagenes = {}
try:
    mapa_imagenes = mapear_fotos_cloudinary()
    print(f"✅ Se cargaron {len(mapa_imagenes)} galerías de fotos desde Cloudinary.")
except Exception as e:
    print(f"⚠️ No se pudieron cargar las fotos: {e}")

# ==========================================
# FASE 2: DESCARGAR EXCELES CON SELENIUM
# ==========================================
opciones = Options()
opciones.add_argument('--headless=new')
opciones.add_argument('--no-sandbox')
opciones.add_argument('--disable-dev-shm-usage')
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

    boton_login = driver.find_element(By.ID, 'login-text')
    driver.execute_script("arguments[0].click();", boton_login)
    time.sleep(5)

    print("📥 Descargando catálogo de Oro Laminado 18K...")
    driver.get('https://belatrizcolombia.com/app/public/biblioteca')
    time.sleep(3)

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
    cat_original = str(row.get('CATEGORIA', '')).strip().upper()
    nombre_original = str(row.get('NOMBRE', '')).strip()
    cat_map = {
        'ARETES': ('Aretes', 'Arete'), 'DIJES': ('Dijes', 'Dije'),
        'CADENAS': ('Cadenas', 'Cadena'), 'PULSERAS': ('Pulseras', 'Pulsera'),
        'ANILLOS': ('Anillos', 'Anillo'), 'HERRAJES': ('Herrajes', 'Herraje'),
        'EXCLUSIVIDADES': ('Exclusividades', 'Exclusividad'),
        'DIFERENCIAL': ('Diferencial', 'Diferencial'),
        'GARGANTILLA': ('Gargantillas', 'Gargantilla'), 'OTROS': ('Otros', 'Otros')
    }
    cat_plural, cat_singular = cat_map.get(cat_original, (cat_original.title(), cat_original.title()))
    if not nombre_original or nombre_original.lower() == 'nan':
        return cat_singular
    palabras = nombre_original.split()
    if not palabras:
        return cat_singular
    primer_elemento = palabras[0]
    if len(primer_elemento) == 1 and primer_elemento.isalpha():
        nombre_limpio = " ".join(palabras[1:])
    elif primer_elemento.lower() in [cat_plural.lower(), cat_singular.lower(), cat_original.lower()]:
        nombre_limpio = " ".join(palabras[1:])
    else:
        nombre_limpio = nombre_original
    return f"{cat_singular} {nombre_limpio}".strip()

def extraer_variantes(referencia):
    if '-' not in referencia:
        return referencia, ''
    partes = referencia.split('-')
    if len(partes) >= 2 and partes[0].isalpha() and len(partes[0]) <= 2:
        codigo_base = f"{partes[0]}-{partes[1]}".strip()
        indice_variante = 2
    else:
        codigo_base = partes[0].strip()
        indice_variante = 1

    variante = partes[indice_variante].strip().upper() if len(partes) > indice_variante else ''
    if not variante:
        return codigo_base, ''
    if len(variante) == 1:
        return referencia, ''
    if variante.startswith('T') and variante[1:].replace('.', '').isdigit():
        return referencia, ''
    return codigo_base, variante

# Función blindada para convertir cualquier precio a entero sin fallar con celdas vacías
def limpiar_precio_seguro(p):
    if pd.isna(p) or str(p).strip() == '':
        return 0
    if isinstance(p, (int, float)):
        if 0 < p < 1000:
            return int(p * 1000)
        return int(p)
    p_str = str(p).replace('$', '').replace('COP', '').replace('.', '').replace(',', '').strip()
    try:
        return int(float(p_str))
    except:
        return 0

df_final = pd.DataFrame()
df_final['id'] = df_proveedor.index + 1

referencias_limpias = df_proveedor['REFERENCIA'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
bases_y_colores = referencias_limpias.apply(extraer_variantes)

df_final['code'] = referencias_limpias
df_final['base_code'] = [x[0] for x in bases_y_colores]
df_final['name'] = df_proveedor.apply(limpiar_y_combinar_nombre, axis=1)
df_final['collection'] = ''
df_final['category'] = df_proveedor['CATEGORIA'].astype(str).str.strip().str.title()
df_final['material'] = df_proveedor['LINEA PRODUCTO'].fillna('Oro Laminado 18K')
df_final['color'] = [x[1] for x in bases_y_colores]

# 🔥 CORRECCIÓN CRÍTICA: Si PRECIO DETAL viene vacío en Belatriz, queda en 0 (no en '')
precios_detal_base = df_proveedor['PRECIO DETAL'].apply(limpiar_precio_seguro)
df_final['price'] = precios_detal_base.apply(lambda x: x + 10000 if x > 0 else 0)
df_final['previousPrice'] = ''

precios_emp_base = df_proveedor['PRECIO EMPRENDEDOR'].apply(limpiar_precio_seguro)
df_final['precioMayorista'] = precios_emp_base

df_final['stock'] = df_proveedor['ESTATUS'].apply(lambda x: 10 if str(x).strip().lower() == 'activo' else 0)
df_final['status'] = df_final['stock'].apply(lambda x: 'Normal' if x > 0 else 'Agotado')
df_final['images'] = df_final['code'].map(mapa_imagenes).fillna(df_final['base_code'].map(mapa_imagenes)).fillna('')
df_final['description'] = 'Hermosa pieza de ' + df_final['material'].astype(str)
df_final['tags'] = ''

# Limpiamos filas inválidas sin código
df_final = df_final[df_final['code'] != ''].copy()
df_final = df_final[df_final['code'].str.lower() != 'nan'].copy()
df_final = df_final.fillna('')
df_final['id'] = range(1, len(df_final) + 1)

# ==========================================
# FASE 3.5: INYECTAR PRODUCTOS PERSONALIZADOS (SHEETS + ERP)
# ==========================================
print("🔗 Inyectando catálogo de Personalizados...")
gc = gspread.service_account(filename='credenciales.json')
sh = gc.open_by_key('1jRTJqwPir1CSGQzKazjgWNo3TWr3b8Nz-RkWHOfatxo')

try:
    worksheet_pers = sh.worksheet('Personalizados')
    datos_pers = worksheet_pers.get_all_records()

    if datos_pers:
        df_pers = pd.DataFrame(datos_pers)
        for col in df_final.columns:
            if col not in df_pers.columns:
                df_pers[col] = ''

        df_pers['price'] = df_pers['price'].apply(limpiar_precio_seguro)
        if 'precioMayorista' not in df_pers.columns:
            df_pers['precioMayorista'] = df_pers['price']
        else:
            df_pers['precioMayorista'] = df_pers['precioMayorista'].apply(limpiar_precio_seguro)

        df_pers['status'] = df_pers['status'].apply(lambda x: 'Normal' if str(x).strip() == '' else x)
        df_pers['description'] = df_pers['description'].apply(lambda x: 'Hermosa pieza exclusiva' if str(x).strip() == '' else x)

        def asignar_imagen(row):
            img_actual = str(row.get('images', '')).strip()
            if img_actual:
                return img_actual
            codigo = str(row.get('code', ''))
            base = str(row.get('base_code', ''))
            if codigo in mapa_imagenes:
                return mapa_imagenes[codigo]
            if base in mapa_imagenes:
                return mapa_imagenes[base]
            return ''

        df_pers['images'] = df_pers.apply(asignar_imagen, axis=1)
        df_pers['stock'] = pd.to_numeric(df_pers['stock'], errors='coerce').fillna(0).astype(int)

        df_final = pd.concat([df_final, df_pers[df_final.columns]], ignore_index=True)
        print(f"✅ ¡{len(df_pers)} productos personalizados de Sheets procesados!")
except Exception as e:
    print(f"⚠️ Omitiendo personalizados de Sheets: {e}")

# Inyectar también personalizados creados directamente desde el ERP con fotos de Cloudinary
try:
    erp_url = os.environ.get('ERP_PUBLIC_URL', 'https://www.diangeljoyeria.com')
    req = urllib.request.Request(f"{erp_url}/api/erp/custom-products", headers={'User-Agent': 'DiangelBot/1.0'})
    with urllib.request.urlopen(req, timeout=15) as resp:
        erp_data = json.loads(resp.read().decode('utf-8'))
        if erp_data.get('success') and erp_data.get('products'):
            df_erp = pd.DataFrame(erp_data['products'])
            for col in df_final.columns:
                if col not in df_erp.columns:
                    df_erp[col] = ''
            # Evitar duplicar códigos que ya estén en df_final
            codigos_existentes = set(df_final['code'].astype(str).str.upper())
            df_erp = df_erp[~df_erp['code'].astype(str).str.upper().isin(codigos_existentes)]
            if len(df_erp) > 0:
                df_final = pd.concat([df_final, df_erp[df_final.columns]], ignore_index=True)
                print(f"✅ ¡{len(df_erp)} productos personalizados desde el ERP inyectados!")
except Exception as e:
    print(f"ℹ️ Sin productos adicionales desde API ERP (o aún no desplegado): {e}")

df_final['price'] = df_final['price'].apply(limpiar_precio_seguro)
df_final['precioMayorista'] = df_final['precioMayorista'].apply(limpiar_precio_seguro)
df_final['stock'] = pd.to_numeric(df_final['stock'], errors='coerce').fillna(0).astype(int)
df_final['id'] = range(1, len(df_final) + 1)

# ==========================================
# FASE 3.8: CREAR HOJA MAESTRA PARA FACTURACIÓN
# ==========================================
print("🗄️ Guardando Inventario Completo (Activos + Agotados)...")
try:
    try:
        ws_completo = sh.worksheet('Inventario_Completo')
    except:
        ws_completo = sh.add_worksheet(title="Inventario_Completo", rows="5000", cols="20")

    ws_completo.clear()
    datos_completo = [df_final.columns.values.tolist()] + df_final.fillna('').values.tolist()
    ws_completo.update(values=datos_completo, range_name='A1')
    print("✅ 'Inventario_Completo' guardado con éxito.")
except Exception as e:
    print(f"❌ Error al subir Inventario Completo: {e}")

# ==========================================
# FASE 4: FILTRAR Y SUBIR SOLO ACTIVOS CON PRECIO > 0 A HOJA 1
# ==========================================
# 🔥 CORRECCIÓN: Exigimos tanto stock > 0 como price > 0
df_activos = df_final[(df_final['stock'] > 0) & (df_final['price'] > 0)].copy()

print("⬆️ Subiendo catálogo filtrado a Hoja 1 (Página web)...")
try:
    worksheet_hoja1 = sh.worksheet('Hoja 1')
    worksheet_hoja1.clear()
    datos_hoja1 = [df_activos.columns.values.tolist()] + df_activos.fillna('').values.tolist()
    worksheet_hoja1.update(values=datos_hoja1, range_name='A1')
    print("✅ 'Hoja 1' actualizada con éxito.")
except Exception as e:
    print(f"❌ Error al subir a Hoja 1: {e}")

# ==========================================
# FASE 5: GENERAR CATÁLOGO PARA META Y WHATSAPP
# ==========================================
print("🛒 Generando feed estructurado para Meta Business...")
df_meta = pd.DataFrame()

df_meta['id'] = df_activos['id']
df_meta['title'] = df_activos['name'].astype(str).str[:200]
df_meta['description'] = df_activos['description'].astype(str).str[:9999]
df_meta['availability'] = df_activos['stock'].apply(lambda x: 'in stock' if x > 0 else 'out of stock')
df_meta['condition'] = 'new'
df_meta['price'] = df_activos['price'].apply(lambda x: f"{float(limpiar_precio_seguro(x)):.2f} COP")
df_meta['link'] = df_activos['code'].apply(lambda x: f"https://www.diangeljoyeria.com/producto/{x}")

def extraer_foto_principal(url_string):
    if not url_string or str(url_string).strip() == '':
        return ""
    urls = [u.strip() for u in str(url_string).split(';') if u.strip()]
    if not urls:
        return ""
    return urls[0]

def extraer_fotos_adicionales(url_string):
    if not url_string or str(url_string).strip() == '':
        return ""
    urls = [u.strip() for u in str(url_string).split(';') if u.strip()]
    if len(urls) <= 1:
        return ""
    return ",".join(urls[1:])

df_meta['image_link'] = df_activos['images'].apply(extraer_foto_principal)
df_meta['additional_image_link'] = df_activos['images'].apply(extraer_fotos_adicionales)

df_meta['brand'] = 'Diangel Joyería'
df_meta['google_product_category'] = 'Apparel & Accessories > Jewelry'
df_meta['inventory'] = df_activos['stock']
df_meta['item_group_id'] = df_activos['base_code']
df_meta['color'] = df_activos['color']
df_meta['material'] = df_activos['material']

def clasificar_whatsapp(row):
    cat = str(row['category']).strip().title()
    nombre = str(row['name']).upper()
    if cat == 'Cadenas' or cat == 'Cadena':
        if '45' in nombre:
            return 'Cadenas 45 cm'
        elif '65' in nombre or '60' in nombre:
            return 'Cadenas 65 cm'
        else:
            return 'Cadenas y Gargantillas'
    return cat

df_meta['custom_label_0'] = df_activos.apply(clasificar_whatsapp, axis=1)
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
    print("✅ Pestaña 'Meta_Feed' actualizada con éxito.")
except Exception as e:
    print(f"❌ Error al subir catálogo a Meta_Feed: {e}")

# ==========================================
# FASE 6: NOTIFICAR AUTOMÁTICAMENTE AL ERP (SUPABASE)
# ==========================================
print("🔄 Sincronizando automáticamente la base de datos del ERP (Supabase)...")
try:
    erp_url = os.environ.get('ERP_PUBLIC_URL', 'https://www.diangeljoyeria.com')
    cron_secret = os.environ.get('ERP_CRON_SECRET', 'diangel_cron_sync_2026')
    req = urllib.request.Request(
        f"{erp_url}/api/erp/cron-sync",
        method="POST",
        headers={
            'Authorization': f'Bearer {cron_secret}',
            'User-Agent': 'DiangelBot/1.0'
        }
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        res_json = json.loads(resp.read().decode('utf-8'))
        print(f"🚀 ERP Supabase actualizado automáticamente: {res_json}")
except Exception as e:
    print(f"ℹ️ Aviso: No se pudo contactar el webhook del ERP (normal si aún no se ha desplegado en Vercel): {e}")
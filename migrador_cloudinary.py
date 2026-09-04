# migrador_cloudinary.py
import time
import io
import gspread
import cloudinary
import cloudinary.uploader
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.http import MediaIoBaseDownload

print("🚀 Iniciando el Gran Migrador a Cloudinary...")

# ==========================================
# 1. CONFIGURACIÓN CLOUDINARY
# ==========================================
cloudinary.config(
    cloud_name="gumwglws",
    api_key="248582635685111",
    api_secret="qfFNrF4LqP4YaNdrvNabhmyYJiE", # 🔥 REEMPLAZA ESTO Y NO LO COMPARTAS
    secure=True
)

# ==========================================
# 2. CONEXIÓN A GOOGLE DRIVE Y SHEETS
# ==========================================
print("🔐 Conectando con Google...")
gc = gspread.service_account(filename='credenciales.json')
sh = gc.open_by_key('1jRTJqwPir1CSGQzKazjgWNo3TWr3b8Nz-RkWHOfatxo')

# Crear pestaña de base de datos si no existe
try:
    ws_fotos = sh.worksheet('Base_Fotos_Cloudinary')
except gspread.exceptions.WorksheetNotFound:
    ws_fotos = sh.add_worksheet(title="Base_Fotos_Cloudinary", rows="1000", cols="2")
    ws_fotos.update(values=[['codigo', 'enlaces_cloudinary']], range_name='A1')

# Leer lo que ya se ha subido para poder pausar/reanudar el script sin duplicar
registros_actuales = ws_fotos.get_all_records()
fotos_ya_subidas = {str(row['codigo']): row['enlaces_cloudinary'] for row in registros_actuales if 'codigo' in row}
fila_actual = len(registros_actuales) + 2

scope = ['https://www.googleapis.com/auth/drive.readonly']
creds = Credentials.from_service_account_file('credenciales.json', scopes=scope)
drive_service = build('drive', 'v3', credentials=creds)

ID_CARPETA_PADRE_FOTOS = '1QG5vmof76jSefx9RmMLtLMkaxi_OsL2L'

# ==========================================
# 3. ESCANEAR DRIVE Y MIGRAR
# ==========================================
print("🔍 Buscando carpetas en Drive...")
page_token = None
carpetas = []
while True:
    query_carpetas = f"'{ID_CARPETA_PADRE_FOTOS}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    resultados = drive_service.files().list(q=query_carpetas, fields="nextPageToken, files(id, name)", pageSize=1000, pageToken=page_token).execute()
    carpetas.extend(resultados.get('files', []))
    page_token = resultados.get('nextPageToken')
    if not page_token: break

print(f"📦 Se encontraron {len(carpetas)} productos con fotos.")

for carp in carpetas:
    codigo_referencia = str(carp['name']).strip()

    # Si ya lo migramos, lo saltamos (ideal por si se va el internet)
    if codigo_referencia in fotos_ya_subidas:
        print(f"⏭️ {codigo_referencia} ya está en Cloudinary. Saltando...")
        continue

    print(f"☁️ Migrando fotos de: {codigo_referencia}...")
    id_carpeta = carp['id']
    # 🔍 OBTENER Y ORDENAR FOTOS (PERFIL PRIMERO)
    query_fotos = f"'{id_carpeta}' in parents and mimeType contains 'image/' and trashed = false"
    fotos = drive_service.files().list(q=query_fotos, fields="files(id, name)").execute().get('files', [])

    foto_perfil = []
    otras_fotos = []

    for foto in fotos:
        nombre_archivo = str(foto.get('name', '')).upper()
        if 'PERFIL' in nombre_archivo:
            foto_perfil.append(foto)
        else:
            otras_fotos.append(foto)

    # Fusionamos asegurando que PERFIL quede de primeras
    fotos_ordenadas = foto_perfil + otras_fotos
    enlaces_optimizados = []

    for foto in fotos_ordenadas:
        # Descargar foto de Drive a la memoria del computador
        request = drive_service.files().get_media(fileId=foto['id'])
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()

        # Subir foto a Cloudinary desde la memoria
        fh.seek(0)
        try:
            # Sube la foto y Cloudinary la convierte automáticamente a formato rápido WebP
            respuesta = cloudinary.uploader.upload(
                fh,
                folder="diangel_joyeria",
                public_id=f"{codigo_referencia}_{foto['name'].split('.')[0]}",
                format="webp"
            )
            enlaces_optimizados.append(respuesta['secure_url'])
        except Exception as e:
            print(f"❌ Error subiendo {foto['name']}: {e}")

    # Si logró subir fotos, las guarda en el Google Sheets inmediatamente
    if enlaces_optimizados:
        texto_enlaces = ";".join(enlaces_optimizados)
        ws_fotos.update(values=[[codigo_referencia, texto_enlaces]], range_name=f'A{fila_actual}')
        fila_actual += 1
        fotos_ya_subidas[codigo_referencia] = texto_enlaces
        print(f"✅ {codigo_referencia} guardado con éxito en Excel.")

print("🎉 ¡Migración finalizada con éxito!")
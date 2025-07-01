import os
import zipfile
import telebot
import pandas as pd
import threading
import webbrowser
from datetime import datetime, time
from telebot import types
from geopy.geocoders import Nominatim
import numpy as np
import time as tm
import requests
import json
import shutil
import glob
import re
from gtts import gTTS
import pygame
import platform
import psutil
import pytz
from telebot.types import Message
from dotenv import load_dotenv
import textwrap
from flask import Flask, request
from estado_global import guardar_estado, cargar_estado  # ✅ nuevas funciones

# Cargar variables de entorno
load_dotenv()

# Zona horaria
tz_lima = pytz.timezone('America/Lima')
inicio_bot = datetime.now(tz_lima)

# Cargar token
TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
if not TOKEN:
    raise ValueError("Error: No se encontró el token de Telegram en las variables de entorno.")

bot = telebot.TeleBot(TOKEN)

bot_activo = True

usuarios_df = pd.DataFrame()

# Carpeta de archivos subidos
CARPETA_ARCHIVOS = "archivos_subidos"
os.makedirs(CARPETA_ARCHIVOS, exist_ok=True)

# Cargar estado desde JSON
estado_excel, ultima_ruta_archivo = cargar_estado()

# Handlers
@bot.message_handler(commands=['estadoexcel'])
def estado_excel_handler(msg):
    estado, _ = cargar_estado()
    bot.send_message(msg.chat.id, estado)

@bot.message_handler(commands=['rutaexcel'])
def ruta_archivo_handler(msg):
    _, ruta = cargar_estado()
    if ruta and os.path.exists(ruta):
        bot.send_message(msg.chat.id, f"📁 Ruta actual del archivo:\n{ruta}")
    else:
        bot.send_message(msg.chat.id, "⚠️ No hay un archivo Excel cargado.")


# Crear la carpeta si no existe
os.makedirs(CARPETA_ARCHIVOS, exist_ok=True)
API_VALIDAR_USUARIO = "https://cybernovasystems.com/prueba/sistema_tlc/modelos/telegran/api_validar_usuario.php"

API_REGISTRAR_ASISTENCIA = "https://cybernovasystems.com/prueba/sistema_tlc/modelos/telegran/api_registrar_asistencia.php"
# Eliminar el webhook si está activo
bot.remove_webhook()
def cargar_datos_excel():
    """Inicia un hilo para cargar el archivo Excel."""
    threading.Thread(target=_cargar_excel_thread).start()

def _cargar_excel_thread():
    """Carga el archivo Excel, detecta la fila de inicio y envía los datos a la API."""
    global usuarios_df
    try:
        # Seleccionar archivo Excel
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx;*.xls")]
        )
        if not file_path:
            print("❌ No se seleccionó ningún archivo.")
            return

        # Guardar el archivo en la carpeta archivos_subidos
        file_name = os.path.basename(file_path)
        new_file_path = os.path.join(CARPETA_ARCHIVOS, file_name)
        shutil.copy(file_path, new_file_path)

        # Detectar la fila donde inicia la data
        fila_inicio = detectar_fila_inicio(new_file_path)
        if fila_inicio is None:
            raise ValueError("⚠️ No se encontró la columna 'CodiSeguiClien'.")

        print(f"✅ Fila detectada correctamente: {fila_inicio}")

        # Cargar datos desde la fila detectada
        df = pd.read_excel(new_file_path, skiprows=fila_inicio - 1)
        df.columns = df.columns.str.strip()
        usuarios_df = df

        # Mostrar una muestra de los datos en consola
        print("\n🔍 **Primeras 5 filas del DataFrame cargado:**")
        print(df.head())

        # Guardar estado actualizado para el bot
        guardar_estado(f"📊 Archivo Excel Cargado: {file_name} ✔️", new_file_path)

        messagebox.showinfo("Éxito", f"Archivo cargado y almacenado en {CARPETA_ARCHIVOS}: {file_name}")

        # Enviar los datos a la API
        enviar_datos_a_api(df)

    except Exception as e:
        messagebox.showerror("Error", f"⚠️ Ocurrió un error: {e}")
        print(f"❌ Error al cargar el archivo Excel: {e}")

def detectar_fila_inicio(file_path):
    """Detecta la fila donde se encuentra la columna 'CodiSeguiClien'."""
    try:
        excel_data = pd.ExcelFile(file_path)
        for sheet in excel_data.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet, header=None)
            for i, row in df.iterrows():
                if 'CodiSeguiClien' in row.values:
                    return i + 1  # Devuelve la fila donde se encuentra
    except Exception as e:
        print(f"❌ Error detectando la fila de inicio: {e}")
    return None
    
usuarios_esperando_ubicacion = {}
usuarios_esperando_imagen = {}

@bot.message_handler(commands=['asistencia'])
def solicitar_ubicacion(message):
    user_id = message.from_user.id
    try:
        response = requests.post(API_VALIDAR_USUARIO, json={"user_id": user_id}, timeout=5)
        data = response.json()

        if not data.get("permitido", False):
            bot.reply_to(message, "⛔ No tienes acceso para registrar asistencia. Contacta a soporte.")
            return  

        usuarios_esperando_ubicacion[user_id] = True
        keyboard = types.ReplyKeyboardMarkup(one_time_keyboard=True, resize_keyboard=True)
        boton_ubicacion = types.KeyboardButton(text="📍 Enviar Ubicación", request_location=True)
        keyboard.add(boton_ubicacion)
        
        bot.send_message(message.chat.id, 
                         "📍 *Comparte tu ubicación para registrar asistencia.*\n\n"
                         "Presiona el botón de abajo para enviarla automáticamente.",
                         reply_markup=keyboard, parse_mode="Markdown")
    except requests.exceptions.RequestException as e:
        bot.reply_to(message, "⚠️ Error al verificar asistencia. Inténtalo más tarde.")
        print(f"❌ Error en la API de asistencia: {e}")

@bot.message_handler(content_types=['location'])
def recibir_ubicacion(message):
    user_id = message.from_user.id
    latitud = message.location.latitude
    longitud = message.location.longitude
    nombre_tecnico = message.from_user.full_name

    usuarios_esperando_imagen[user_id] = {
        "latitud": latitud,
        "longitud": longitud,
        "nombre_tecnico": nombre_tecnico
    }

    bot.reply_to(message, "📸 Ahora envía una foto rotulada para completar tu asistencia.")
@bot.message_handler(content_types=['photo'])
def recibir_imagen(message: Message):
    user_id = message.from_user.id
    
    if user_id not in usuarios_esperando_imagen:
        bot.reply_to(message, "⚠️ Primero debes enviar tu ubicación antes de la foto.")
        return
    
    mensaje_carga = bot.reply_to(message, "⏳ Procesando tu solicitud...")
    
    datos_ubicacion = usuarios_esperando_imagen[user_id]
    file_id = message.photo[-1].file_id
    file_info = bot.get_file(file_id)
    file_path = file_info.file_path
    image_url = f"https://api.telegram.org/file/bot{TOKEN}/{file_path}"
    image_path = f"imagenes/{user_id}.jpg"
    
    print(f"📸 Recibida imagen de {user_id}")
    print(f"🖼️ File ID: {file_id}")
    print(f"📂 File Path: {file_path}")
    print(f"🌍 URL de la imagen: {image_url}")
    
    # Crear la carpeta si no existe
    if not os.path.exists("imagenes"):
        os.makedirs("imagenes")
    
    # Descargar la imagen
    response = requests.get(image_url, stream=True)
    if response.status_code == 200:
        with open(image_path, "wb") as file:
            for chunk in response.iter_content(1024):
                file.write(chunk)
        print(f"✅ Imagen guardada en: {image_path}")
    else:
        print("❌ Error al descargar la imagen")
        bot.edit_message_text("⚠️ Error al descargar la imagen.", message.chat.id, mensaje_carga.message_id)
        return
    
    # Datos a enviar a la API
    datos = {
        "user_id": str(user_id),
        "latitud": str(datos_ubicacion["latitud"]),
        "longitud": str(datos_ubicacion["longitud"])
     
    }
    
    # Enviar imagen y datos a la API
    with open(image_path, "rb") as image_file:
        files = {"imagen": image_file}
        try:
            response = requests.post(API_REGISTRAR_ASISTENCIA, data=datos, files=files, timeout=10)
            print(f"🔄 Código de respuesta API: {response.status_code}")
            print(f"📩 Respuesta API: {response.text}")
            
            if response.status_code == 200:
                try:
                    data = response.json()
                    print(f"📊 Datos de la API: {data}")
                except ValueError:
                    print(f"⚠️ Error al interpretar JSON: {response.text}")
                    bot.edit_message_text("⚠️ Respuesta inválida del servidor. Inténtalo más tarde.", message.chat.id, mensaje_carga.message_id)
                    return
                
                if data.get("asistencia_registrada", False):
                    mensaje_confirmacion = (
                        f"✅ *👷‍♂️ {datos_ubicacion['nombre_tecnico']} (ID: {user_id}) ha enviado su asistencia correctamente.*\n\n"
                        "📌 La gestora revisará tu solicitud y te dará acceso al bot. Por favor, espera su aprobación. ⏳"
                    )
                    bot.edit_message_text(mensaje_confirmacion, message.chat.id, mensaje_carga.message_id)
                else:
                    mensaje_api = data.get("mensaje", "⛔ No puedes marcar asistencia desde esta ubicación.")
                    bot.edit_message_text(f"⛔ {mensaje_api}", message.chat.id, mensaje_carga.message_id)
            else:
                bot.edit_message_text("⚠️ Error al registrar asistencia. Inténtalo más tarde.", message.chat.id, mensaje_carga.message_id)
        except requests.exceptions.RequestException as e:
            print(f"❌ Error al conectar con la API: {e}")
            bot.edit_message_text("⚠️ Error al registrar asistencia. Inténtalo más tarde.", message.chat.id, mensaje_carga.message_id)
    
    # Eliminar la imagen después de enviarla
    if os.path.exists(image_path):
        try:
            os.remove(image_path)
            print(f"🗑️ Imagen eliminada: {image_path}")
        except Exception as e:
            print(f"⚠️ No se pudo eliminar la imagen: {e}")
    
    usuarios_esperando_imagen.pop(user_id, None)
    
PASSWORD_CORRECTA = "1"
usuarios_autorizados = {}

@bot.message_handler(commands=['subir'])
def pedir_contraseña(message: Message):
    """Solicita la contraseña antes de permitir subir archivos."""
    chat_id = message.chat.id
    bot.send_message(chat_id, "🔑 Ingresa la contraseña para subir un archivo:")
    bot.register_next_step_handler(message, verificar_contraseña)

def verificar_contraseña(message: Message):
    """Verifica la contraseña y da permiso temporal si es correcta."""
    chat_id = message.chat.id
    if message.text is not None and message.text.strip() == PASSWORD_CORRECTA:
        usuarios_autorizados[chat_id] = tm.time()  # Guarda el tiempo de autorización
        bot.send_message(chat_id, "✅ Contraseña correcta. Ahora envía tu archivo Excel.")
    else:
        bot.send_message(chat_id, "🚫 Contraseña incorrecta. Intenta de nuevo con /subir.")

@bot.message_handler(content_types=['document'])
def recibir_archivo(message: Message):
    """Recibe un archivo Excel solo si el usuario fue autorizado."""
    global usuarios_df, estado_excel
    chat_id = message.chat.id

    if chat_id not in usuarios_autorizados or (tm.time() - usuarios_autorizados[chat_id] > 300):
        bot.send_message(chat_id, "⛔ No tienes permiso para subir archivos. Usa /subir para autenticarte primero.")
        return

    file_name = ""
    file_path = ""
    estado_excel_anterior = estado_excel  # Guardamos el estado anterior

    try:
        nombre_usuario = message.from_user.first_name
        apellido_usuario = message.from_user.last_name or ""
        nombre_completo = f"{nombre_usuario} {apellido_usuario}".strip()

        file_info = bot.get_file(message.document.file_id)
        file_extension = file_info.file_path.split('.')[-1].lower()

        if file_extension not in ['xlsx', 'xls']:
            mensaje_error = f"❌ Error. {nombre_completo}, solo se permiten archivos Excel (xlsx o xls)."
            bot.send_message(chat_id, mensaje_error)
            return

        estado_excel = f"📥 {nombre_completo} está subiendo un archivo..."
        bot.send_message(chat_id, estado_excel)

        file_name = message.document.file_name
        file_path = os.path.join(CARPETA_ARCHIVOS, file_name)
        downloaded_file = bot.download_file(file_info.file_path)

        with open(file_path, "wb") as new_file:
            new_file.write(downloaded_file)

        fila_inicio = detectar_fila_inicio(file_path)
        if fila_inicio is None:
            raise ValueError("No se encontró la fila de inicio.")

        df = pd.read_excel(file_path, skiprows=fila_inicio - 1, engine="openpyxl")
        df.columns = df.columns.str.strip()
        usuarios_df = df

        estado_excel = f"📊 {nombre_completo}, tu archivo {file_name} fue cargado con éxito. ✔️"
        bot.send_message(chat_id, estado_excel)

        print("\n📊 Archivo cargado con éxito. Datos extraídos:")
        print(df.head())  # Mostrar las primeras filas del DataFrame en la consola

        if not df.empty:
            enviar_datos_a_api(df)
        else:
            bot.send_message(chat_id, f"⚠️ {nombre_completo}, el archivo está vacío o no tiene datos válidos.")

        usuarios_autorizados.pop(chat_id, None)

    except Exception as e:
        mensaje_error = f"❌ {nombre_completo}, hubo un error. Sube de nuevo. {str(e)}"
        bot.send_message(chat_id, mensaje_error)

        if file_path and os.path.exists(file_path):
            os.remove(file_path)

        estado_excel = estado_excel_anterior  # Restauramos el estado anterior




def manejar_exito(chat_id, nombre, archivo):
    """Maneja el proceso de éxito después de cargar el archivo."""
    threading.Thread(target=enviar_mensaje_voz_por_telegram, args=(chat_id, f"Hola {nombre}, el archivo {archivo} fue subido con éxito y está listo para su uso.")).start()


def manejar_error(chat_id, nombre, mensaje, archivo=None, file_path=None):
    """Maneja errores en la carga de archivos y elimina solo los archivos Excel fallidos."""

    # Verificar si el archivo existe antes de intentar eliminarlo
    if file_path and os.path.exists(file_path):
        try:
            os.remove(file_path)
            print(f"✅ Archivo eliminado correctamente: {file_path}")
        except Exception as e:
            print(f"⚠️ No se pudo eliminar el archivo {file_path}: {e}")
    else:
        print(f"⚠️ Archivo no encontrado para eliminar: {file_path}")

    bot.send_message(chat_id, mensaje)
    threading.Thread(target=enviar_mensaje_voz_por_telegram, args=(chat_id, mensaje)).start()


def enviar_mensaje_voz_por_telegram(chat_id, texto):
    """Envía un mensaje de voz a través de Telegram."""
    try:
        audio_path = os.path.join(CARPETA_ARCHIVOS, "mensaje_voz.ogg")
        tts = gTTS(text=texto, lang="es")
        tts.save(audio_path)
        with open(audio_path, 'rb') as audio_file:
            bot.send_voice(chat_id, audio_file)
        time.sleep(2)
        os.remove(audio_path)
    except Exception as e:
        print(f"❌ Error al enviar mensaje de voz por Telegram: {e}")


def liberar_archivo(file_path):
    """Espera hasta que el archivo esté disponible para su eliminación."""
    import time
    while True:
        try:
            os.remove(file_path)
            break
        except PermissionError:
            time.sleep(1)


# Handler para el comando /zip (solo archivos Excel del 21 de mayo de 2025)
@bot.message_handler(commands=['zip'])
def enviar_zip(message: Message):
    chat_id = message.chat.id
    zip_path = comprimir_excel_por_fecha("2025-05-21")  # Fecha objetivo: 2025-05-21

    if zip_path and os.path.exists(zip_path):
        with open(zip_path, 'rb') as archivo_zip:
            bot.send_document(chat_id, archivo_zip)
        print("✅ ZIP enviado correctamente")
    else:
        bot.send_message(chat_id, "❌ No se encontraron archivos Excel para esa fecha o no se pudo crear el ZIP.")

def comprimir_excel_por_fecha(fecha_str):
    """
    Comprime en un ZIP todos los archivos .xls/.xlsx de la carpeta 'archivos_subidos'
    cuya fecha de modificación coincida con fecha_str (formato 'YYYY-MM-DD').
    Devuelve la ruta del ZIP creado o None si no hay archivos que coincidan.
    """
    carpeta = "archivos_subidos"
    zip_dest = "excel_21_mayo_2025.zip"

    if not os.path.exists(carpeta):
        print("❌ La carpeta 'archivos_subidos' no existe.")
        return None

    # Convertir la cadena a objeto date
    try:
        fecha_obj = datetime.strptime(fecha_str, "%Y-%m-%d").date()
    except ValueError:
        print("❌ Formato de fecha inválido. Debe ser 'YYYY-MM-DD'.")
        return None

    archivos_filtrados = []
    for nombre in os.listdir(carpeta):
        ruta = os.path.join(carpeta, nombre)
        if os.path.isfile(ruta):
            # Solo considerar archivos .xls y .xlsx
            ext = nombre.lower().split('.')[-1]
            if ext not in ['xls', 'xlsx']:
                continue

            # Obtener fecha de modificación
            fecha_mod = datetime.fromtimestamp(os.path.getmtime(ruta)).date()
            if fecha_mod == fecha_obj:
                archivos_filtrados.append(ruta)

    if not archivos_filtrados:
        print(f"⚠️ No se encontraron archivos Excel modificados el {fecha_str}.")
        return None

    # Crear el ZIP con los archivos filtrados
    with zipfile.ZipFile(zip_dest, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for ruta_archivo in archivos_filtrados:
            arcname = os.path.basename(ruta_archivo)
            zipf.write(ruta_archivo, arcname=arcname)

    print(f"✅ ZIP creado: {zip_dest} ({len(archivos_filtrados)} archivos)")
    return zip_dest



def enviar_datos_a_api(df):
    """Convierte los datos del DataFrame en JSON y los envía a la API automáticamente."""
    try:
        # 🔍 Verificar columnas disponibles
        print("🧪 Columnas en DataFrame:", df.columns.tolist())

        # 🚧 Convertir columnas numéricas
        df['OrdenId'] = pd.to_numeric(df['OrdenId'], errors='coerce')
        df['CodiSeguiClien'] = pd.to_numeric(df['CodiSeguiClien'], errors='coerce')

        # 🧹 Eliminar filas con orden_id inválido
        df = df.dropna(subset=['OrdenId'])

        # 🕒 Formatear fechas
        fechas_a_formatear = ['FechaUltiEsta', 'FechaIniVisi', 'FechaFinVisi', 'F.Soli']
        for col in fechas_a_formatear:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d/%m/%Y %H:%M:%S').fillna('00/00/0000 00:00:00')

        ordenes = []

        for _, row in df.iterrows():
            # 🧠 Datos técnicos y CTO
            datos_tecnicos_raw = str(row.get('Datos Técnicos', 'Desconocida'))
            match_cto = re.search(r'(W-[^;]+)', datos_tecnicos_raw, re.IGNORECASE)
            codigo_cto = match_cto.group(1) if match_cto else None
            datos_tecnicos = datos_tecnicos_raw.strip()

            # ☎️ Manejo de teléfonos
            telefono_movil = str(row.get('TeleMovilNume', 'No disponible'))
            telefono_fijo = str(row.get('TeleFijoNume', 'No disponible'))
            if telefono_fijo.endswith('.0'):
                telefono_fijo = telefono_fijo[:-2]

            orden = {
                "orden_id": int(row['OrdenId']),
                "codigo_seguimiento": int(row['CodiSeguiClien']) if pd.notna(row['CodiSeguiClien']) else None,
                "cuadrilla": str(row.get('Cuadrilla', 'Desconocida')),
                "cliente": str(row.get('Cliente', 'No especificado')),
                "estado": str(row.get('Estado', 'En Revisión')),
                "direccion": str(row.get('Direccion', 'No especificada')),
                "dni": str(row.get('Número Documento', 'No disponible')),
                "telefono": telefono_movil,
                "telefono_fijo": telefono_fijo,
                "ticket": str(row.get('CodiSegui', 'No asignado')),
                "zona": str(row.get('Zona', 'No especificada')),
                "tipo_traba": str(row.get('TipoTraba', 'No especificado')),
                "fecha_ulti_esta": row['FechaUltiEsta'],
                "fecha_inicio_visita": row['FechaIniVisi'],
                "fecha_fin_visita": row['FechaFinVisi'],
                "fecha_solicitud": row['F.Soli'],
                "codigo_cto": codigo_cto,
                "datos_tecnicos": datos_tecnicos,
                "sector_operativo": str(row.get('Sector Operativo', 'Desconocido')),
                "producto": str(row.get('Producto', '')),
                "tipo": str(row.get('Tipo', ''))
            }
            ordenes.append(orden)

        # 🌐 Enviar datos a API
        payload = {"ordenes": ordenes}
        url_api = "https://cybernovasystems.com/prueba/sistema_tlc/modelos/telegran/api_guardar_ordenes.php"
        headers = {'Content-Type': 'application/json'}

        print("\n📤 Datos enviados:")
        print(json.dumps(payload, indent=4, ensure_ascii=False))

        response = requests.post(url_api, json=payload, headers=headers)
        respuesta_api = response.json()

        print("\n📥 Respuesta de la API:")
        print(respuesta_api)

        if 'mensaje' in respuesta_api:
            print(f"✅ Éxito: {respuesta_api['mensaje']}")
        else:
            print(f"❌ Error: {respuesta_api.get('errores', ['Error desconocido'])}")

    except requests.exceptions.RequestException as e:
        print(f"❌ Error de conexión: {e}")
    except json.JSONDecodeError as e:
        print(f"❌ Error en la respuesta JSON: {e}")
    except Exception as e:
        print(f"❌ Error inesperado: {e}")

# hola
# def actualizar_estado_excel(texto, color):
#     estado_excel_label.config(text=f"{texto}", foreground=color)


def seleccionar_plantilla(tipo_trabajo, dni, ordenid, cliente, sn_actual, direccion, producto, cuadrilla, motivo_trabajo, estado, region, motivo_regestion,distrito,tecnico,Zona,telefono,codigo,ticket,ot):
    plantillas = {
        'REGISTRO DE LLEGADA': textwrap.dedent(
            f" **TICKET:** {ticket} \n"
            f" **NOMBRE DEL CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **DIRECCION :** {direccion}\n"
            f" **CONTRATA:** TLI\n"
            f" **CUADRILLA:** {cuadrilla}\n"
            f" **OBS:** \n"
        ),
        'PLANTILLA DE LLEGADA ALTO VALOR': textwrap.dedent(
            f" **CLIENTE:** {cliente} \n"
            f" **DNI:** {dni} \n"
            f" **DIRECCIÓN:** {direccion} \n"
            f" **TELÉFONO:** {telefono} \n"
            f" **CÓDIGO:** {codigo} \n"
            f" **DISTRITO:** {distrito} \n"
            f" **SEGMENTO:** ALTO VALOR \n"
            f" **PLACA:**  \n"
            f" **TICKET:** {ticket} \n"
            f" **CUADRILLA:** {cuadrilla} \n"
            f" **OBSERVACIONES:** REGISTRAR LLEGA\n"
        ),
        'VALIDACION DE LLAMADAS': textwrap.dedent(
            f" **CLIENTE/A:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **CONTRATA:** TLI\n"
            f" **FECHA Y TRAMO SOLICITADO:** \n"
            f" **NUMERO DEL CLIENTE AL CUAL SE COMUNICO:**  {telefono}\n"
            f" **MOTIVO:** \n"
            f" **TICKET:** {ticket} \n"
            f" **OBSERVACIONES:** \n"
        ),
        'ORDEN CON RESERVA': textwrap.dedent(
            f" **NOMBRE DEL CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **TELEFONO:** \n"
            f" **CODIGO:** {codigo}\n"
            f" **OT:** {ot}\n"
            f" **DIRECCION:** {direccion}\n"
            f" **MOTIVO:** \n"
            f" **OBS:** \n"
            f" **DIA:**        TRAMO: \n"
        ),
        'AUTORIZACIÓN PARA RECABLEADO O TRASLADO': textwrap.dedent(
            f" **TICKET:** {ticket} \n"
            f" **CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **CTO O CAJA NAP:** \n"
            f" **TORRE / PISO DE CAJA NAP:** \n"
            f" **DIRECCION O NOMBRE DE EDIFICIO / CONDOMINIO / RESIDENCIAL:** {direccion}\n"
            f" **NUMERO DE PUERTO:** \n"
            f" **POTENCIA DE PUERTO:** \n"
            f" **METRAJE UTILIZADO:** \n"
            f" **NOMBRE DEL TECNICO:** {cuadrilla} \n"
            f" **OBSERVACION:** \n"
            f" **MOTIVO CAMBIO DE CTO O CAJA NAP:** \n"
            f" **1. RECABLEADO A LA MISMA CTO O CAJA NAP** \n"
            f" **2. CAMBIO DE CTO O CAJA NAP** \n"
            f" **3. TRASLADO ** \n"
            f" **4. REINSTALACION** \n"
            f" **5. REUBICACION CON RESERVA** \n"
            f" **6. REUBICACION SIN RESERVA** \n"
        ),
        'UTILIZACION DE PUERTO DE BAJA': textwrap.dedent(
            f" **TICKET:** {ticket} \n"
            f" **CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **CTO O CAJA NAP:** \n"
            f" **CTO O CAJA NAP:** \n"
            f" **POTENCIA DE PUERTO UTILIZADO:** \n"
            f" **DNI DEL CLIENTE AFECTADO:** \n"
            f" **CONTRATA:** TLI \n"
            f" **OBS:** \n"
        ),
        'ALERTA POR CLIENTE DESCONECTADO': textwrap.dedent(
            f" **TICKET:** {ticket} \n"
            f" **CLIENTE ATENDIDO:** {cliente}\n"
            f" **DNI DEL CLIENTE ATENDIDO:** {dni}\n"
            f" **CTO o CAJA NAP:** \n"
            f" **PUERTO:** \n"
            f" **DNI DEL CLIENTE CONECTADO ACTUALMENTE EN EL PUERTO:** \n"
            f" **DATOS DEL TECNICO QUE REPORTA:** {cuadrilla}\n"
            f" **CONTRATA:** TLI\n"
            f" **OBS:** \n"
        ),
        'VALIDAR CASOS CON COSTO / CAMBIO DE TICKET': textwrap.dedent(
            f" **NOMBRE DEL CLIENTE(A):** {cliente} \n"
            f" **DNI:** {dni}\n"
            f" **TELEFONO:** {telefono}\n"
            f" **DIRECCION:** {direccion}\n"
            f" **TICKET:** {ticket} \n"
            f" **SOLICITUD O DAÑO DETECTAD:** \n"
            f" **OBSERVACION:** \n"
        ),
        'SUPERVISOR': textwrap.dedent(
            f" **NOMBRE DEL CLIENTE:** {cliente} \n"
            f" **DNI:** {dni}\n"
            f" **NUMERO:** {telefono}\n"
            f" **DIRECCION:** {direccion}\n"
            f" **DISTRITO:** {Zona}\n"
            f" **CONTRATA:** TLI\n"
            f" **TICKET:** {ticket} \n"
            f" **TIPO DE ATENCION:** SUPERVISION \n"
            f" **OBSERVACION:** \n"
        ),
        'REPORTE PEXT': textwrap.dedent(
            f" **TICKET:** {ticket} \n"
            f" **NOMBRE DEL CLIENTE:** {cliente} \n"
            f" **DNI:** {dni}\n"
            f" **DISTRITO:** {Zona}\n"
            f" **COORDENADAS DE LA CTO:** \n"
            f" **CTO:** \n"
            f" **NOMBRE Y APELLIDO DEL TEC:** {cuadrilla}\n"
            f" **CONTRATA:** TLI \n"
            f" **OBS:** \n"
            f" **EN CASO DE SER CONDOMINIO:** \n"
            f" **  . NOMBRE DEL CONDOMINIO** \n"
            f" **  . NUMERO DE PISO DE LA CAJA ** \n"
            f" **FOTOS PARA PRESENTAR PEXT:** \n"
            f" **FOTO DE LA CTO ROTULADA CERRADA:** \n"
            f" **FOTO DE LA CTO ROTULADA ABIERTA:** \n"
            f" **FOTOS DE LOS PUERTOS DEGRADADOS:** \n"
        ),
        'CAMBIO DE ONT': (
            f" **TICKET:** {ticket} \n"
            f" **CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **SN ONT ANTIGUO:** \n"
            f" **SN ONT NUEVO:** \n"
            f" **PRODUCT ID ONT NUEVA:** \n"
            f" **CODIGO:** {codigo}\n"
            f" **OT:** {ot}\n"
            f" **CONTRATA:** TLI\n"
            f" **TECNICO:** {tecnico}\n"
            f" **DISTRITO:** {Zona}\n"
            f" **MOTIVO DE CAMBIO:** \n"
        ),
        'CAMBIO DE ONT v23': textwrap.dedent(
            f" **TICKET:** {ticket} \n"
            f" **NOMBRE DEL CLIENTE(A):** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **PRO ID:** \n"
            f" **SN DEL ONT ANTIGUO:** \n"
            f" **PRO ID:** \n"
            f" **CODIGO:** {codigo}\n"
            f" **OT:** {ot}\n"
            f" **NOMBRE DEL TECNICO:** {tecnico}\n"
            f" **CONTRATA O PARTNER: TLI** \n"
            f" **OBSERVACION:** \n"
        ),
        'CAMBIO DE ONT v2': textwrap.dedent(
            f" **TICKET:** {ticket} \n"
            f" **CODIGO:** {codigo}\n"
            f" **DNI:** {dni}\n"
            f" **SN DEL ONT ANTIGUO:** \n"
            f" **PRODUCT ID ONT ANTIGUO:** \n"
            f" **SN DEL ONT NUEVA:** \n"
            f" **PRODUCT ID ONT NUEVA:** \n"
            f" **DIRECCION:** {direccion}\n"
            f" **CTO:** \n"
            f" **PUERTO:** \n"
            f" **DIRECCION:** {direccion}\n"
            f" **CONTRATA: TLI** \n"
            f" **TECNICO:** {tecnico}\n"
            f" **DISTRITO:** {Zona}\n"
        ),
         'ACTIVACIÓN': textwrap.dedent(
            f" **TICKET:** {ticket} \n"
            f" **CODIGO:** {codigo}\n"
            f" **DNI:** {dni}\n"
            f" **SN DEL ONT ANTIGUO:** \n"
            f" **PRODUCT ID ONT ANTIGUO:** \n"
            f" **SN DEL ONT NUEVA:** \n"
            f" **PRODUCT ID ONT NUEVA:** \n"
            f" **CTO:** \n"
            f" **PUERTO:** \n"
            f" **DIRECCION:** {direccion}\n"
            f" **CONTRATA: TLI** \n"
            f" **TECNICO:** {tecnico}\n"
            f" **DISTRITO:** {Zona}\n"
        ),
        'CAMBIO DE CTO / CAMBIO DE PUERTO / TRASLADO': textwrap.dedent(
            f" **TICKET:** {ticket}\n"
            f" **CODIGO:** {codigo}\n"
            f" **NOMBRE DEL CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **SN DEL ONT ACTUAL:** \n"
            f" **DIRECCION:** {direccion} \n"
            f" **CONDOMINIO/PISO:** \n"
            f" **DISTRITO:** {Zona}\n"
            f" **CTO:** \n"
            f" **PUERTO:** \n"
            f" **CONTRATA:** TLI\n"
            f" **TECNICO:** {tecnico}\n"
            f" **MOTIVO DEL CAMBIO DE CTO/PUERTO:** \n"
        ),
        'REMATRICULACION': textwrap.dedent(
            f" **TICKET:** {ticket}\n"
            f" **CODIGO DEL CLIENTE:** {codigo}\n"
            f" **NOMBRE DEL CLIENTE(A):** {cliente} \n"
            f" **DNI:** {dni}\n"
            f" **SN DEL ONT ACTUAL:** \n"
            f" **PRODUCT ID ONT NUEVA:** \n"
            f" **DISTRITO:** {Zona} \n"
            f" **CONTRATA:** TLI\n"
            f" **TECNICO:** {tecnico}\n"
            f" **MOTIVO DE LA REMATRICULACION:**  \n"
        ),
        'CAMBIO_DE_ONT_v2': textwrap.dedent(  # Cambié el nombre para evitar duplicados
            f"📋 **Código del Cliente:** \n"
            f"📋 **Ticket:** {ticket}\n"
            f"👤 **Cliente(A):** {cliente} \n"
            f"🆔 **DNI:** {dni}\n"
            f"🔖 **SN ONT ANTIGUO:** \n"
            f"🔖 **SN ONT NUEVO:** \n"
            f"📦 **Product ID ONT Actual:** \n"
            f"📋 **Contrata:** \n"
            f"🛠 **Técnico:** {tecnico}\n"
            f"📋 **Plan:** \n"
            f"📍 **Distrito:** {Zona} \n"
            f"📝 **Observación:** \n"
        ),
        'REMATRICULACIONv2': textwrap.dedent(
            f"📋 **Código del Cliente:** \n"
            f"📋 **Ticket:** {codigo}\n"
            f"👤 **Nombre del Cliente(A):** {cliente} \n"
            f"🆔 **DNI:** {dni}\n"
            f"🔖 **SN del ONT Actual:**\n"
            f"📦 **Product ID ONT Actual:** \n"
            f"📦 **Plan:** \n"
            f"📋 **Contrata:** \n"
            f"🛠️ **Técnico:** {tecnico}\n"
            f"📍 **Distrito:** {distrito} \n"
            f"📝 **Motivo de la Rematriculación:** \n"
        ),
        'TRASLADO + CAMBIO DE ONT': textwrap.dedent(
            f" **TICKET:** {ticket}\n"
            f" **NOMBRE DEL CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **SN ONT ANTIGUO:** \n"
            f" **SN ONT NUEVO:** \n"
            f" **ID ONT ANTIGUA:** \n"
            f" **PRODUCT ID ONT NUEVA:** \n"
            f" **DIRECCION:** {Zona}\n"
            f" **CONDOMINIO/PISO:** \n"
            f" **DISTRITO: {distrito}** \n"
            f" **CTO:** \n"
            f" **PUERTO:** \n"
            f" **CONTRATA:** TLI\n"
            f" **TECNICO:** {tecnico}\n"
        ),
        'CAMBIO DE CTO / TRASLADO / CAMBIO DE PUERTO v2': textwrap.dedent(
            f"📋 **Código de cliente:** \n"
            f"📋 **Ticket:** {codigo}\n"
            f"👤 **Nombre del Cliente:** {cliente}\n"
            f"🆔 **DNI:** {dni}\n"
            f"🔖 **SN del ONT Actual:** \n"
            f"🏠 **Dirección:** {Zona} \n"
            f"📍 **CTO:** \n"
            f"🔌 **Puerto:** \n"
            f"📋️ **Plan:** \n"
            f"📋 **Contrata:** \n"
            f"🛠️ **Técnico:** {tecnico} \n"
            f"📍 **Distrito:** {distrito}\n"
            f"📝 **Observación:** \n"
            f"🛠️ **Motivo del Cambio:** \n"
        ),
        'SPLITTER': textwrap.dedent(
           f"**Foto de la NAP/CTO cerrada**\n"
           f"**TICKET:** {ticket}\n"
           f" **CODIGO:** {codigo}\n"
           f"**CLIENTE COLOCADO EN SPLITTER:** {cliente}\n"
           f"**DNI DEL CLIENTE COLOCADO:** {dni}\n"
           f"-------------------------------------\n"
           f"**DNI DEL CLIENTE AFECTADO:**\n"
           f"**CTO:**\n"
           f"**COORDENADAS CTO:**\n"
           f"**PUERTO UTILIZADO:**\n"
           f"-------------------------------------\n"
           f"**POTENCIA EN EL PUERTO:**\n"
           f"**POTENCIA EN EL PUERTO CON SPLITTER:**\n"
           f"**CONTRATA:** TLI\n"
          ),

        'CAMBIO DE MESH': textwrap.dedent(
            f" **Foto de la MAC del MESH nuevo y antiguo**\n"
            f" **NOMBRE DEL CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **SN MESH NUEVO:** \n"
            f" **MAC MESH NUEVO:** \n"
            f" **SN MESH ACTUAL:** \n"
            f" **MAC MESH ACTUAL:** \n"
        ),
        'ENTREGA DE MESH': textwrap.dedent(
            f" **Foto de la MAC del MESH actual**\n"
            f" **NOMBRE DEL CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **SN MESH:** \n"
            f" **MAC MESH:** \n"
        ),
        'ENTREGA DE TELEFONO': textwrap.dedent(
            f" **Foto de la MAC del teléfono actual**\n"
            f" **NOMBRE DEL CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **MAC TELEFONO:** \n"
        ),
        'CAMBIO DE TELEFONO': textwrap.dedent(
            f" **Foto de la MAC del teléfono nuevo y antiguo**\n"
            f" **NOMBRE DEL CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **MAC TELEFONO NUEVO:** \n"
            f" **MAC TELEFONO ACTUAL:** \n"
        ),
         'ENTREGA DE MESH v2': textwrap.dedent(
            f"📷 **Foto de la MAC del Mesh Actual**\n"
            f"👤 **Nombre del cliente:** {cliente}\n"
            f"🆔 **DNI:** {dni}\n"
            f"🔄 **Mac Teléfono Nuevo:** \n"
        ),
         'CAMBIO DE WINBOX': textwrap.dedent(
            f" **Foto de la MAC del Winbox nuevo y antiguo**\n"
            f" **NOMBRE DEL CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **SN WINBOX NUEVO:** \n"
            f" **MAC WINBOX NUEVO:** \n"
            f" **SN WINBOX NUEVO** \n"
            f" **MAC WINBOX NUEVO:** \n"
        ),
         'ENTREGA DE WINBOX': textwrap.dedent(
            f" **Foto de la mac del winbox actual**\n"
            f" **NOMBRE DEL CLIENTE:** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **CODIGO:** {codigo}\n"
            f" **SN WINBOX ACTUAL:** \n"
            f" **MAC WINBOX:** \n"
        ),
        'PLANTILLA DE CIERRE V1': textwrap.dedent(
            f" **TICKET:** {ticket}\n"
            f" **NOMBRE DEL CLIENTE(A):** {cliente}\n"
            f" **DNI:** {dni}\n"
            f" **TRABAJO REALIZADO:** \n"
            f" **DESCARTES REALIZADOS:** \n"
        ),
    }


    return plantillas.get(tipo_trabajo, "📋 **Información Adicional:** No se dispone de instrucciones específicas.")

# Definición de las categorías y sus respectivas plantillas
categorias = {
    'ENVIADO POR LA CONTRATA': [
        'REGISTRO DE LLEGADA',
        'PLANTILLA DE LLEGADA ALTO VALOR',
        'VALIDACION DE LLAMADAS',
        'ORDEN CON RESERVA'
    ],
    'SOLICITANDO AUTORIZACION': [
        'AUTORIZACIÓN PARA RECABLEADO O TRASLADO',
        'UTILIZACION DE PUERTO DE BAJA',
        'ALERTA POR CLIENTE DESCONECTADO',
        'VALIDAR CASOS CON COSTO / CAMBIO DE TICKET',
        'SUPERVISOR',
        'REPORTE PEXT',
        'CAMBIO DE ONT'
    ],
    'SOLICITANDO ACTIVACIONES': [
        'CAMBIO DE ONT v2',
        'ACTIVACIÓN',
        'CAMBIO DE CTO / CAMBIO DE PUERTO / TRASLADO',
        'TRASLADO + CAMBIO DE ONT',
        'REMATRICULACION'
    ],
    'ATIPICOS': [
        'SPLITTER',
        'CAMBIO DE MESH',
        'ENTREGA DE MESH',
        'CAMBIO DE TELEFONO',
        'ENTREGA DE TELEFONO',
        'CAMBIO DE WINBOX',
        'ENTREGA DE WINBOX'
    ],
    'PLANTILLA DE CIERRE': [
        'PLANTILLA DE CIERRE V1'
    ]
}


@bot.message_handler(commands=['vt'])
def buscar_orden(message):
    global usuarios_df
    user_id = message.from_user.id

    try:
        # Verificar acceso
        response = requests.post(API_VALIDAR_USUARIO, json={"user_id": user_id}, timeout=5)
        data = response.json()
        print(f"🔍 Respuesta de la API para user_id {user_id}: {data}")

        if not data.get("permitido"):
            bot.reply_to(message, "⛔ No tienes permiso para usar este bot. Contacta a soporte.")
            return
        if not data.get("asistencia_marcada"):
            bot.reply_to(message, "⚠️ Debes marcar asistencia con /asistencia antes de usar el bot.")
            return
        if data.get("estado_asistencia") in ["Pendiente", "Rechazado"]:
            bot.reply_to(message, f"⏳ Tu solicitud de asistencia está en estado: {data.get('estado_asistencia')}.")
            return
        if data.get("estado_asistencia") != "Acceso":
            bot.reply_to(message, "⛔ No tienes acceso en este momento. Contacta a soporte.")
            return

        # Siempre cargar Excel actualizado
        from estado_global import cargar_estado
        _, ruta = cargar_estado()
        if ruta and os.path.exists(ruta):
            fila_inicio = detectar_fila_inicio(ruta)
            if fila_inicio:
                df = pd.read_excel(ruta, skiprows=fila_inicio - 1)
                df.columns = df.columns.str.strip().str.replace('\n', '').str.replace('\r', '')
                usuarios_df = df
                print(f"✅ Excel recargado desde: {ruta}")
            else:
                bot.reply_to(message, "⚠️ No se encontró la fila de inicio en el Excel.")
                return
        else:
            bot.reply_to(message, "⚠️ No hay archivo Excel cargado. Usa /subir para cargar uno.")
            return

        # Buscar orden
        try:
            ordenid = message.text.split()[1]
            print(f"👤 Usuario {user_id} busca orden {ordenid}")

            resultado = usuarios_df[usuarios_df['CodiSeguiClien'] == int(ordenid)]

            if resultado.empty:
                bot.reply_to(message, "⚠️ No se encontró ninguna orden con ese ID.")
                return

            codi_segui_clien = resultado['CodiSeguiClien'].values[0]
            markup = crear_teclado_categorias(ordenid)

            # Mostrar nombre del archivo actual
            nombre_excel = os.path.basename(ruta) if ruta else "desconocido"

            bot.send_message(
                message.chat.id,
                f"🔍 **CodiSeguiClien Seleccionado:** {codi_segui_clien}\n"
                f"📁 *Origen:* `{nombre_excel}`\n"
                f"📋 **Selecciona una categoría:**",
                reply_markup=markup,
                parse_mode='Markdown'
            )

        except IndexError:
            bot.reply_to(message, "⚠️ Proporcione un OrdenId. Ejemplo: /vt 1617625")
        except ValueError:
            bot.reply_to(message, "⚠️ OrdenId no válido.")
        except KeyError:
            bot.reply_to(message, "⚠️ La columna 'CodiSeguiClien' no se encuentra en el archivo Excel.")

    except requests.exceptions.RequestException as e:
        bot.reply_to(message, "⚠️ Error al verificar acceso. Inténtalo más tarde.")
        print(f"❌ Error en la API de validación: {e}")


def crear_teclado_categorias(codi_segui_clien):
    markup = types.InlineKeyboardMarkup()

    for categoria in categorias.keys():
        callback_data = f"{codi_segui_clien}|{categoria}"
        markup.add(types.InlineKeyboardButton(text=categoria, callback_data=callback_data))

    return markup

def obtener_codi_segui_clien(ordenid):
    resultado = usuarios_df[usuarios_df['CodiSeguiClien'] == int(ordenid)]
    if not resultado.empty:
        return resultado['CodiSeguiClien'].values[0]
    return None

@bot.callback_query_handler(func=lambda call: call.data.split('|')[1] in categorias.keys())
def categoria_seleccionada(call):
    try:
        ordenid, categoria = call.data.split('|')
        opciones = categorias[categoria]
        codi_segui_clien = obtener_codi_segui_clien(ordenid)

        markup = types.InlineKeyboardMarkup()
        for opcion in opciones:
            callback_data = f"{codi_segui_clien}|{opcion}"
            markup.add(types.InlineKeyboardButton(text=opcion, callback_data=callback_data))

        bot.send_message(
            call.message.chat.id,
            f"📋 **Categoría Seleccionada:** {categoria}\n"
            "📋 **¿Qué plantilla deseas usar?**",
            reply_markup=markup,
            parse_mode='Markdown'
        )
    except IndexError:
        bot.send_message(call.message.chat.id, "⚠️ Datos de llamada no válidos.")
    except KeyError:
        bot.send_message(call.message.chat.id, "⚠️ La categoría seleccionada no existe.")
    except Exception as e:
        print(f"Error: {e}")
        bot.send_message(call.message.chat.id, "⚠️ Ocurrió un error al seleccionar la categoría.")
geolocator = Nominatim(user_agent="ninapro")
# Función para obtener el distrito a partir de coordenadas
def obtener_distrito(latitud, longitud):
    try:
        location = geolocator.reverse((latitud, longitud), exactly_one=True)
        if location:
            # Obtener la dirección completa
            address = location.raw['address']
            # Retornar el distrito
            return address.get('suburb', 'Distrito no encontrado')
    except Exception as e:
        return f"Error: {str(e)}"


def safe_str(value):
    # Asegura que el valor sea siempre un string limpio
    return str(value).strip() if value else 'N/D'
    
def escape_markdown_v2(text):
    caracteres_especiales = r"*_[]()~`>#+-=|{}.!"
    for char in caracteres_especiales:
        text = text.replace(char, f"\\{char}")  # Escapar con '\'
    return text

@bot.callback_query_handler(func=lambda call: '|' in call.data)
def plantilla_seleccionada(call):
    try:
        # Separar ID de orden y tipo de trabajo desde callback_data
        ordenid, tipo_trabajo = call.data.split('|')

        # Confirmar la plantilla seleccionada al usuario
        bot.send_message(call.message.chat.id, f"📝 Plantilla seleccionada: *{tipo_trabajo}*", parse_mode='Markdown')

        # Convertir ordenid a número de forma segura
        ordenid = int(float(ordenid))  # ✅ Solución aplicada

        # Buscar datos del cliente con el ID de la orden
        resultado = usuarios_df[usuarios_df['CodiSeguiClien'] == ordenid]

        if resultado.empty:
            bot.send_message(call.message.chat.id, "⚠️ No se encontraron datos de la orden seleccionada.")
            return

        datos = resultado.iloc[0]

        # Obtener la información del cliente con valores predeterminados
        dni = datos.get('Número Documento', 'N/D')
        cliente = datos.get('Cliente', 'N/D')
        sn_actual = datos.get('Tipo', 'N/D')  # SN ONT Actual

        # Manejo de dirección evitando nulos o vacíos
        direccion = datos.get('Direccion') or datos.get('Direccion1', 'N/D')
        direccion = direccion.split('||REFERENCIA:')[0].strip()

        # Función para validar números y evitar errores con NaN
        def safe_int(value):
            return str(int(value)) if pd.notna(value) and value != 0 else 'N/D'

        # Función para manejar cadenas de forma segura
        def safe_str(value):
            return str(value) if pd.notna(value) else 'N/D'

        # Manejo seguro de los teléfonos
        telefono_movil = safe_str(datos.get('TeleMovilNume'))
        telefono_fijo = safe_str(datos.get('TeleFijoNume'))
        if telefono_fijo.endswith('.0'):
             telefono_fijo = telefono_fijo[:-2]
        telefono = ' - '.join(filter(lambda x: x != 'N/D', [telefono_movil, telefono_fijo])) or 'N/D'

        # Obtener otros datos con manejo de valores faltantes
        producto = datos.get('Producto', 'N/D')
        cuadrilla = datos.get('Cuadrilla', 'N/D')
        motivo_trabajo = datos.get('Motivo Trabajo', 'N/D')
        estado = datos.get('Estado', 'N/D')
        region = datos.get('Region', 'N/D')
        motivo_regestion = datos.get('Motivo Regestión', 'N/D')
        tecnico = datos.get('Cuadrilla', 'N/D')
        zona = datos.get('Zona', 'N/D')
        ticket = datos.get('CodiSegui', 'N/D')
        codigo = datos.get('CodiSeguiClien', 'N/D')
        ot = datos.get('OrdenId', 'N/D')

        # Obtener la georreferencia de manera segura
        georeferencia = datos.get('Georeferencia', '0.0,0.0')
        try:
            latitud, longitud = map(float, georeferencia.split(','))
        except ValueError:
            latitud, longitud = 0.0, 0.0  # Valores predeterminados

        distrito = obtener_distrito(latitud, longitud) or 'N/D'

        # Enviar mensaje de "Cargando plantilla..."
        loading_message = bot.send_message(call.message.chat.id, "🔄 Cargando plantilla...")

        mensaje_plantilla = seleccionar_plantilla(
          tipo_trabajo, dni, ordenid, cliente, sn_actual, direccion,
          producto, cuadrilla, motivo_trabajo, estado, region,
          motivo_regestion, distrito, tecnico, zona, telefono, codigo, ticket, ot
        )
        # Eliminar los `**` del mensaje
        mensaje_plantilla = mensaje_plantilla.replace("**", "")
        # Escapar caracteres para evitar errores de MarkdownV2
        mensaje_plantilla = escape_markdown_v2(mensaje_plantilla)

        # Editar el mensaje con MarkdownV2
        bot.edit_message_text(
         chat_id=call.message.chat.id,
         message_id=loading_message.message_id,
         text=mensaje_plantilla,
         parse_mode='MarkdownV2'
        )

    except ValueError as e:
        bot.send_message(call.message.chat.id, f"⚠️ Hubo un problema con los datos recibidos: {str(e)}")



@bot.message_handler(commands=['start'])
def enviar_bienvenida(message):
    user_id = message.from_user.id

    try:
        # 1️⃣ Consultar la API para validar usuario y asistencia
        response = requests.post(API_VALIDAR_USUARIO, json={"user_id": user_id}, timeout=5)
        data = response.json()

        # 📌 Imprimir en consola todo lo que devuelve la API para depuración
        print(f"🔍 Respuesta de la API para user_id {user_id}: {data}")

        if not data.get("permitido"):
            bot.reply_to(message, "⛔ No tienes permiso para usar este bot. Contacta a soporte.")
            return  

        # 2️⃣ Verificar si el usuario marcó asistencia
        if not data.get("asistencia_marcada"):
            bot.reply_to(message, "⚠️ Debes marcar asistencia con /asistencia antes de usar el bot.")
            return  

        # 3️⃣ Verificar si la asistencia fue aprobada
        estado_asistencia = data.get("estado_asistencia", "Pendiente")

        if estado_asistencia == "Pendiente":
            bot.reply_to(message, "⏳ Tu solicitud de asistencia está en revisión. Espera a que sea aprobada antes de continuar.")
            return  

        if estado_asistencia == "Rechazado":
            bot.reply_to(message, "❌ Tu solicitud de asistencia fue rechazada. Contacta a tu gestora para más información.")
            return  

        if estado_asistencia != "Acceso":
            bot.reply_to(message, "⛔ No tienes acceso en este momento. Contacta a soporte.")
            return  

        # ✅ Si la asistencia fue aprobada, permitir acceso
        mensaje = (
            "✅ ¡Bienvenido al bot de seguimiento de órdenes! 🎉\n\n"
            "Aquí puedes buscar información sobre tus órdenes utilizando el comando:\n"
            "/vt [CodiSeguiClien] \n\n"
            "👤 Información del creador: /creador\n"
        )
        bot.reply_to(message, mensaje)

    except requests.exceptions.RequestException as e:
        bot.reply_to(message, "⚠️ Error al verificar acceso. Inténtalo más tarde.")
        print(f"❌ Error en la API de validación: {e}")


@bot.message_handler(commands=['vt'])
def buscar_orden(message):
    try:
        ordenid = message.text.split()[1]
        info = buscar_por_ordenid(ordenid)
        agregar_registro(f"Consulta de OrdenId: {CodiSeguiClien}")
        bot.reply_to(message, info, parse_mode='Markdown')
    except IndexError:
        bot.reply_to(message, "⚠️ Proporcione un OrdenId. Ejemplo: /vt 1617625")
    except ValueError:
        bot.reply_to(message, "⚠️ OrdenId no válido.")


@bot.message_handler(commands=['creador'])
def mostrar_creador(message):
    creador_info = "👤 Este bot fue creado por NinaProgramming. ¡Gracias por usarlo!"
    bot.reply_to(message, creador_info)

@bot.message_handler(commands=['ayuda'])
def mostrar_ayuda(message):
    ayuda = (
        "🆘 Aquí tienes algunos comandos disponibles:\n"
        "/vt [OrdenId] - Busca información sobre una orden.\n"
        "/creador - Muestra información sobre el creador del bot.\n"
        "/ayuda - Muestra este mensaje de ayuda."
    )
    bot.reply_to(message, ayuda)
    
@bot.message_handler(commands=['info'])
def mostrar_info(message):
    global usuarios_df  # Acceder al DataFrame global

    estado, _ = cargar_estado()  # Lee el estado desde JSON

    # Contar cuántas órdenes hay cargadas (filas del Excel)
    total_ordenes = len(usuarios_df) if not usuarios_df.empty else 0

    info = (
        f"📄 Estado del Excel:\n{estado}\n"
        f"📦 Órdenes cargadas: {total_ordenes}"
    )

    bot.reply_to(message, info)



def actualizar_estado(texto, color):
    # Asegúrate de que 'estado_label' esté definido en tu interfaz
    estado_label.config(text=f"Estado: {texto}", foreground=color)

def iniciar_bot():
    global bot_activo
    if not bot_activo:
        bot_activo = True
        threading.Thread(target=bot_polling_con_reintento).start()
        actualizar_estado("Activo 🟢", "#00ff00")

def detener_bot():
    global bot_activo
    if bot_activo:
        bot.stop_polling()
        bot_activo = False
        actualizar_estado("Detenido 🔴", "red")

def bot_polling_con_reintento():
    global bot_activo  # Asegúrate de declarar bot_activo como global aquí
    while bot_activo:
        try:
            bot.polling(none_stop=True)  # Iniciar el polling de forma continua
        except Exception as e:
            print(f"Error en la conexión: {e}")
            actualizar_estado("Reconectando... ⏳", "orange")
            time.sleep(5)  # Espera 5 segundos antes de intentar reconectar
        else:
            # Si el polling finaliza sin errores, reinicia el proceso
            bot_activo = False
            actualizar_estado("Detenido 🔴", "red")
            break
        # Cuando se recupere la conexión, actualizar el estado
        if bot_activo:
            actualizar_estado("Reconectado 🟢", "#00ff00")


# Zona horaria de Lima
tz_lima = pytz.timezone('America/Lima')
# Usuario al que se enviará el mensaje a las 10 PM
USER_ID = 5540982553
@bot.message_handler(commands=['estado'])
def ver_estado(message):
    """Muestra el estado del bot, el tiempo activo y detalles importantes."""
    chat_id = message.chat.id
    estado_bot = "🟢 Encendido" if bot_activo else "🔴 Apagado"

    if bot_activo:
        ahora = datetime.now(tz_lima)
        tiempo_activado = ahora - inicio_bot
        tiempo_activado_str = str(tiempo_activado).split(".")[0]  # Sin milisegundos
        mensaje_tiempo = f"⏳ Tiempo activo: {tiempo_activado_str}"
        mensaje_inicio = f"🕒 Iniciado en: {inicio_bot.strftime('%Y-%m-%d %H:%M:%S')}"
    else:
        mensaje_tiempo = "⏳ Tiempo activo: No disponible"
        mensaje_inicio = "🕒 No disponible"

    python_version = platform.python_version()
    pid = os.getpid()
    memoria = psutil.virtual_memory().percent
    cpu = psutil.cpu_percent(interval=1)

    mensaje = (
        f"🤖 Estado del bot: {estado_bot}\n"
        f"{mensaje_inicio}\n"
        f"{mensaje_tiempo}\n"
        f"📂 Estado de Excel: {estado_excel}\n\n"
        f"🖥️ Python {python_version} | PID: {pid}\n"
        f"💾 RAM usada: {memoria}% | CPU: {cpu}%"
    )

    bot.send_message(chat_id, mensaje)

hora_programada = time(22, 0)  # 10:00 PM

@bot.message_handler(commands=['cuantofalta'])
def cuanto_falta(message):
    """Calcula cuánto falta para el mensaje programado de las 10 PM y muestra la hora exacta."""
    ahora = datetime.now(tz_lima)
    hora_objetivo = ahora.replace(hour=hora_programada.hour, minute=hora_programada.minute, second=0, microsecond=0)

    if ahora >= hora_objetivo:
        # Si ya pasó la hora objetivo, calcular para el día siguiente
        hora_objetivo += datetime.timedelta(days=1)

    tiempo_restante = hora_objetivo - ahora
    horas, segundos_restantes = divmod(tiempo_restante.total_seconds(), 3600)
    minutos, _ = divmod(segundos_restantes, 60)

    mensaje = (
        f"⏳ Falta {int(horas)} horas y {int(minutos)} minutos para el mensaje programado.\n"
        f"🕒 Hora programada: {hora_objetivo.strftime('%I:%M %p')} (10:00 PM Lima)"
    )
    bot.send_message(message.chat.id, mensaje)


# Crear la aplicación Flask
app = Flask(__name__)

@app.route("/", methods=["GET"])
def home():
    return "🤖 Bot funcionando localmente con Long Polling 🚀"

@app.route(f"/{TOKEN}", methods=["POST"])
def webhook():
    """Recibe actualizaciones del Webhook y las procesa."""
    update = request.get_json()
    if update:
        bot.process_new_updates([telebot.types.Update.de_json(update)])
    return "OK", 200

if __name__ == "__main__":
    print("🚀 Iniciando bot con Long Polling localmente")

    while True:
        try:
            bot.polling(none_stop=True, timeout=60, long_polling_timeout=60)
        except telebot.apihelper.ApiException as e:
            print(f"⚠️ Error en la API de Telegram: {e}")
        except Exception as e:
            print(f"⚠️ Error inesperado: {e}")
        
        print("🔄 Reintentando en 5 segundos...")
        tm.sleep(5)  # ✅ Corrección: usar `tm.sleep(5)` en lugar de `datetime.datetime.sleep(5)`

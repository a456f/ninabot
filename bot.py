import os
import time
import platform
import threading
import traceback
import telebot
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

import estado_global
from main import detectar_fila_inicio, enviar_datos_a_api

TOKEN_2 = '7922512452:AAGhfzYMzhJPfV1TA1dBy2w6hICCIXHdNds'
bot2 = telebot.TeleBot(TOKEN_2)

if platform.system() == "Windows":
    DOWNLOAD_FOLDER = os.path.join(os.getcwd(), "descargas")
else:
    DOWNLOAD_FOLDER = "/mnt/data"
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_FOLDER,
    "download.prompt_for_download": False,
    "profile.default_content_settings.popups": 0,
    "directory_upgrade": True
}
options.add_experimental_option("prefs", prefs)
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

modo_activo_2 = False
chat_id_global_2 = None
usuarios_autorizados_2 = {}
CLAVE_ENCENDER_2 = "185946"
CLAVE_APAGAR_2 = "4582"

def barra(seg, total=10, largo=10):
    llenado = int((seg / total) * largo)
    porcentaje = int((seg / total) * 100)
    return '🟦' * llenado + '⬜️' * (largo - llenado) + f' ({porcentaje}%)'

def actualizar_mensaje(bot, chat_id, msg_id, estado_actual, barra_progreso=""):
    pasos = {
        1: "🔑 Iniciando sesión...",
        2: "📁 Entrando al módulo...",
        3: "🔍 Filtrando información...",
        4: "📄 Exportando archivo...",
        5: "⏳ Descargando archivo...",
        6: "✅ Finalizando proceso..."
    }
    texto = f"🚀 Ejecutando...\n\n{pasos.get(estado_actual, '⏳ Procesando...')} 🔄"
    if barra_progreso:
        texto += f"\n{barra_progreso}"
    try:
        bot.edit_message_text(texto.strip(), chat_id, msg_id)
    except:
        pass

def obtener_ultimo_archivo_xlsx(folder, segundos_max=60):
    ahora = time.time()
    archivos = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith(".xlsx")]
    archivos_recientes = [f for f in archivos if ahora - os.path.getmtime(f) < segundos_max]
    return sorted(archivos_recientes, key=os.path.getmtime, reverse=True)[0] if archivos_recientes else None

[Estado]
def bucle_automatico_2():
    while True:
        if modo_activo_2 and chat_id_global_2:
            hora_actual = datetime.now().hour
            if 7 <= hora_actual < 20:
                try:
                    print("[INFO] Ejecutando automático...")
                    bot2.send_message(chat_id_global_2, "⏳ Iniciando proceso automático...")
                    exportar_y_enviar_2(chat_id_global_2)
                    bot2.send_message(chat_id_global_2, "✅ Proceso automático terminado.")
                except Exception as e:
                    print(f"[ERROR] bucle_automatico_2: {e}")
                    bot2.send_message(chat_id_global_2, f"⚠️ Error en automático:\n{e}")
            else:
                print("[INFO] Fuera de horario (7:00 a.m. a 8:00 p.m.). Esperando...")
        time.sleep(300)

@bot2.message_handler(commands=['estadoexcel'])
def estado_excel_handler(msg):
    estado, _ = estado_global.cargar_estado()
    bot2.send_message(msg.chat.id, estado)

@bot2.message_handler(commands=['exportar'])
def exportar_handler(msg):
    try:
        exportar_y_enviar_2(msg.chat.id)
    except Exception as e:
        error = traceback.format_exc()
        bot2.send_message(msg.chat.id, f"❌ Error inesperado:\n{e}")

@bot2.message_handler(commands=['encender'])
def encender_handler(msg):
    global modo_activo_2, chat_id_global_2
    if modo_activo_2:
        bot2.send_message(msg.chat.id, "⚠️ El bot ya está ENCENDIDO.")
        return
    if msg.chat.id not in usuarios_autorizados_2:
        bot2.send_message(msg.chat.id, "🔐 Envía la clave para activar el modo automático.")
        return
    modo_activo_2 = True
    chat_id_global_2 = msg.chat.id
    bot2.send_message(msg.chat.id, "✅ Modo automático ACTIVADO.")

@bot2.message_handler(commands=['apagar'])
def apagar_handler(msg):
    global modo_activo_2, chat_id_global_2
    if not modo_activo_2:
        bot2.send_message(msg.chat.id, "⚠️ El bot ya está APAGADO.")
        return
    if msg.chat.id not in usuarios_autorizados_2:
        bot2.send_message(msg.chat.id, "🔐 Envía la clave para apagar.")
        return
    modo_activo_2 = False
    chat_id_global_2 = None
    usuarios_autorizados_2.pop(msg.chat.id, None)
    bot2.send_message(msg.chat.id, "🛑 Modo automático DESACTIVADO.")

@bot2.message_handler(commands=['estado'])
def estado_handler(msg):
    estado = "✅ El bot está *ENCENDIDO*." if modo_activo_2 else "❌ El bot está *APAGADO*."

    # Obtener nombre del usuario que activó el bot
    if modo_activo_2 and chat_id_global_2:
        try:
            usuario = bot2.get_chat(chat_id_global_2)
            nombre_usuario = f"{usuario.first_name or ''} {usuario.last_name or ''}".strip()
            usuario_info = f"\n👤 *Activado por:* {nombre_usuario} (`{chat_id_global_2}`)"
        except Exception:
            usuario_info = f"\n👤 *Activado por:* `{chat_id_global_2}`"
    else:
        usuario_info = ""

    horario = "\n🕒 *Horario de funcionamiento:* 7:00 a.m. a 8:00 p.m."

    comandos = """
📦 *Comandos disponibles:*
/exportar - Ejecutar exportación manual
/encender - Activar modo automático
/apagar - Desactivar modo automático
/estado - Ver estado del bot
/estadoexcel - Ver estado del archivo Excel
"""

    bot2.send_message(
        msg.chat.id,
        f"{estado}{usuario_info}{horario}\n{comandos}",
        parse_mode="Markdown"
    )


@bot2.message_handler(func=lambda m: True)
def clave_handler(msg):
    global modo_activo_2, chat_id_global_2
    if msg.text == CLAVE_ENCENDER_2:
        usuarios_autorizados_2[msg.chat.id] = True
        modo_activo_2 = True
        chat_id_global_2 = msg.chat.id
        bot2.send_message(msg.chat.id, "✅ Clave correcta. Modo automático ACTIVADO.")
    elif msg.text == CLAVE_APAGAR_2:
        if msg.chat.id in usuarios_autorizados_2:
            modo_activo_2 = False
            chat_id_global_2 = None
            usuarios_autorizados_2.pop(msg.chat.id, None)
            bot2.send_message(msg.chat.id, "🛑 Clave correcta. Bot APAGADO.")
        else:
            bot2.send_message(msg.chat.id, "🔐 No estás autorizado para apagar el bot.")
    else:
        if msg.chat.id not in usuarios_autorizados_2:
            bot2.send_message(msg.chat.id, "❌ Clave incorrecta.")

threading.Thread(target=bucle_automatico_2, daemon=True).start()
print("🤖 Segundo bot ejecutándose...")
bot2.polling()

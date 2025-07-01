import os
import time
import platform
import threading
import traceback
from datetime import datetime, timedelta
import telebot
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pytz

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


def obtener_fecha_filtrado():
    zona_lima = pytz.timezone("America/Lima")
    ahora = datetime.now(zona_lima)
    if ahora.hour < 7:
        ahora -= timedelta(days=1)
    return ahora.strftime("%d/%m/%Y")

def obtener_ultimo_archivo_xlsx(folder, segundos_max=60):
    ahora = time.time()
    archivos = [os.path.join(folder, f) for f in os.listdir(folder) if f.endswith(".xlsx")]
    archivos_recientes = [f for f in archivos if ahora - os.path.getmtime(f) < segundos_max]
    return sorted(archivos_recientes, key=os.path.getmtime, reverse=True)[0] if archivos_recientes else None

def esperar_descarga_completa(filepath, timeout=30):
    for _ in range(timeout):
        if os.path.exists(filepath) and not filepath.endswith(".crdownload"):
            try:
                with open(filepath, "rb"):
                    if os.path.getsize(filepath) > 10 * 1024:  # mínimo 10 KB
                        return True
            except Exception:
                pass
        time.sleep(1)
    return False



def exportar_y_enviar_2(chat_id):
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 30)
    progreso_msg = bot2.send_message(chat_id, "📱 Iniciando proceso...")

    hora_inicio = hora_actual_lima()

    try:
        # Login
        driver.get("https://winbo-phx.azurewebsites.net/login.aspx")
        wait.until(EC.presence_of_element_located((By.ID, "txtUsuario"))).send_keys("brubio")
        wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys("M123456789")
        driver.find_element(By.ID, "BtnLoginInicial").click()

        for i in range(1, 6):
            actualizar_mensaje(bot2, chat_id, progreso_msg.message_id, 1, barra(i, 5))
            time.sleep(0.25)

        # Abrir módulo
        wait.until(EC.presence_of_element_located((By.ID, "menuSistema")))
        driver.execute_script("AbrirPagi('Paginas/OperadoresBO/misOrdenes.aspx?to=1&nombre=Seguimiento+de+Ordenes&id=74&icono=&edit=S','74');")

        for i in range(1, 6):
            actualizar_mensaje(bot2, chat_id, progreso_msg.message_id, 2, barra(i, 5))
            time.sleep(0.25)

        # Filtrar por fecha actual
        hoy = obtener_fecha_filtrado()
        wait.until(EC.presence_of_element_located((By.ID, "txtDesdeFechaVisi74")))
        wait.until(EC.presence_of_element_located((By.ID, "txtHastaFechaVisi74")))
        driver.execute_script(f"document.getElementById('txtDesdeFechaVisi74').value = '{hoy}'")
        driver.execute_script(f"document.getElementById('txtHastaFechaVisi74').value = '{hoy}'")

        filtrar_btn = wait.until(EC.presence_of_element_located((By.ID, "BtnFiltrar74")))
        driver.execute_script("arguments[0].click();", filtrar_btn)

        for i in range(1, 11):
            actualizar_mensaje(bot2, chat_id, progreso_msg.message_id, 3, barra(i))
            time.sleep(0.35)

        # Exportar
        exportar_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(., 'Exportar')]")))
        driver.execute_script("arguments[0].click();", exportar_btn)

        # Esperar 5 segundos para que se genere el archivo
        time.sleep(5)

        for i in range(1, 11):
            actualizar_mensaje(bot2, chat_id, progreso_msg.message_id, 4, barra(i))
            time.sleep(0.35)

        # Abrir notificaciones y descargar archivo
        driver.execute_script("arguments[0].click();", wait.until(EC.presence_of_element_located((By.ID, "spnNotiCampa"))))
        time.sleep(2)

        enlaces = driver.find_elements(By.XPATH, "//p[@class='noti-text']/a[contains(@href, '.xlsx')]")
        if not enlaces:
            bot2.edit_message_text("❌ No se encontró ningún archivo .xlsx.", chat_id, progreso_msg.message_id)
            return

        url_archivo = enlaces[0].get_attribute("href")
        driver.get(url_archivo)

        filename = os.path.basename(url_archivo)
        ruta_descarga = os.path.join(DOWNLOAD_FOLDER, filename)

        if not esperar_descarga_completa(ruta_descarga):
            bot2.edit_message_text("❌ La descarga del archivo no se completó correctamente.", chat_id, progreso_msg.message_id)
            return

        fila_inicio = detectar_fila_inicio(ruta_descarga)
        if fila_inicio is None:
            raise ValueError("No se encontró la fila de inicio.")

        df = pd.read_excel(ruta_descarga, skiprows=fila_inicio - 1, engine="openpyxl")
        df.columns = df.columns.str.strip()

        if df.empty:
            bot2.edit_message_text("⚠️ El archivo exportado está vacío o mal estructurado.", chat_id, progreso_msg.message_id)
            return

        estado_global.guardar_estado(f"📁 Archivo descargado automáticamente: {filename}", ruta_descarga)
        enviar_datos_a_api(df)

        hora_fin = hora_actual_lima()
        duracion = hora_fin - hora_inicio

        # Mensaje final al usuario
        bot2.edit_message_text(
            f"✅ Archivo exportado y procesado correctamente.\n"
            f"📎 Nombre: `{filename}`\n"
            f"📅 Fecha filtrada: {hoy}\n"
            f"🕒 Inicio: {hora_inicio.strftime('%H:%M:%S')}\n"
            f"🕓 Fin: {hora_fin.strftime('%H:%M:%S')}\n"
            f"⏱️ Duración total: {str(duracion).split('.')[0]}",
            chat_id, progreso_msg.message_id, parse_mode="Markdown"
        )

        # Enviar datos del proceso a la API
        proxima_actualizacion = (hora_fin + timedelta(seconds=334)).strftime('%H:%M:%S')
        payload = {
            "nombre_archivo": filename,
            "fecha_filtrada": hoy,
            "hora_inicio": hora_inicio.strftime('%H:%M:%S'),
            "hora_fin": hora_fin.strftime('%H:%M:%S'),
            "duracion": str(duracion).split('.')[0],
            "proxima_actualizacion": proxima_actualizacion
        }

        try:
            response = requests.post(
                "https://cybernovasystems.com/prueba/sistema_tlc/modelos/telegran/api_guardar_exportacion.php",
                json=payload,
                timeout=10
            )
            print("[INFO] Registro exportación enviado. Código:", response.status_code)
            print("[INFO] Respuesta:", response.text)
        except Exception as e:
            print("[ERROR] Falló envío a API exportación:", e)

    except Exception as e:
        error = traceback.format_exc()
        print(f"[ERROR] exportar_y_enviar_2:\n{error}")
        bot2.edit_message_text(f"⚠️ Error durante el proceso:\n{e}", chat_id, progreso_msg.message_id)

    finally:
        driver.quit()



# Función para obtener la hora actual en Lima
def hora_actual_lima():
    zona_lima = pytz.timezone("America/Lima")
    return datetime.now(zona_lima)

def bucle_automatico_2():
    while True:
        try:
            if modo_activo_2 and chat_id_global_2:
                ahora = hora_actual_lima()
                hora_actual = ahora.hour

                print(f"[DEBUG] Hora actual Lima: {ahora.strftime('%Y-%m-%d %H:%M:%S')}")

                if 7 <= hora_actual < 21:
                    print("[INFO] Ejecutando automático...")
                    bot2.send_message(chat_id_global_2, "⏳ Iniciando proceso automático...")
                    exportar_y_enviar_2(chat_id_global_2)
                    bot2.send_message(chat_id_global_2, "✅ Proceso automático terminado.")
                else:
                    print("[INFO] Fuera de horario (7:00 a.m. a 9:00 p.m.). Esperando...")
            else:
                print("[INFO] Modo automático desactivado o chat_id no definido.")
        except Exception as e:
            print(f"[ERROR] bucle_automatico_2: {e}")
            if chat_id_global_2:
                bot2.send_message(chat_id_global_2, f"⚠️ Error en automático:\n{e}")

        time.sleep(300)  # Espera 5 minutos

@bot2.message_handler(commands=['info'])
def info_handler(msg):
    if not modo_activo_2:
        bot2.send_message(msg.chat.id, "❌ El modo automático está apagado.")
        return

    ahora = time.time()
    segundos_restantes = 300 - int(ahora) % 300
    minutos = segundos_restantes // 60
    segundos = segundos_restantes % 60
    texto = f"🕓 Faltan *{minutos}* minutos y *{segundos}* segundos para la siguiente ejecución automática."
    bot2.send_message(msg.chat.id, texto, parse_mode="Markdown")


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

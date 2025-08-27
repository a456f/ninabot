import os
import time
import platform
import threading
import traceback
from datetime import datetime, timedelta
import telebot
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pytz
import multiprocessing

# ===============================
# Configuración zona horaria
# ===============================
os.environ['TZ'] = 'America/Lima'
time.tzset()

import estado_global
from main import detectar_fila_inicio, enviar_datos_a_api

# ===============================
# Configuración Bot y Descargas
# ===============================
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

# ===============================
# Variables Globales
# ===============================
modo_activo_2 = False
chat_id_global_2 = None
usuarios_autorizados_2 = {5540982553: "185946"}
bloqueo_auto = threading.Lock()

# ===============================
# Funciones Auxiliares
# ===============================
def hora_actual_lima():
    zona_lima = pytz.timezone("America/Lima")
    return datetime.now(zona_lima)

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
    ahora = hora_actual_lima()
    if ahora.hour < 1:
        ahora -= timedelta(days=1)
    return ahora.strftime("%d/%m/%Y")

def esperar_descarga_completa(filepath, timeout=30):
    for _ in range(timeout):
        if os.path.exists(filepath) and not filepath.endswith(".crdownload"):
            try:
                with open(filepath, "rb"):
                    if os.path.getsize(filepath) > 10 * 1024:
                        return True
            except:
                pass
        time.sleep(1)
    return False

# ===============================
# Función principal de exportación
# ===============================
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
        time.sleep(5)

        for i in range(1, 11):
            actualizar_mensaje(bot2, chat_id, progreso_msg.message_id, 4, barra(i))
            time.sleep(0.35)

        # Descargar archivo desde notificación
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

        # Mensaje final
        bot2.edit_message_text(
            f"✅ Archivo exportado y procesado correctamente.\n"
            f"📎 Nombre: {filename}\n"
            f"📅 Fecha filtrada: {hoy}\n"
            f"🕒 Inicio: {hora_inicio.strftime('%H:%M:%S')}\n"
            f"🕓 Fin: {hora_fin.strftime('%H:%M:%S')}\n"
            f"⏱️ Duración total: {str(duracion).split('.')[0]}",
            chat_id, progreso_msg.message_id, parse_mode="Markdown"
        )

        # ====== Envío a API con reintentos ======
        payload = {
            "nombre_archivo": filename,
            "fecha_filtrada": hoy,
            "hora_inicio": hora_inicio.strftime('%H:%M:%S'),
            "hora_fin": hora_fin.strftime('%H:%M:%S'),
            "duracion": str(duracion).split('.')[0],
            "proxima_actualizacion": (hora_fin + timedelta(seconds=334)).strftime('%H:%M:%S')
        }

        for intento in range(3):
            try:
                response = requests.post(
                    "https://tliperu.com/prueba/telegran/api_guardar_exportacion.php",
                    json=payload,
                    timeout=10
                )
                print("[INFO] Registro exportación enviado. Código:", response.status_code)
                break
            except Exception as e:
                print(f"[WARNING] Intento {intento+1} falló: {e}")
                time.sleep(5)

    except Exception as e:
        error = traceback.format_exc()
        print(f"[ERROR] exportar_y_enviar_2:\n{error}")
        bot2.edit_message_text(f"⚠️ Error durante el proceso:\n{e}", chat_id, progreso_msg.message_id)
    finally:
        driver.quit()

# ===============================
# Ejecutor con timeout
# ===============================
def ejecutar_con_timeout(func, args=(), timeout=300):
    p = multiprocessing.Process(target=func, args=args)
    p.start()
    p.join(timeout)
    if p.is_alive():
        print(f"[ERROR] Proceso superó {timeout}s, forzando cierre...")
        p.terminate()
        p.join()
        return False
    return True

# ===============================
# Función robusta para enviar mensajes con reintento
# ===============================
def enviar_mensaje_telegram(chat_id, texto, intentos=3, espera=3):
    for i in range(intentos):
        try:
            bot2.send_message(chat_id, texto)
            return True
        except requests.exceptions.SSLError as e:
            print(f"[SSL ERROR] Intento {i+1} al enviar '{texto}': {e}")
            time.sleep(espera)
        except Exception as e:
            print(f"[ERROR] Intento {i+1} al enviar '{texto}': {e}")
            time.sleep(espera)
    print(f"[FATAL] No se pudo enviar mensaje después de {intentos} intentos: {texto}")
    return False


# ===============================
# Bucle Automático Mejorado con reintentos
# ===============================
def bucle_automatico_2():
    global mensaje_buenos_dias_enviado, mensaje_descanso_enviado
    mensaje_buenos_dias_enviado = False
    mensaje_descanso_enviado = False
    ultimo_dia = None

    print("[INFO] Bucle automático iniciado y funcionando...")

    while True:
        try:
            ahora = hora_actual_lima()
            dia_actual = ahora.date()

            # Reinicia banderas al cambiar de día
            if dia_actual != ultimo_dia:
                mensaje_buenos_dias_enviado = False
                mensaje_descanso_enviado = False
                ultimo_dia = dia_actual

            # Calcula próxima ejecución múltiplo de 5 min
            minuto_siguiente = (ahora.minute // 5 + 1) * 5
            if minuto_siguiente >= 60:
                siguiente_hora = ahora.replace(hour=(ahora.hour + 1) % 24, minute=0, second=0, microsecond=0)
            else:
                siguiente_hora = ahora.replace(minute=minuto_siguiente, second=0, microsecond=0)

            espera_segundos = max((siguiente_hora - ahora).total_seconds(), 0)
            print(f"[DEBUG] {ahora.strftime('%Y-%m-%d %H:%M:%S')} - Esperando {espera_segundos}s hasta próxima ejecución...")
            time.sleep(espera_segundos)

            if not (modo_activo_2 and chat_id_global_2):
                print("[INFO] Modo automático inactivo o sin chat definido.")
                continue

            ahora = hora_actual_lima()
            hora_actual = ahora.hour
            minuto_actual = ahora.minute

            # Enviar saludo matutino
            if hora_actual == 7 and minuto_actual == 0 and not mensaje_buenos_dias_enviado:
                if enviar_mensaje_telegram(chat_id_global_2, "☀️ ¡Buen día! Estoy iniciando mi horario de trabajo."):
                    mensaje_buenos_dias_enviado = True

            # Ejecutar dentro del horario
            if 7 <= hora_actual < 21 or (hora_actual == 21 and minuto_actual == 0):
                print(f"[INFO] Ejecutando proceso automático a las {ahora.strftime('%H:%M:%S')}")
                with bloqueo_auto:
                    enviar_mensaje_telegram(chat_id_global_2, "⏳ Iniciando proceso automático...")
                    inicio_proceso = hora_actual_lima()
                    exito = ejecutar_con_timeout(exportar_y_enviar_2, (chat_id_global_2,), 300)
                    fin_proceso = hora_actual_lima()

                    if exito:
                        duracion = fin_proceso - inicio_proceso
                        print(f"[INFO] Proceso automático finalizado en {duracion}")
                        enviar_mensaje_telegram(chat_id_global_2, "✅ Proceso automático terminado.")
                    else:
                        print("[ERROR] Proceso automático cancelado por exceder el tiempo límite.")
                        enviar_mensaje_telegram(chat_id_global_2, "⚠️ Proceso automático cancelado (timeout).")

                # Enviar mensaje de descanso nocturno
                if hora_actual == 21 and minuto_actual == 0 and not mensaje_descanso_enviado:
                    time.sleep(30)
                    if enviar_mensaje_telegram(chat_id_global_2, "🌙 Buen trabajo por hoy. Me retiro a descansar."):
                        mensaje_descanso_enviado = True
            else:
                print("[INFO] Fuera de horario (7:00 a.m. – 9:00 p.m.). Esperando...")

        except Exception as e:
            error = traceback.format_exc()
            print(f"[ERROR] bucle_automatico_2: {error}")
            if chat_id_global_2:
                enviar_mensaje_telegram(chat_id_global_2, f"⚠️ Error en automático:\n{e}")

# ===============================
# Handlers de Telegram
# ===============================
@bot2.message_handler(commands=['info'])
def info_handler(msg):
    if not modo_activo_2:
        bot2.send_message(msg.chat.id, "❌ El modo automático está apagado.")
        return
    ahora = time.time()
    segundos_restantes = 300 - int(ahora) % 300
    bot2.send_message(msg.chat.id, f"🕓 Faltan *{segundos_restantes//60}* minutos y *{segundos_restantes%60}* segundos para la siguiente ejecución automática.", parse_mode="Markdown")

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
    global chat_id_global_2
    chat_id_global_2 = msg.chat.id
    if msg.chat.id not in usuarios_autorizados_2:
        bot2.send_message(msg.chat.id, "❌ Usuario no autorizado.")
        return
    bot2.send_message(msg.chat.id, "🔑 Ingresa la contraseña para ENCENDER el modo automático:")
    bot2.register_next_step_handler(msg, recibir_clave)

def recibir_clave(msg):
    global modo_activo_2
    if msg.text == usuarios_autorizados_2.get(msg.chat.id):
        modo_activo_2 = True
        bot2.send_message(msg.chat.id, "✅ Modo automático ENCENDIDO.")
    else:
        bot2.send_message(msg.chat.id, "❌ Contraseña incorrecta. No se pudo ENCENDER.")

@bot2.message_handler(commands=['apagar'])
def apagar_handler(msg):
    global modo_activo_2, chat_id_global_2
    modo_activo_2 = False
    chat_id_global_2 = None
    bot2.send_message(msg.chat.id, "⚠️ Modo automático APAGADO.")

# ===============================
# Inicializar Hilo Automático
# ===============================
hilo_bucle_2 = threading.Thread(target=bucle_automatico_2, daemon=True)
hilo_bucle_2.start()

# ===============================
# Ejecutar Bot
# ===============================
bot2.infinity_polling()

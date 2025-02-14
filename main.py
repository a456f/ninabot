import os
import datetime
from telebot import TeleBot
from dotenv import load_dotenv
from telebot.types import Message

# Cargar variables de entorno
load_dotenv()
TELEGRAM_BOT_TOKEN: str = os.getenv('TELEGRAM_BOT_TOKEN')

# Inicializar bot
bot = TeleBot(token=TELEGRAM_BOT_TOKEN)


@bot.message_handler(commands=['start'])
def send_welcome(message: Message):
    """ Maneja el comando /start enviando un mensaje de bienvenida. """
    chat_id = message.chat.id
    bot.send_message(chat_id, "👋 ¡Hola! Soy tu bot de Telegram.\nUsa /help para ver los comandos disponibles.")


@bot.message_handler(commands=['help'])
def send_help(message: Message):
    """ Muestra la lista de comandos disponibles. """
    chat_id = message.chat.id
    help_text = (
        "📌 *Comandos disponibles:*\n"
        "/start - Iniciar el bot\n"
        "/help - Mostrar ayuda\n"
        "/about - Información sobre el bot\n"
        "/echo <mensaje> - Repetir un mensaje\n"
        "/time - Mostrar la hora actual\n"
    )
    bot.send_message(chat_id, help_text, parse_mode="Markdown")


@bot.message_handler(commands=['about'])
def send_about(message: Message):
    """ Información sobre el bot. """
    chat_id = message.chat.id
    bot.send_message(chat_id, "🤖 Soy un bot de prueba creado con Python y Telebot.")


@bot.message_handler(commands=['echo'])
def echo_text(message: Message):
    """ Repite el mensaje enviado después del comando /echo. """
    chat_id = message.chat.id
    text = message.text.split(" ", 1)  # Dividir el mensaje después del comando
    if len(text) > 1:
        bot.send_message(chat_id, f"🔄 {text[1]}")
    else:
        bot.send_message(chat_id, "⚠️ Debes escribir un mensaje después de /echo")


@bot.message_handler(commands=['time'])
def send_time(message: Message):
    """ Muestra la hora actual. """
    chat_id = message.chat.id
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    bot.send_message(chat_id, f"⏰ Hora actual: {now}")


@bot.message_handler(func=lambda message: True)
def echo_message(message: Message):
    """ Repite cualquier mensaje enviado. """
    chat_id = message.chat.id
    bot.send_message(chat_id, message.text)


# Mantener el bot en ejecución
bot.infinity_polling()


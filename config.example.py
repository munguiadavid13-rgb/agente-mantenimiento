# =====================================================================
# config.example.py  -  PLANTILLA de configuracion (sin credenciales reales)
#
# COMO USAR:
#   1. Copia este archivo y renombralo a  config.py
#   2. Rellena cada valor con tus datos reales
#   config.py esta protegido por .gitignore y NO se sube a internet.
# =====================================================================

# ---------- CORREO (Gmail) ----------
# Necesitas una "Contrasena de aplicacion" de Gmail (16 letras), NO tu clave normal.
CORREO_REMITENTE    = "tucorreo@gmail.com"        # desde que cuenta se envia
CORREO_APP_PASSWORD = "xxxx xxxx xxxx xxxx"        # la contrasena de aplicacion de Gmail
# A quien(es) le llega la alerta: una o varias direcciones en la lista.
CORREO_DESTINATARIOS = [
    "tucorreo@gmail.com",
    # "otra.persona@ejemplo.com",
]

# ---------- TELEGRAM ----------
# El TOKEN te lo da @BotFather al crear el bot.
# Cada CHAT_ID se obtiene en https://api.telegram.org/bot<TOKEN>/getUpdates
# (cada persona debe escribirle al bot primero para que aparezca su chat_id).
TELEGRAM_TOKEN    = "123456789:AAAA-tu-token-de-BotFather"
TELEGRAM_CHAT_IDS = [
    "000000000",
    # "111111111",
]

# ---------- DASHBOARD WEB ----------
# Si publicas el dashboard en internet (ej. GitHub Pages), pon aqui su URL y
# aparecera como enlace en el correo y en Telegram. Dejalo "" si no lo publicas.
DASHBOARD_URL = ""

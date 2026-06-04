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
CORREO_DESTINATARIO = "tucorreo@gmail.com"        # a quien le llega (puede ser el mismo)

# ---------- TELEGRAM ----------
# El TOKEN te lo da @BotFather al crear el bot.
# El CHAT_ID se obtiene en https://api.telegram.org/bot<TOKEN>/getUpdates
TELEGRAM_TOKEN   = "123456789:AAAA-tu-token-de-BotFather"
TELEGRAM_CHAT_ID = "000000000"

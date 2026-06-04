"""
notificaciones.py
-----------------
Envia el reporte del agente por CORREO (Gmail) y por TELEGRAM.
Las credenciales NO van aqui: se leen del archivo config.py (privado).
Solo usa librerias que Python ya trae (smtplib, urllib) -> no instala nada.
"""

import smtplib
import ssl
import urllib.parse
import urllib.request
from email.message import EmailMessage

import config   # tu archivo con las credenciales (no se comparte)


def _contexto_ssl():
    """
    La red de la empresa inspecciona el trafico HTTPS con su propio certificado,
    que Python no reconoce. Desactivamos la verificacion para que el envio no falle.
    (Suficiente para este proyecto; lo "correcto" seria instalar el certificado raiz
    de la empresa, pero eso requiere permisos de TI.)
    """
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def enviar_correo(asunto, cuerpo, cuerpo_html=None, imagenes=None):
    """
    Envia un correo usando la cuenta de Gmail definida en config.py.
    cuerpo      -> version en texto plano (respaldo si el cliente no muestra HTML).
    cuerpo_html -> version con tablas/colores; si se da, el cliente la prefiere.
    imagenes    -> lista de (cid, bytes_png) que el HTML referencia como
                   <img src="cid:..."> para mostrar los graficos incrustados.
    """
    msg = EmailMessage()
    msg["From"] = config.CORREO_REMITENTE
    msg["To"] = config.CORREO_DESTINATARIO
    msg["Subject"] = asunto
    msg.set_content(cuerpo)                      # texto plano (siempre va)
    if cuerpo_html:
        msg.add_alternative(cuerpo_html, subtype="html")   # version "bonita"
        if imagenes:
            # La parte HTML es la ultima alternativa; le "adjuntamos" las
            # imagenes como contenido relacionado (cid) para que se vean
            # dentro del cuerpo, no como archivos sueltos.
            parte_html = msg.get_payload()[-1]
            for cid, datos_png in imagenes:
                parte_html.add_related(datos_png, "image", "png", cid=f"<{cid}>")

    with smtplib.SMTP("smtp.gmail.com", 587) as servidor:
        servidor.starttls(context=_contexto_ssl())                # cifra la conexion
        servidor.login(config.CORREO_REMITENTE, config.CORREO_APP_PASSWORD)
        servidor.send_message(msg)
    print("  -> Correo enviado a", config.CORREO_DESTINATARIO)


def enviar_telegram(texto, parse_mode=None):
    """
    Envia un mensaje por Telegram usando el bot definido en config.py.
    parse_mode: None (texto plano) o "HTML" para usar negritas/cursivas/etc.
    Telegram rechaza mensajes de mas de 4096 caracteres, asi que recortamos
    en un salto de linea para no cortar una etiqueta HTML a la mitad.
    """
    LIMITE = 4096
    if len(texto) > LIMITE:
        corte = texto.rfind("\n", 0, LIMITE - 40)
        if corte == -1:
            corte = LIMITE - 40
        texto = texto[:corte] + "\n\n... (mensaje recortado, ver correo)"

    url = f"https://api.telegram.org/bot{config.TELEGRAM_TOKEN}/sendMessage"
    campos = {
        "chat_id": config.TELEGRAM_CHAT_ID,
        "text": texto,
    }
    if parse_mode:
        campos["parse_mode"] = parse_mode
    datos = urllib.parse.urlencode(campos).encode()
    with urllib.request.urlopen(url, data=datos, context=_contexto_ssl()) as respuesta:
        respuesta.read()
    print("  -> Mensaje de Telegram enviado")


def notificar(asunto, cuerpo_correo, texto_telegram=None, parse_mode_telegram=None,
              cuerpo_html_correo=None, imagenes_correo=None):
    """
    Intenta enviar por los dos canales; si uno falla, avisa pero no se cae.
    Si no se da texto_telegram, usa el mismo cuerpo del correo.
    parse_mode_telegram: "HTML" para enviar el mensaje de Telegram con formato.
    cuerpo_html_correo : version HTML del correo (tablas/colores); opcional.
    imagenes_correo    : lista de (cid, bytes_png) con los graficos a incrustar.
    """
    if texto_telegram is None:
        texto_telegram = cuerpo_correo
    try:
        enviar_correo(asunto, cuerpo_correo, cuerpo_html=cuerpo_html_correo,
                      imagenes=imagenes_correo)
    except Exception as e:
        print("  -> Error enviando correo:", e)
    try:
        enviar_telegram(texto_telegram, parse_mode=parse_mode_telegram)
    except Exception as e:
        print("  -> Error enviando Telegram:", e)

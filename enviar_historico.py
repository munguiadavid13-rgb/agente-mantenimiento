"""
enviar_historico.py
-------------------
Envia a una persona (o a varias) el HISTORICO de las ultimas N visitas de un
equipo. Pensado para cuando alguien pide, por ejemplo, "mandame las ultimas 3
visitas del ISL-32L44".

USO (desde la carpeta del proyecto):

  # Ultimas 3 visitas del equipo, a un correo puntual:
  python enviar_historico.py ISL-32L44 --n 3 --correo persona@ejemplo.com

  # A un chat de Telegram puntual:
  python enviar_historico.py CTE-32L11 --telegram 584397421

  # A varios destinatarios a la vez:
  python enviar_historico.py SIS-32L27 --correo a@x.com b@y.com --telegram 584397421

  # Sin --correo ni --telegram  -> usa los destinatarios de config.py
  python enviar_historico.py TEL-32L9

  # Solo VER el mensaje en pantalla, sin enviar nada (para probar):
  python enviar_historico.py ISL-32L44 --demo

Notas:
  - Si das --correo, solo manda correo; si das --telegram, solo Telegram.
  - Si no das ninguno, manda a los destinatarios por defecto de config.py.
"""

import re
import sys
import argparse

# La consola de Windows (cp1252) no sabe imprimir emojis; forzamos UTF-8 para
# que el modo --demo no se caiga al mostrar el mensaje. No afecta el envio real.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from openpyxl import load_workbook

import agente   # reutilizamos construir_modelo, COLUMNAS, _badge, _esc, _COLORES


def _clave_codigo(codigo):
    """Normaliza un codigo para compararlo: mayusculas, sin guiones/espacios y
    sin ceros a la izquierda en los numeros (asi CTE-32L08 == CTE-32L8)."""
    base = str(codigo).upper().replace("-", "").replace(" ", "")
    return re.sub(r"0+(\d)", r"\1", base)


def buscar_equipo(subestaciones, codigo):
    """Encuentra el Equipo cuyo codigo coincide (tolerante a ceros/guiones)."""
    objetivo = _clave_codigo(codigo)
    for e in agente.todos_los_equipos(subestaciones):
        if _clave_codigo(e.codigo) == objetivo:
            return e
    return None


# ------------------------------------------------------------------
# Construccion de los mensajes (texto plano, Telegram HTML y correo HTML)
# ------------------------------------------------------------------
def _texto_plano(equipo, visitas):
    L = [f"Agente Buchholz - Historico de {equipo.codigo}",
         f"{equipo.tipo} - {equipo.subestacion}",
         f"Ultimas {len(visitas)} visita(s):", ""]
    for v in visitas:
        L.append(f"== {str(v.fecha)[:10]}  (Cuadrilla {v.cuadrilla}) ==")
        for columna, parametro, unidad in agente.COLUMNAS:
            m = v.medicion(parametro)
            if m is None:
                continue
            valor = "(sin dato)" if m.valor is None else f"{m.valor} {m.unidad}"
            estado = m.clasificar()
            marca = "  <-- REVISAR" if estado == "REVISAR" else ""
            L.append(f"   {parametro}: {valor}{marca}")
        if v.observacion:
            L.append(f"   Obs: {v.observacion}")
        L.append("")
    return "\n".join(L)


def _texto_telegram(equipo, visitas):
    e = agente._esc
    L = [f"\U0001F4D2 <b>Historico {e(equipo.codigo)}</b> <i>({e(equipo.tipo)})</i>",
         f"\U0001F4CD {e(equipo.subestacion)} · ultimas {len(visitas)} visita(s)"]
    for v in visitas:
        L.append("")
        L.append(f"\U0001F4C5 <b>{str(v.fecha)[:10]}</b> · Cuadrilla {e(v.cuadrilla)}")
        for columna, parametro, unidad in agente.COLUMNAS:
            m = v.medicion(parametro)
            if m is None or m.valor is None:
                continue
            punto = "\U0001F534" if m.clasificar() == "REVISAR" else "•"
            L.append(f"  {punto} {e(parametro)}: <b>{m.valor} {e(m.unidad)}</b>")
        if v.observacion:
            L.append(f"  \U0001F4DD <i>{e(v.observacion)}</i>")
    enlace = agente._enlace_telegram_dashboard()
    if enlace:
        L.append("")
        L.append(enlace)
    return "\n".join(L)


def _html(equipo, visitas):
    e = agente._esc
    H = ['<div style="font-family:Arial,Helvetica,sans-serif;color:#212529;'
         'max-width:760px;margin:auto;">']
    H.append('<div style="background:#0d6efd;color:#fff;padding:16px 20px;'
             'border-radius:8px 8px 0 0;">'
             f'<h2 style="margin:0;">&#128210; Agente Buchholz &mdash; Historico {e(equipo.codigo)}</h2>'
             f'<p style="margin:4px 0 0;opacity:.9;">({e(equipo.tipo)}) &mdash; '
             f'{e(equipo.subestacion)} &middot; ultimas {len(visitas)} visita(s)</p></div>')
    H.append('<div style="border:1px solid #dee2e6;border-top:none;padding:20px;'
             'border-radius:0 0 8px 8px;">')
    for v in visitas:
        H.append('<h3 style="border-bottom:2px solid #0d6efd;padding-bottom:4px;">'
                 f'&#128197; {str(v.fecha)[:10]} '
                 '<span style="font-weight:normal;color:#6c757d;font-size:14px;">'
                 f'&middot; Cuadrilla {e(v.cuadrilla)}</span></h3>')
        H.append('<table style="border-collapse:collapse;width:100%;font-size:14px;'
                 'margin-bottom:14px;">')
        for columna, parametro, unidad in agente.COLUMNAS:
            m = v.medicion(parametro)
            if m is None:
                continue
            valor = "(sin dato)" if m.valor is None else f"{m.valor} {m.unidad}"
            H.append('<tr>'
                     + agente._celda(e(parametro), "width:45%;")
                     + agente._celda(valor)
                     + agente._celda(agente._badge(m.clasificar()),
                                     "text-align:center;width:120px;")
                     + '</tr>')
        H.append('</table>')
        if v.observacion:
            H.append('<p style="font-size:12px;color:#6c757d;margin:0 0 12px;">'
                     f'&#128221; {e(v.observacion)}</p>')
    H.append(agente._boton_html_dashboard())
    H.append('<p style="font-size:12px;color:#adb5bd;margin-top:8px;">'
             'Enviado por el Agente Buchholz &middot; Jose David Perez Munguia '
             '&mdash; UTH 2026.4</p>')
    H.append('</div></div>')
    return "\n".join(H)


def main():
    p = argparse.ArgumentParser(
        description="Envia el historico de las ultimas N visitas de un equipo.")
    p.add_argument("codigo", help="Codigo del equipo, ej. ISL-32L44")
    p.add_argument("--n", type=int, default=3,
                   help="Cuantas visitas (las mas recientes). Por defecto 3.")
    p.add_argument("--correo", nargs="+", metavar="EMAIL",
                   help="Correo(s) destino solo para este envio.")
    p.add_argument("--telegram", nargs="+", metavar="CHAT_ID",
                   help="Chat_id(s) de Telegram destino solo para este envio.")
    p.add_argument("--demo", action="store_true",
                   help="Solo muestra el mensaje en pantalla, no envia nada.")
    args = p.parse_args()

    wb = load_workbook(agente.ARCHIVO, data_only=True)
    subs = agente.construir_modelo(wb)

    equipo = buscar_equipo(subs, args.codigo)
    if equipo is None:
        print(f"No encontre el equipo '{args.codigo}'. Revisa el codigo.")
        return
    if not equipo.visitas:
        print(f"El equipo {equipo.codigo} no tiene visitas registradas.")
        return

    visitas = equipo.visitas_ordenadas()[-args.n:]
    asunto = f"Agente Buchholz - Historico {equipo.codigo} (ultimas {len(visitas)} visitas)"
    texto = _texto_plano(equipo, visitas)
    texto_tg = _texto_telegram(equipo, visitas)
    cuerpo_html = _html(equipo, visitas)

    if args.demo:
        print("===== ASUNTO =====")
        print(asunto)
        print("\n===== TEXTO PLANO / CORREO (respaldo) =====")
        print(texto)
        print("\n===== TELEGRAM (HTML) =====")
        print(texto_tg)
        print("\n(modo --demo: no se envio nada)")
        return

    # A quien enviar: si das --correo/--telegram, solo esos; si no, los de config.
    sin_destino = not args.correo and not args.telegram
    import notificaciones

    if args.correo or sin_destino:
        try:
            notificaciones.enviar_correo(asunto, texto, cuerpo_html=cuerpo_html,
                                         destinatarios=args.correo)
        except Exception as ex:
            print("  -> Error enviando correo:", ex)

    if args.telegram or sin_destino:
        try:
            notificaciones.enviar_telegram(texto_tg, parse_mode="HTML",
                                           chat_ids=args.telegram)
        except Exception as ex:
            print("  -> Error enviando Telegram:", ex)

    print(f"Listo: historico de {equipo.codigo} ({len(visitas)} visitas) enviado.")


if __name__ == "__main__":
    main()

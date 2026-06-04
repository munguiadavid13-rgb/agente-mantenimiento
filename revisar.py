"""
revisar.py
----------
Revision automatica del agente (pensada para correr cada cierto tiempo en
segundo plano, ej. cada 50 minutos):

  1. Relee el Excel y arma el modelo.
  2. Regenera el dashboard (dashboard.html) con los datos frescos.
  3. Escribe una linea de estado en 'revisiones_log.txt'.

NO envia correo ni Telegram y NO abre el navegador (para no molestar ni
saturar de notificaciones). Las notificaciones se envian con 'python agente.py'.

Autor: Jose David Perez Munguia - UTH 2026.4
"""

import os
import datetime

from openpyxl import load_workbook

import agente
import dashboard

LOG = os.path.join(agente.CARPETA, "revisiones_log.txt")


def revisar():
    momento = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    try:
        wb = load_workbook(agente.ARCHIVO, data_only=True)
        subs = agente.construir_modelo(wb)
        equipos = agente.todos_los_equipos(subs)
        hallazgos = agente.todos_los_hallazgos(subs)

        # regenerar el dashboard sin abrir el navegador
        html = dashboard.construir_html()
        with open(dashboard.SALIDA, "w", encoding="utf-8") as f:
            f.write(html)

        codigos = sorted({e.codigo for e, v, m in hallazgos})
        detalle = (" Hallazgos en: " + ", ".join(codigos)) if codigos else ""
        linea = (f"[{momento}] OK - {len(equipos)} equipos, "
                 f"{len(hallazgos)} hallazgo(s). Dashboard actualizado.{detalle}")
    except Exception as e:
        linea = f"[{momento}] ERROR: {e}"

    with open(LOG, "a", encoding="utf-8") as f:
        f.write(linea + "\n")
    print(linea)


if __name__ == "__main__":
    revisar()

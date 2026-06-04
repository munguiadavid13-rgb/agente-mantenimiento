"""
agente.py
---------
Agente de Mantenimiento Predictivo de subestaciones electricas.

Flujo:
  1. Lee el Excel 'Mantenimiento_Subestaciones.xlsx' (hojas Mediciones y Limites).
  2. Construye el arbol de objetos: Subestacion -> Equipo -> Visita -> Medicion.
  3. Genera un reporte legible (resumen, hallazgos, estado por equipo, historico).
  4. Si hay hallazgos, envia alerta por Telegram y reporte completo por correo.

Autor: Jose David Perez Munguia - UTH 2026.4
"""

import os
import html
import datetime
from openpyxl import load_workbook

from modelos import Medicion, Visita, Subestacion

# Ruta del Excel: misma carpeta que este archivo
CARPETA = os.path.dirname(os.path.abspath(__file__))
ARCHIVO = os.path.join(CARPETA, "Mantenimiento_Subestaciones.xlsx")

# Todos los parametros que se miden: (columna en Excel, nombre, unidad)
COLUMNAS = [
    ("Voltaje VDC (V)",            "Voltaje VDC",        "V"),
    ("Voltaje AC (V)",             "Voltaje AC",         "V"),
    ("Resistencia tierra (Ohm)",   "Resistencia tierra", "Ohm"),
    ("Corriente Fase A (A)",       "Corriente Fase A",   "A"),
    ("Corriente Fase B (A)",       "Corriente Fase B",   "A"),
    ("Corriente Fase C (A)",       "Corriente Fase C",   "A"),
    ("Potencia Activa (MW)",       "Potencia Activa",    "MW"),
    ("Potencia Reactiva (MVAR)",   "Potencia Reactiva",  "MVAR"),
    ("Num Operaciones",            "Num Operaciones",    "oper"),
]


# ============================================================
# Lectura del Excel -> arbol de objetos
# ============================================================
def cargar_limites(wb):
    """Devuelve {parametro: {criterio, limite}} desde la hoja Limites."""
    ws = wb["Limites"]
    limites = {}
    for fila in ws.iter_rows(min_row=2, values_only=True):
        parametro, criterio, limite = (list(fila) + [None]*3)[:3]
        if parametro:
            limites[parametro] = {"criterio": criterio, "limite": limite}
    return limites


def _agregar_medicion(visita, fila, limites):
    """Llena una visita con las mediciones de una fila (dict columna->valor)."""
    for columna, parametro, unidad in COLUMNAS:
        lim = limites.get(parametro)
        if lim:
            visita.agregar(Medicion(parametro, fila.get(columna), unidad,
                                    lim["criterio"], lim["limite"]))
        else:
            visita.agregar(Medicion(parametro, fila.get(columna), unidad))


def construir_modelo(wb):
    """
    Arma el arbol Subestacion -> Equipo -> Visita -> Medicion leyendo:
      - la hoja indice 'Equipos' (Subestacion, Codigo, Tipo), y
      - una hoja por equipo (nombrada con su codigo) con sus visitas.
    Asi cada equipo tiene su propia tabla y los datos no se mezclan.
    """
    limites = cargar_limites(wb)
    subestaciones = {}                      # nombre -> Subestacion

    indice = wb["Equipos"]
    for fila in indice.iter_rows(min_row=2, values_only=True):
        nombre_sub, codigo, tipo = (list(fila) + [None] * 3)[:3]
        if not codigo:                      # fila vacia en el indice
            continue
        if codigo not in wb.sheetnames:     # equipo listado pero sin hoja propia
            print(f"  (aviso) El equipo {codigo} no tiene hoja de datos; se omite.")
            continue

        sub = subestaciones.setdefault(nombre_sub, Subestacion(nombre_sub))
        equipo = sub.obtener_equipo(codigo, tipo)

        # leer la hoja propia del equipo
        hoja = wb[codigo]
        titulos = [c.value for c in hoja[1]]
        for fila_v in hoja.iter_rows(min_row=2, values_only=True):
            datos = dict(zip(titulos, fila_v))
            if datos.get("Fecha") is None:  # fila vacia
                continue
            visita = Visita(str(datos["Fecha"]), datos.get("Cuadrilla"),
                            datos.get("Observacion") or "")
            _agregar_medicion(visita, datos, limites)
            equipo.agregar_visita(visita)

    return subestaciones


# ============================================================
# Utilidades para recorrer el modelo
# ============================================================
def todos_los_equipos(subestaciones):
    return [e for s in subestaciones.values() for e in s.equipos.values()]


def todos_los_hallazgos(subestaciones):
    """Lista de (equipo, visita, medicion) que estan en REVISAR."""
    salida = []
    for e in todos_los_equipos(subestaciones):
        for v in e.visitas:
            for m in v.hallazgos():
                salida.append((e, v, m))
    return salida


# ============================================================
# Reporte COMPLETO (para correo y consola)
# ============================================================
SEP = "=" * 60
SUB = "-" * 60


def _linea_param(visita, parametro, etiqueta):
    """Formatea una linea legible de una medicion: etiqueta ... valor [estado]."""
    m = visita.medicion(parametro)
    if m is None or m.valor is None:
        return f"        {etiqueta:.<24} (sin dato)"
    valor = f"{m.valor} {m.unidad}"
    if m.tiene_limite():
        return f"        {etiqueta:.<24} {valor:<11} [{m.clasificar()}]"
    return f"        {etiqueta:.<24} {valor}"


def reporte_completo(subestaciones, hoy):
    eq = todos_los_equipos(subestaciones)
    visitas = [v for e in eq for v in e.visitas]
    hallazgos = todos_los_hallazgos(subestaciones)
    L = []

    # ---- encabezado ----
    L.append(SEP)
    L.append("        AGENTE DE MANTENIMIENTO PREDICTIVO")
    L.append(f"        Reporte generado: {hoy}")
    L.append(SEP)
    L.append("")
    L.append(" RESUMEN")
    L.append(f"   Subestaciones : {len(subestaciones)}")
    L.append(f"   Equipos       : {len(eq)}")
    L.append(f"   Visitas       : {len(visitas)}")
    L.append(f"   Hallazgos     : {len(hallazgos)}  (requieren accion)")

    # ---- 1) hallazgos ----
    L.append("")
    L.append(SEP)
    L.append(" 1) HALLAZGOS QUE REQUIEREN ACCION")
    L.append(SEP)
    if not hallazgos:
        L.append("   Sin hallazgos: todos los parametros con limite estan NORMAL.")
    for e, v, m in hallazgos:
        L.append("")
        L.append(f"   [!] {e.codigo}  ({e.tipo})")
        L.append(f"       Subestacion : {e.subestacion}")
        L.append(f"       Fecha       : {v.fecha}        Cuadrilla: {v.cuadrilla}")
        L.append(f"       Parametro   : {m.parametro}")
        L.append(f"       Medido      : {m.valor} {m.unidad}   (limite {m.criterio}: {m.limite})")
        if v.observacion:
            L.append(f"       Observacion : {v.observacion}")

    # ---- 2) estado actual por equipo ----
    L.append("")
    L.append(SEP)
    L.append(" 2) ESTADO ACTUAL POR EQUIPO (ultima visita)")
    L.append(SEP)
    for e in eq:
        v = e.ultima_visita()
        if v is None:
            continue
        L.append("")
        L.append(f"   {e.codigo}  ({e.tipo})  -  {e.subestacion}")
        L.append(f"   Ultima visita: {v.fecha}   Cuadrilla: {v.cuadrilla}")
        L.append("     Voltajes:")
        L.append(_linea_param(v, "Voltaje VDC", "Voltaje VDC"))
        L.append(_linea_param(v, "Voltaje AC", "Voltaje AC"))
        L.append("     Puesta a tierra:")
        L.append(_linea_param(v, "Resistencia tierra", "Resistencia tierra"))
        L.append("     Corrientes (monitoreo):")
        L.append(_linea_param(v, "Corriente Fase A", "Fase A"))
        L.append(_linea_param(v, "Corriente Fase B", "Fase B"))
        L.append(_linea_param(v, "Corriente Fase C", "Fase C"))
        L.append("     Potencias (monitoreo):")
        L.append(_linea_param(v, "Potencia Activa", "Activa"))
        L.append(_linea_param(v, "Potencia Reactiva", "Reactiva"))
        L.append("     Otros:")
        L.append(_linea_param(v, "Num Operaciones", "Num operaciones"))
        if v.observacion:
            L.append(f"     Observacion: {v.observacion}")
        L.append("   " + "." * 50)

    # ---- 3) historico de tierra ----
    L.append("")
    L.append(SEP)
    L.append(" 3) HISTORICO Y TENDENCIA - Resistencia de tierra (Ohm)")
    L.append(SEP)
    for e in eq:
        datos = e.historico("Resistencia tierra")
        if not datos:
            continue
        L.append("")
        L.append(f"   {e.codigo} ({e.tipo}) - {e.subestacion}")
        for fecha, valor in datos:
            estado = Medicion("Resistencia tierra", valor, "Ohm", "maximo", 5).clasificar()
            marca = "  <-- REVISAR" if estado == "REVISAR" else ""
            L.append(f"        {fecha}   {valor:>5} Ohm   {estado}{marca}")
        L.append(f"        Tendencia: {e.tendencia('Resistencia tierra')}")

    L.append("")
    L.append(SEP)
    return "\n".join(L)


# ============================================================
# Reporte CORTO (alerta para Telegram, con formato HTML)
# ============================================================
def _emoji_tendencia(texto):
    """Convierte el texto de tendencia ('subiendo (...)') en una flecha."""
    if texto.startswith("subiendo"):
        return "\U0001F4C8"          # grafico subiendo
    if texto.startswith("bajando"):
        return "\U0001F4C9"          # grafico bajando
    if texto.startswith("estable"):
        return "➡️"        # flecha a la derecha
    return ""


def _esc(texto):
    """Escapa <, > y & para que no rompan el formato HTML de Telegram."""
    return html.escape(str(texto))


def reporte_corto(subestaciones, hoy):
    """
    Alerta breve para Telegram con formato HTML (emojis + negritas).
    Telegram HTML solo admite <b>, <i>, <u>, <code>, etc.; los saltos de
    linea son '\\n' normales y el texto dinamico se escapa con _esc().
    """
    hallazgos = todos_los_hallazgos(subestaciones)
    L = []
    L.append("\U0001F527 <b>AGENTE DE MANTENIMIENTO - ALERTA</b>")
    L.append(f"\U0001F4C5 {hoy}")
    if hallazgos:
        L.append(f"⚠️ <b>{len(hallazgos)} hallazgo(s)</b> requieren accion")
    else:
        L.append("✅ <b>Sin hallazgos</b>: todos los parametros con limite estan NORMAL")

    for e, v, m in hallazgos:
        criterio = "max" if m.criterio == "maximo" else "min"
        flecha = _emoji_tendencia(e.tendencia(m.parametro))
        fecha_corta = str(v.fecha)[:10]
        L.append("")
        L.append(f"\U0001F534 <b>{_esc(e.codigo)}</b> <i>({_esc(e.tipo)})</i> - {_esc(e.subestacion)}")
        L.append(f"   • {_esc(m.parametro)}: <b>{m.valor} {_esc(m.unidad)}</b> "
                 f"(limite {criterio} {m.limite}) {flecha}".rstrip())
        L.append(f"   • \U0001F4CD Cuadrilla {_esc(v.cuadrilla)} · {fecha_corta}")
        if v.observacion:
            L.append(f"   • \U0001F4DD <i>{_esc(v.observacion)}</i>")

    if hallazgos:
        L.append("")
        L.append("\U0001F4E7 <i>Detalle completo enviado al correo.</i>")
    return "\n".join(L)


# ============================================================
# Reporte HTML SEGMENTADO (para correo): solo el equipo en alerta
# ============================================================
# Cada estado tiene su par de colores (fondo, texto):
_COLORES = {
    "REVISAR":   ("#f8d7da", "#842029"),   # rojo
    "NORMAL":    ("#d1e7dd", "#0f5132"),   # verde
    "MONITOREO": ("#e2e3e5", "#41464b"),   # gris
    "SIN DATO":  ("#fff3cd", "#664d03"),   # amarillo
}


def _badge(estado):
    """Devuelve una 'etiqueta' HTML coloreada segun el estado."""
    fondo, texto = _COLORES.get(estado, ("#e2e3e5", "#41464b"))
    return (f'<span style="background:{fondo};color:{texto};padding:2px 8px;'
            f'border-radius:10px;font-size:12px;font-weight:bold;">{estado}</span>')


def _celda(contenido, extra=""):
    """Una celda <td> con el estilo base de las tablas."""
    return f'<td style="padding:6px 10px;border:1px solid #dee2e6;{extra}">{contenido}</td>'


def _cid(equipo, nombre):
    """Content-ID seguro (solo alfanumerico) para una imagen del correo."""
    base = f"{nombre}_{equipo.codigo}"
    return "".join(c if c.isalnum() else "_" for c in base)


def reporte_html(subestaciones, hoy):
    """
    Genera el correo HTML ENFOCADO en los equipos que tienen alerta.
    No incluye los equipos que estan bien: solo el/los equipo(s) en alerta
    con su motivo, su ultima visita completa y graficos de su historico.

    Devuelve (html, imagenes) donde imagenes es una lista de (cid, bytes_png)
    que el HTML referencia como <img src="cid:..."> para incrustar los graficos.
    """
    import graficos   # se importa aqui para no exigir matplotlib si no se usa

    hallazgos = todos_los_hallazgos(subestaciones)

    # equipos unicos con alerta, conservando el orden de aparicion
    equipos_alerta = []
    for e, v, m in hallazgos:
        if e not in equipos_alerta:
            equipos_alerta.append(e)

    imagenes = []
    H = []
    H.append('<div style="font-family:Arial,Helvetica,sans-serif;color:#212529;'
             'max-width:760px;margin:auto;">')

    # ---- encabezado (rojo: es una alerta) ----
    H.append('<div style="background:#dc3545;color:#fff;padding:16px 20px;'
             'border-radius:8px 8px 0 0;">')
    H.append('<h2 style="margin:0;">&#9888;&#65039; Alerta de Mantenimiento Predictivo</h2>')
    H.append(f'<p style="margin:4px 0 0;opacity:.9;">Reporte generado: {hoy} '
             f'&middot; {len(equipos_alerta)} equipo(s) en alerta</p>')
    H.append('</div>')
    H.append('<div style="border:1px solid #dee2e6;border-top:none;padding:20px;'
             'border-radius:0 0 8px 8px;">')

    # ---- una seccion por equipo en alerta ----
    for e in equipos_alerta:
        suyos = [(vv, mm) for (ee, vv, mm) in hallazgos if ee is e]

        H.append('<h3 style="border-bottom:2px solid #dc3545;padding-bottom:4px;">'
                 f'&#128295; {_esc(e.codigo)} '
                 '<span style="font-weight:normal;color:#6c757d;font-size:15px;">'
                 f'({_esc(e.tipo)}) &mdash; {_esc(e.subestacion)}</span></h3>')

        # motivo de la alerta
        H.append('<div style="background:#f8d7da;color:#842029;padding:10px 14px;'
                 'border-radius:6px;margin-bottom:14px;">')
        H.append('<b>Motivo de la alerta:</b>'
                 '<ul style="margin:6px 0 0;padding-left:20px;">')
        for vv, mm in suyos:
            criterio = "max" if mm.criterio == "maximo" else "min"
            H.append(f'<li>{_esc(mm.parametro)}: <b>{mm.valor} {_esc(mm.unidad)}</b> '
                     f'(limite {criterio} {mm.limite}) &middot; '
                     f'visita {str(vv.fecha)[:10]}</li>')
        H.append('</ul></div>')

        # ultima visita completa
        v = e.ultima_visita()
        if v is not None:
            H.append(f'<p style="margin:0 0 4px;"><b>Ultima visita:</b> '
                     f'{str(v.fecha)[:10]} &middot; Cuadrilla {_esc(v.cuadrilla)}</p>')
            H.append('<table style="border-collapse:collapse;width:100%;'
                     'font-size:14px;margin-bottom:16px;">')
            for columna, parametro, unidad in COLUMNAS:
                m = v.medicion(parametro)
                if m is None:
                    continue
                valor = "(sin dato)" if m.valor is None else f"{m.valor} {m.unidad}"
                H.append('<tr>')
                H.append(_celda(_esc(parametro), "width:45%;"))
                H.append(_celda(valor))
                H.append(_celda(_badge(m.clasificar()), "text-align:center;width:120px;"))
                H.append('</tr>')
            H.append('</table>')
            if v.observacion:
                H.append('<p style="font-size:12px;color:#6c757d;margin:0 0 12px;">'
                         f'&#128221; {_esc(v.observacion)}</p>')

        # graficos del historico (PNG incrustados via cid)
        pngs = graficos.graficos_equipo(e)
        for nombre, etiqueta in [("tierra", "Resistencia de tierra (con limite)"),
                                 ("corrientes", "Corrientes por fase"),
                                 ("potencias", "Potencias: activa / reactiva / aparente")]:
            cid = _cid(e, nombre)
            imagenes.append((cid, pngs[nombre]))
            H.append(f'<p style="margin:14px 0 4px;font-weight:bold;color:#495057;">'
                     f'{etiqueta}</p>')
            H.append(f'<img src="cid:{cid}" alt="{etiqueta}" '
                     'style="width:100%;max-width:600px;border:1px solid #dee2e6;'
                     'border-radius:6px;">')

        H.append('<hr style="border:none;border-top:1px solid #dee2e6;margin:24px 0;">')

    H.append('<p style="font-size:12px;color:#adb5bd;margin-top:8px;">'
             'Generado automaticamente por el Agente de Mantenimiento Predictivo '
             '&middot; Jose David Perez Munguia &mdash; UTH 2026.4</p>')
    H.append('</div></div>')
    return "\n".join(H), imagenes


# ============================================================
# Programa principal
# ============================================================
def main():
    if not os.path.exists(ARCHIVO):
        print("No encuentro el Excel en:", ARCHIVO)
        return

    hoy = datetime.date.today().isoformat()
    wb = load_workbook(ARCHIVO, data_only=True)
    subestaciones = construir_modelo(wb)

    completo = reporte_completo(subestaciones, hoy)
    print(completo)

    # Notificar solo si hay algo que revisar
    if todos_los_hallazgos(subestaciones):
        try:
            import notificaciones
            print("\nEnviando notificaciones...")
            html_correo, imagenes = reporte_html(subestaciones, hoy)   # HTML + graficos
            notificaciones.notificar(
                "Agente de Mantenimiento - Hallazgos",
                completo,                                  # correo: texto plano (respaldo)
                reporte_corto(subestaciones, hoy),         # telegram: alerta corta
                parse_mode_telegram="HTML",                # telegram con formato
                cuerpo_html_correo=html_correo,            # correo: HTML segmentado
                imagenes_correo=imagenes,                  # correo: graficos incrustados
            )
        except ImportError:
            print("\n(Notificaciones aun no configuradas: falta config.py)")
        except Exception as e:
            print("\n(No se pudo notificar:", e, ")")


if __name__ == "__main__":
    main()

"""
dashboard.py
------------
Genera un DASHBOARD WEB ESTATICO (un archivo dashboard.html) del agente de
mantenimiento predictivo. No levanta ningun servidor: solo crea el archivo y
lo abre en el navegador.

Muestra UN equipo a la vez, con dos menus desplegables para navegar:
  - Subestacion  (Ceiba Termica, San Isidro, Isletas, ...)
  - Equipo       (los interruptores/restauradores de esa subestacion)

Para cada equipo: su estado de la ultima visita (con colores) y sus 3 graficos
de historico (resistencia de tierra, corrientes y potencias).

Uso:
    python dashboard.py
    -> crea 'dashboard.html' y lo abre en el navegador.
    Para ver datos actualizados, vuelve a ejecutarlo.

Autor: Jose David Perez Munguia - UTH 2026.4
"""

import os
import base64
import datetime
import html
import webbrowser

from openpyxl import load_workbook

import agente
import graficos

SALIDA = os.path.join(agente.CARPETA, "dashboard.html")


# ------------------------------------------------------------------
# Utilidades de presentacion
# ------------------------------------------------------------------
def _img(png_bytes):
    """Devuelve una etiqueta <img> con el PNG embebido (base64), sin archivos."""
    b64 = base64.b64encode(png_bytes).decode()
    return (f'<img src="data:image/png;base64,{b64}" '
            'style="width:100%;max-width:560px;border:1px solid #dee2e6;'
            'border-radius:6px;margin:6px 0;">')


def _badge_html(estado):
    """Etiqueta coloreada reutilizando los colores del agente."""
    fondo, texto = agente._COLORES.get(estado, ("#e2e3e5", "#41464b"))
    return f'<span class="badge" style="background:{fondo};color:{texto};">{estado}</span>'


def _ficha_html(equipo):
    """Muestra la ficha tecnica (datos de placa) como chips; solo campos con valor."""
    items = []
    for clave, etiqueta in agente.FICHA_CAMPOS:
        val = equipo.ficha.get(clave)
        if val is not None and str(val).strip() != "":
            items.append(f'<span class="it"><b>{etiqueta}:</b> {html.escape(str(val))}</span>')
    if not items:
        return '<div class="sub" style="color:#adb5bd;">Sin ficha tecnica registrada.</div>'
    return '<div class="ficha">' + "".join(items) + '</div>'


ESTILOS = """
* { box-sizing: border-box; }
body { font-family: Arial, Helvetica, sans-serif; color:#212529;
       background:#f1f3f5; margin:0; padding:0 0 40px; }
.barra { background:#0d6efd; color:#fff; padding:14px 24px; display:flex;
         justify-content:space-between; align-items:baseline; flex-wrap:wrap; gap:6px; }
.barra h1 { margin:0; font-size:19px; }
.barra span { font-size:12px; opacity:.9; }
.contenido { max-width:960px; margin:18px auto; padding:0 16px; }
.resumen { display:flex; gap:10px; flex-wrap:wrap; margin-bottom:14px; }
.stat { flex:1; min-width:110px; background:#fff; border:1px solid #dee2e6;
        border-radius:8px; padding:10px 14px; }
.stat .n { font-size:22px; font-weight:bold; line-height:1.1; }
.stat .e { font-size:12px; color:#6c757d; }
.stat.rojo .n { color:#dc3545; }
.stat.verde .n { color:#198754; }
.ok { background:#d1e7dd; color:#0f5132; border:1px solid #badbcc;
      padding:10px 14px; border-radius:8px; margin-bottom:14px; }
.panel { background:#fff; border:1px solid #dee2e6; border-radius:8px;
         padding:14px 16px; margin-bottom:14px; overflow-x:auto; }
.panel h3 { margin:0 0 8px; font-size:14px; color:#842029; }
table { border-collapse:collapse; width:100%; font-size:13px; }
th { text-align:left; font-size:11px; color:#6c757d; font-weight:normal;
     padding:4px 8px; border-bottom:1px solid #dee2e6; white-space:nowrap; }
td { padding:7px 8px; border-bottom:1px solid #eef0f2; }
tr.alerta-fila { cursor:pointer; }
tr.alerta-fila:hover { background:#fff5f5; }
.cod { font-weight:bold; }
.dim { color:#6c757d; }
.filtros { display:flex; gap:14px; flex-wrap:wrap; margin-bottom:14px; }
.filtros label { display:block; font-size:11px; color:#6c757d; margin-bottom:3px; }
.filtros select { font-size:14px; padding:6px 10px; border:1px solid #ced4da;
                  border-radius:6px; min-width:200px; background:#fff; }
.equipo { background:#fff; border:1px solid #dee2e6; border-radius:8px;
          padding:16px 18px; margin-bottom:18px; }
.equipo.rojo { border-left:6px solid #dc3545; }
.equipo.verde { border-left:6px solid #198754; }
.equipo h2 { margin:0 0 2px; font-size:18px; }
.sub { color:#6c757d; font-size:13px; margin-bottom:12px; }
.obs { font-size:13px; background:#f8f9fa; border-radius:6px; padding:8px 12px;
       margin:10px 0; }
.tiles { display:flex; gap:10px; flex-wrap:wrap; margin:10px 0 8px; }
.tile { flex:1; min-width:150px; border:1px solid #dee2e6; border-radius:8px;
        padding:9px 12px; }
.tile.rev { background:#f8d7da; border-color:#f5c2c7; }
.tile .p { font-size:11px; color:#6c757d; }
.tile .v { font-size:18px; font-weight:bold; }
.tile .l { font-size:11px; color:#6c757d; }
.mon { font-size:13px; color:#41464b; margin:6px 0 12px; }
.mon b { color:#212529; }
details { border-top:1px solid #eef0f2; padding:8px 0; }
summary { cursor:pointer; font-size:13px; font-weight:bold; color:#41464b; }
details[open] > summary { margin-bottom:8px; }
.badge { padding:2px 8px; border-radius:10px; font-size:11px; font-weight:bold;
         white-space:nowrap; }
.ficha { display:flex; flex-wrap:wrap; gap:8px; }
.ficha .it { background:#eef2ff; border:1px solid #dfe3ff; border-radius:6px;
             padding:4px 9px; font-size:12px; color:#41464b; }
.ficha .it b { color:#3b3f5c; }
.graficos { display:flex; gap:12px; flex-wrap:wrap; }
.graficos > div { flex:1; min-width:260px; }
.hist-wrap { overflow-x:auto; }
td.rev { background:#f8d7da; color:#842029; font-weight:bold; }
.pie { text-align:center; color:#adb5bd; font-size:12px; margin-top:20px; }
"""

# JavaScript: desplegables + clic en una alerta para saltar a su equipo.
SCRIPT = """
const paneles = Array.from(document.querySelectorAll('.equipo'));
const selSub = document.getElementById('selSub');
const selEq  = document.getElementById('selEquipo');

function llenarSubs() {
  [...new Set(paneles.map(p => p.dataset.sub))].forEach(s => {
    const o = document.createElement('option');
    o.value = s; o.textContent = s; selSub.appendChild(o);
  });
}
function llenarEquipos(sub) {
  selEq.innerHTML = '';
  paneles.filter(p => p.dataset.sub === sub).forEach(p => {
    const o = document.createElement('option');
    o.value = p.dataset.codigo;
    o.textContent = (p.classList.contains('rojo') ? '\\u26A0 ' : '') +
                    p.dataset.codigo + ' (' + p.dataset.tipo + ')';
    selEq.appendChild(o);
  });
}
function mostrar(sub, cod) {
  paneles.forEach(p => {
    p.style.display = (p.dataset.sub === sub && p.dataset.codigo === cod) ? 'block' : 'none';
  });
}
function ir(sub, cod) {
  selSub.value = sub; llenarEquipos(sub); selEq.value = cod; mostrar(sub, cod);
}
selSub.addEventListener('change', () => { llenarEquipos(selSub.value); mostrar(selSub.value, selEq.value); });
selEq.addEventListener('change', () => mostrar(selSub.value, selEq.value));
document.querySelectorAll('.alerta-fila').forEach(r =>
  r.addEventListener('click', () => {
    ir(r.dataset.sub, r.dataset.codigo);
    document.getElementById('filtros').scrollIntoView({behavior:'smooth'});
  }));

llenarSubs();
// al abrir, se muestra el primer equipo en alerta (o el primero de la lista)
const inicial = paneles.find(p => p.classList.contains('rojo')) || paneles[0];
if (inicial) ir(inicial.dataset.sub, inicial.dataset.codigo);
"""

# Parametros con limite que se muestran como "tarjetas" en la ultima visita
CLAVE = ["Voltaje VDC", "Voltaje AC", "Resistencia tierra", "Corriente de fuga"]


def _tendencia_flecha(texto):
    if texto.startswith("subiendo"):
        return "&#9650; sube"
    if texto.startswith("bajando"):
        return "&#9660; baja"
    if texto.startswith("estable"):
        return "&#9654; estable"
    return "&mdash;"


def _lim_texto(m):
    if not m.tiene_limite():
        return ""
    return f'limite {"max" if m.criterio == "maximo" else "min"} {m.limite}'


def _val(v, parametro, vacio="&ndash;"):
    m = v.medicion(parametro)
    return vacio if m is None or m.valor is None else f"{m.valor}"


def _tabla_historial(e):
    """Todas las visitas (mas reciente primero); resalta celdas en REVISAR."""
    cols = [("Tierra (Ohm)", "Resistencia tierra"), ("Fuga (mA)", "Corriente de fuga"),
            ("VDC (V)", "Voltaje VDC"), ("VAC (V)", "Voltaje AC"),
            ("I A/B/C (A)", None), ("P (MW)", "Potencia Activa"),
            ("Q (MVAR)", "Potencia Reactiva"), ("Oper.", "Num Operaciones")]
    H = ['<div class="hist-wrap"><table><tr><th>Fecha</th><th>Cuadrilla</th>']
    H += [f"<th>{c}</th>" for c, _ in cols] + ["<th>Estado</th></tr>"]
    for v in reversed(e.visitas_ordenadas()):
        H.append(f'<tr><td>{html.escape(str(v.fecha)[:10])}</td>'
                 f'<td>{html.escape(str(v.cuadrilla))}</td>')
        for _, p in cols:
            if p is None:
                val = "/".join(_val(v, f"Corriente Fase {f}") for f in "ABC")
                H.append(f"<td>{val}</td>")
                continue
            m = v.medicion(p)
            rev = m is not None and m.clasificar() == "REVISAR"
            H.append(f'<td class="rev">{_val(v, p)}</td>' if rev else f"<td>{_val(v, p)}</td>")
        H.append(f"<td>{_badge_html('REVISAR' if v.hallazgos() else 'NORMAL')}</td></tr>")
    H.append("</table></div>")
    return "".join(H)


# ------------------------------------------------------------------
# Construccion del HTML
# ------------------------------------------------------------------
def construir_html():
    hoy = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    wb = load_workbook(agente.ARCHIVO, data_only=True)
    subs = agente.construir_modelo(wb)
    equipos = agente.todos_los_equipos(subs)
    hallazgos = agente.todos_los_hallazgos(subs)
    visitas = [v for e in equipos for v in e.visitas]
    codigos_alerta = {e.codigo for e, v, m in hallazgos}
    esc = html.escape

    H = []
    H.append('<!doctype html><html lang="es"><head><meta charset="utf-8">')
    H.append('<meta name="viewport" content="width=device-width, initial-scale=1">')
    H.append('<title>Agente Buchholz - Dashboard de Mantenimiento</title>')
    H.append(f'<style>{ESTILOS}</style></head><body>')
    H.append('<div class="barra"><h1>&#9889; Agente Buchholz &middot; Mantenimiento predictivo</h1>'
             f'<span>Generado: {hoy}</span></div>')
    H.append('<div class="contenido">')

    # ---- resumen en una sola fila ----
    n_alerta = len(codigos_alerta)
    H.append('<div class="resumen">')
    for n, et, cls in [(len(equipos), f"Equipos &middot; {len(subs)} subestaciones", ""),
                       (n_alerta, "En alerta", "rojo" if n_alerta else "verde"),
                       (len(hallazgos), "Hallazgos", "rojo" if hallazgos else "verde"),
                       (len(visitas), "Visitas registradas", "")]:
        H.append(f'<div class="stat {cls}"><div class="n">{n}</div><div class="e">{et}</div></div>')
    H.append('</div>')

    # ---- alertas activas: lo primero que se ve ----
    if hallazgos:
        H.append('<div class="panel"><h3>&#9888;&#65039; Alertas activas '
                 '<span class="dim" style="font-weight:normal;">(clic para ver el equipo)</span></h3>'
                 '<table><tr><th>Equipo</th><th>Subestacion</th><th>Parametro</th>'
                 '<th>Medido</th><th>Limite</th><th>Tendencia</th><th>Visita</th></tr>')
        for e, v, m in hallazgos:
            H.append(f'<tr class="alerta-fila" data-sub="{esc(e.subestacion, True)}" '
                     f'data-codigo="{esc(e.codigo, True)}">'
                     f'<td class="cod">{esc(e.codigo)}</td><td>{esc(e.subestacion)}</td>'
                     f'<td>{esc(m.parametro)}</td>'
                     f'<td><b style="color:#842029;">{m.valor} {esc(m.unidad)}</b></td>'
                     f'<td class="dim">{_lim_texto(m).replace("limite ", "")}</td>'
                     f'<td>{_tendencia_flecha(e.tendencia(m.parametro))}</td>'
                     f'<td class="dim">{esc(str(v.fecha)[:10])}</td></tr>')
        H.append('</table></div>')
    else:
        H.append('<div class="ok">&#9989; Todos los equipos en estado NORMAL.</div>')

    # ---- navegacion ----
    H.append('<div class="filtros" id="filtros">'
             '<div><label>Subestacion</label><select id="selSub"></select></div>'
             '<div><label>Equipo</label><select id="selEquipo"></select></div></div>')

    # ---- un panel por equipo (el JS muestra solo uno) ----
    for e in equipos:
        en_alerta = e.codigo in codigos_alerta
        v = e.ultima_visita()
        H.append(f'<div class="equipo {"rojo" if en_alerta else "verde"}" '
                 f'data-sub="{esc(e.subestacion, True)}" data-codigo="{esc(e.codigo, True)}" '
                 f'data-tipo="{esc(e.tipo, True)}" style="display:none">')
        H.append(f'<h2>{esc(e.codigo)} {_badge_html("REVISAR" if en_alerta else "NORMAL")}</h2>')
        H.append(f'<div class="sub">{esc(e.tipo)} &mdash; {esc(e.subestacion)}'
                 + (f' &middot; ultima visita {esc(str(v.fecha)[:10])} '
                    f'(cuadrilla {esc(str(v.cuadrilla))})' if v else '') + '</div>')

        if v is None:
            H.append('<div class="obs">&#8505;&#65039; Sin visitas registradas todavia. '
                     f'Agrega filas en la hoja <b>{esc(e.codigo)}</b> del Excel.</div>')
        else:
            if v.observacion:
                H.append(f'<div class="obs">&#128221; {esc(v.observacion)}</div>')

            # tarjetas: parametros con limite
            H.append('<div class="tiles">')
            for p in CLAVE:
                m = v.medicion(p)
                if m is None:
                    continue
                rev = m.clasificar() == "REVISAR"
                valor = "&ndash;" if m.valor is None else f"{m.valor} {esc(m.unidad)}"
                H.append(f'<div class="tile{" rev" if rev else ""}"><div class="p">{esc(p)}</div>'
                         f'<div class="v">{valor}</div>'
                         f'<div class="l">{_lim_texto(m)} {_badge_html(m.clasificar())}</div></div>')
            H.append('</div>')

            # monitoreo en una sola linea
            H.append('<div class="mon">Monitoreo &mdash; Corriente A/B/C: '
                     f'<b>{"/".join(_val(v, f"Corriente Fase {f}") for f in "ABC")} A</b> &middot; '
                     f'P: <b>{_val(v, "Potencia Activa")} MW</b> &middot; '
                     f'Q: <b>{_val(v, "Potencia Reactiva")} MVAR</b> &middot; '
                     f'Operaciones: <b>{_val(v, "Num Operaciones")}</b></div>')

            # detalle plegado
            H.append('<details><summary>Historial de visitas '
                     f'({len(e.visitas)})</summary>{_tabla_historial(e)}</details>')
            pngs = graficos.graficos_equipo(e)
            H.append('<details><summary>Graficos del historico</summary><div class="graficos">'
                     + "".join(f'<div>{_img(pngs[k])}</div>' for k in ("tierra", "corrientes", "potencias"))
                     + '</div></details>')

        H.append(f'<details><summary>Ficha tecnica</summary>{_ficha_html(e)}</details>')
        H.append('</div>')   # fin .equipo

    H.append('<div class="pie">Agente Buchholz &middot; Mantenimiento Predictivo '
             '&middot; Jose David Perez Munguia &mdash; UTH 2026.4</div>')
    H.append('</div>')                       # fin .contenido
    H.append(f'<script>{SCRIPT}</script>')
    H.append('</body></html>')
    return "\n".join(H)


def main():
    html = construir_html()
    with open(SALIDA, "w", encoding="utf-8") as f:
        f.write(html)
    print("Dashboard generado:", SALIDA)
    webbrowser.open("file:///" + SALIDA.replace("\\", "/"))


if __name__ == "__main__":
    main()

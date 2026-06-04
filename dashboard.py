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
.barra { background:#0d6efd; color:#fff; padding:18px 24px; }
.barra h1 { margin:0; font-size:20px; }
.barra p { margin:4px 0 0; opacity:.9; font-size:13px; }
.contenido { max-width:920px; margin:20px auto; padding:0 16px; }
.cards { display:flex; gap:12px; flex-wrap:wrap; margin-bottom:18px; }
.card { flex:1; min-width:130px; background:#fff; border:1px solid #dee2e6;
        border-radius:8px; padding:14px; text-align:center; }
.card .n { font-size:26px; font-weight:bold; }
.card .e { font-size:12px; color:#6c757d; }
.alerta { background:#f8d7da; color:#842029; border:1px solid #f5c2c7;
          padding:12px 16px; border-radius:8px; margin-bottom:18px; }
.filtros { display:flex; gap:14px; flex-wrap:wrap; align-items:flex-end;
           background:#fff; border:1px solid #dee2e6; border-radius:8px;
           padding:14px; margin-bottom:18px; }
.filtros label { display:block; font-size:12px; color:#6c757d; margin-bottom:4px; }
.filtros select { font-size:15px; padding:7px 10px; border:1px solid #ced4da;
                  border-radius:6px; min-width:200px; background:#fff; }
.equipo { background:#fff; border:1px solid #dee2e6; border-radius:8px;
          padding:18px; margin-bottom:18px; }
.equipo.rojo { border-left:6px solid #dc3545; }
.equipo.verde { border-left:6px solid #198754; }
.equipo h2 { margin:0 0 2px; font-size:18px; }
.sub { color:#6c757d; font-size:13px; margin-bottom:12px; }
.ficha { display:flex; flex-wrap:wrap; gap:8px; margin:2px 0 16px; }
.ficha .it { background:#eef2ff; border:1px solid #dfe3ff; border-radius:6px;
             padding:4px 9px; font-size:12px; color:#41464b; }
.ficha .it b { color:#3b3f5c; }
table { border-collapse:collapse; width:100%; font-size:14px; margin-bottom:14px; }
td { padding:6px 10px; border:1px solid #dee2e6; }
.badge { padding:2px 8px; border-radius:10px; font-size:12px; font-weight:bold; }
.graficos { display:flex; gap:14px; flex-wrap:wrap; }
.graficos > div { flex:1; min-width:280px; }
.pie { text-align:center; color:#adb5bd; font-size:12px; margin-top:20px; }
"""

# JavaScript: llena los desplegables y muestra UN equipo a la vez.
SCRIPT = """
const paneles = Array.from(document.querySelectorAll('.equipo'));
const selSub = document.getElementById('selSub');
const selEq  = document.getElementById('selEquipo');

function subestaciones() {
  return [...new Set(paneles.map(p => p.dataset.sub))];
}
function llenarSubs() {
  subestaciones().forEach(s => {
    const o = document.createElement('option');
    o.value = s; o.textContent = s; selSub.appendChild(o);
  });
}
function llenarEquipos(sub) {
  selEq.innerHTML = '';
  paneles.filter(p => p.dataset.sub === sub).forEach(p => {
    const o = document.createElement('option');
    o.value = p.dataset.codigo;
    o.textContent = p.dataset.codigo + ' (' + p.dataset.tipo + ')';
    selEq.appendChild(o);
  });
}
function mostrar(sub, cod) {
  paneles.forEach(p => {
    p.style.display = (p.dataset.sub === sub && p.dataset.codigo === cod) ? 'block' : 'none';
  });
}
selSub.addEventListener('change', () => {
  llenarEquipos(selSub.value);
  mostrar(selSub.value, selEq.value);
});
selEq.addEventListener('change', () => mostrar(selSub.value, selEq.value));

// inicializar
llenarSubs();
selSub.selectedIndex = 0;
llenarEquipos(selSub.value);
mostrar(selSub.value, selEq.value);
"""


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

    H = []
    H.append('<!doctype html><html lang="es"><head><meta charset="utf-8">')
    H.append('<meta name="viewport" content="width=device-width, initial-scale=1">')
    H.append('<title>Agente Buchholz - Dashboard de Mantenimiento</title>')
    H.append(f'<style>{ESTILOS}</style></head><body>')

    # barra superior
    H.append('<div class="barra">'
             '<h1>&#9889; Agente Buchholz</h1>'
             '<p style="margin:2px 0 0;font-size:14px;">Mantenimiento predictivo '
             'de subestaciones electricas</p>'
             f'<p>Dashboard generado: {hoy} &middot; '
             'para actualizar, vuelve a ejecutar <b>python dashboard.py</b></p></div>')

    H.append('<div class="contenido">')

    # tarjetas de resumen (globales)
    H.append('<div class="cards">')
    for n, e in [(len(subs), "Subestaciones"), (len(equipos), "Equipos"),
                 (len(visitas), "Visitas"), (len(hallazgos), "Hallazgos")]:
        H.append(f'<div class="card"><div class="n">{n}</div>'
                 f'<div class="e">{e}</div></div>')
    H.append('</div>')

    # banner de alerta (global)
    if hallazgos:
        nombres = ", ".join(sorted(codigos_alerta))
        H.append(f'<div class="alerta">&#9888;&#65039; <b>{len(hallazgos)} hallazgo(s)</b> '
                 f'que requieren accion en: <b>{nombres}</b></div>')
    else:
        H.append('<div class="alerta" style="background:#d1e7dd;color:#0f5132;'
                 'border-color:#badbcc;">&#9989; Todos los equipos en estado NORMAL.</div>')

    # desplegables de navegacion
    H.append('<div class="filtros">'
             '<div><label>Subestacion</label><select id="selSub"></select></div>'
             '<div><label>Equipo</label><select id="selEquipo"></select></div>'
             '</div>')

    # una tarjeta por equipo (todas se generan; el JS muestra solo una)
    for e in equipos:
        en_alerta = e.codigo in codigos_alerta
        clase = "rojo" if en_alerta else "verde"
        v = e.ultima_visita()
        H.append(f'<div class="equipo {clase}" data-sub="{e.subestacion}" '
                 f'data-codigo="{e.codigo}" data-tipo="{e.tipo}" style="display:none">')
        H.append(f'<h2>{("&#128308;" if en_alerta else "&#128994;")} {e.codigo}</h2>')
        H.append(f'<div class="sub">{e.tipo} &mdash; {e.subestacion}'
                 + (f' &middot; ultima visita: {str(v.fecha)[:10]} '
                    f'(cuadrilla {v.cuadrilla})' if v else '') + '</div>')

        # ficha tecnica (datos de placa)
        H.append(_ficha_html(e))

        if v is None:
            # equipo registrado pero sin visitas todavia
            H.append('<div class="sub" style="background:#fff3cd;color:#664d03;'
                     'padding:10px 14px;border-radius:6px;">&#8505;&#65039; '
                     'Sin visitas registradas todavia. Agrega filas en la hoja '
                     f'<b>{e.codigo}</b> del Excel para ver su estado y graficos.</div>')
        else:
            # tabla de estado de la ultima visita
            H.append('<table>')
            for columna, parametro, unidad in agente.COLUMNAS:
                m = v.medicion(parametro)
                if m is None:
                    continue
                valor = "(sin dato)" if m.valor is None else f"{m.valor} {m.unidad}"
                H.append(f'<tr><td style="width:45%">{parametro}</td>'
                         f'<td>{valor}</td>'
                         f'<td style="text-align:center;width:120px">'
                         f'{_badge_html(m.clasificar())}</td></tr>')
            H.append('</table>')
            if v.observacion:
                H.append(f'<div class="sub">&#128221; {v.observacion}</div>')

            # graficos del equipo (solo si hay datos)
            pngs = graficos.graficos_equipo(e)
            H.append('<div class="graficos">')
            H.append('<div>' + _img(pngs["tierra"]) + '</div>')
            H.append('<div>' + _img(pngs["corrientes"]) + '</div>')
            H.append('<div>' + _img(pngs["potencias"]) + '</div>')
            H.append('</div>')

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

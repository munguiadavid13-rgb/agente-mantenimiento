"""
graficos.py
-----------
Genera graficos de linea (PNG en memoria) del historico de un equipo,
para incrustarlos dentro del correo HTML del agente.

Usa matplotlib con el backend 'Agg' (sin ventana), asi funciona en
cualquier maquina o servidor aunque no haya pantalla.

Autor: Jose David Perez Munguia - UTH 2026.4
"""

import io
import math

import matplotlib
matplotlib.use("Agg")              # backend sin pantalla (no abre ventanas)
import matplotlib.pyplot as plt    # noqa: E402  (debe ir despues de use())


def _figura_a_png(fig):
    """Convierte una figura de matplotlib en bytes PNG y libera la memoria."""
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", dpi=100, bbox_inches="tight")
    plt.close(fig)                 # importante: cierra la figura para no acumular
    return buffer.getvalue()


# Cuantos registros recientes mostrar en los graficos. La tabla se ira
# llenando con el tiempo; graficar TODO se veria amontonado y dificil de
# leer, asi que mostramos solo las ultimas N visitas.
MAX_REGISTROS = 5


def _serie(equipo, parametro, maximo=MAX_REGISTROS):
    """
    Devuelve (fechas, valores) del historico de un parametro, recortado a
    las ultimas 'maximo' visitas (las mas recientes) para que el grafico
    no se sature cuando haya muchos datos.
    """
    datos = equipo.historico(parametro)        # ya viene ordenado por fecha
    datos = datos[-maximo:]                     # solo las ultimas N
    fechas = [str(f)[:10] for f, _ in datos]    # solo AAAA-MM-DD
    valores = [v for _, v in datos]
    return fechas, valores


def _estilo_ejes(ax, titulo, eje_y):
    """Aplica el estilo comun a todos los graficos."""
    ax.set_title(titulo, fontsize=11, fontweight="bold")
    ax.set_ylabel(eje_y, fontsize=9)
    ax.grid(True, alpha=.3)
    ax.legend(fontsize=8)
    ax.tick_params(axis="x", labelsize=8, rotation=20)
    ax.tick_params(axis="y", labelsize=8)


def grafico_corrientes(equipo):
    """Linea con el historico de Corriente Fase A, B y C."""
    fig, ax = plt.subplots(figsize=(6, 3))
    for fase in ["A", "B", "C"]:
        fechas, valores = _serie(equipo, f"Corriente Fase {fase}")
        if valores:
            ax.plot(fechas, valores, marker="o", label=f"Fase {fase}")
    _estilo_ejes(ax, f"Corrientes por fase - {equipo.codigo}", "Corriente (A)")
    return _figura_a_png(fig)


def grafico_potencias(equipo):
    """Linea con Potencia Activa, Reactiva y Aparente (calculada S=sqrt(P^2+Q^2))."""
    fig, ax = plt.subplots(figsize=(6, 3))
    fechas, activa = _serie(equipo, "Potencia Activa")
    _, reactiva = _serie(equipo, "Potencia Reactiva")
    if activa:
        ax.plot(fechas, activa, marker="o", label="Activa (MW)")
    if reactiva:
        ax.plot(fechas, reactiva, marker="o", label="Reactiva (MVAR)")
    # Aparente: solo si tenemos activa y reactiva alineadas en las mismas visitas
    if activa and reactiva and len(activa) == len(reactiva):
        aparente = [math.sqrt(p * p + q * q) for p, q in zip(activa, reactiva)]
        ax.plot(fechas, aparente, marker="o", linestyle="--", label="Aparente (MVA)")
    _estilo_ejes(ax, f"Potencias - {equipo.codigo}", "Potencia")
    return _figura_a_png(fig)


def grafico_tierra(equipo, limite=5):
    """Linea con la resistencia de tierra y una linea roja con el limite."""
    fechas, valores = _serie(equipo, "Resistencia tierra")
    fig, ax = plt.subplots(figsize=(6, 3))
    if valores:
        ax.plot(fechas, valores, marker="o", color="#0d6efd", label="Tierra (Ohm)")
    ax.axhline(limite, color="#dc3545", linestyle="--",
               label=f"Limite ({limite} Ohm)")
    _estilo_ejes(ax, f"Resistencia de tierra - {equipo.codigo}", "Ohm")
    return _figura_a_png(fig)


def graficos_equipo(equipo):
    """Devuelve {nombre: png_bytes} con los tres graficos del equipo."""
    return {
        "tierra": grafico_tierra(equipo),
        "corrientes": grafico_corrientes(equipo),
        "potencias": grafico_potencias(equipo),
    }

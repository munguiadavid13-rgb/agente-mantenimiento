"""
modelos.py
----------
Clases POO (Programacion Orientada a Objetos) del agente de mantenimiento.

Relacion entre clases (composicion: unas contienen a otras):

    Subestacion  ->  contiene varios  Equipo
    Equipo       ->  tiene historial de  Visita
    Visita       ->  contiene varias  Medicion

Autor: Jose David Perez Munguia - UTH 2026.4
"""


# ============================================================
# 1) MEDICION: un dato medido que sabe clasificarse solo
# ============================================================
class Medicion:
    def __init__(self, parametro, valor, unidad, criterio=None, limite=None):
        self.parametro = parametro    # ej "Resistencia tierra"
        self.valor = valor            # el numero medido
        self.unidad = unidad          # ej "Ohm"
        self.criterio = criterio      # "minimo", "maximo" o None (sin limite)
        self.limite = limite          # numero contra el que se compara (o None)

    def tiene_limite(self):
        """True si este parametro se compara contra un limite."""
        return self.criterio is not None and self.limite is not None

    def clasificar(self):
        if self.valor is None:
            return "SIN DATO"
        if not self.tiene_limite():
            return "MONITOREO"        # parametro sin limite -> solo se registra
        if self.criterio == "minimo":
            return "REVISAR" if self.valor <= self.limite else "NORMAL"
        if self.criterio == "maximo":
            return "REVISAR" if self.valor >= self.limite else "NORMAL"
        return "NORMAL"

    def __str__(self):
        if self.tiene_limite():
            return f"{self.parametro} = {self.valor} {self.unidad} -> {self.clasificar()}"
        return f"{self.parametro} = {self.valor} {self.unidad}"


# ============================================================
# 2) VISITA: una inspeccion en una fecha, con sus mediciones
# ============================================================
class Visita:
    def __init__(self, fecha, cuadrilla, observacion=""):
        self.fecha = fecha
        self.cuadrilla = cuadrilla
        self.observacion = observacion
        self.mediciones = []          # lista de objetos Medicion

    def agregar(self, medicion):
        self.mediciones.append(medicion)

    def medicion(self, parametro):
        """Devuelve la Medicion de un parametro (o None si no esta)."""
        for m in self.mediciones:
            if m.parametro == parametro:
                return m
        return None

    def hallazgos(self):
        """Mediciones que salieron en REVISAR en esta visita."""
        return [m for m in self.mediciones if m.clasificar() == "REVISAR"]


# ============================================================
# 3) EQUIPO: un interruptor/restaurador con su historial
# ============================================================
class Equipo:
    def __init__(self, codigo, tipo, subestacion):
        self.codigo = codigo               # ej "CTE-32L8"
        self.tipo = tipo                   # "interruptor" o "restaurador"
        self.subestacion = subestacion     # nombre de la S/E
        self.visitas = []                  # historial de visitas

    def agregar_visita(self, visita):
        self.visitas.append(visita)

    def visitas_ordenadas(self):
        return sorted(self.visitas, key=lambda v: v.fecha)

    def ultima_visita(self):
        vs = self.visitas_ordenadas()
        return vs[-1] if vs else None

    def historico(self, parametro):
        """Lista de (fecha, valor) de un parametro a lo largo del tiempo."""
        datos = []
        for v in self.visitas_ordenadas():
            m = v.medicion(parametro)
            if m and m.valor is not None:
                datos.append((v.fecha, m.valor))
        return datos

    def tendencia(self, parametro):
        """Compara las dos ultimas lecturas: sube, baja o estable."""
        datos = self.historico(parametro)
        if len(datos) < 2:
            return "sin historico suficiente"
        anterior, ultimo = datos[-2][1], datos[-1][1]
        if ultimo > anterior:
            return f"subiendo ({anterior} -> {ultimo})"
        if ultimo < anterior:
            return f"bajando ({anterior} -> {ultimo})"
        return f"estable ({ultimo})"


# ============================================================
# 4) SUBESTACION: agrupa a sus equipos
# ============================================================
class Subestacion:
    def __init__(self, nombre):
        self.nombre = nombre
        self.equipos = {}             # codigo -> Equipo

    def obtener_equipo(self, codigo, tipo):
        """Devuelve el equipo; si no existe todavia, lo crea."""
        if codigo not in self.equipos:
            self.equipos[codigo] = Equipo(codigo, tipo, self.nombre)
        return self.equipos[codigo]

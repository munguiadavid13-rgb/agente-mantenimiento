# Agente de Mantenimiento Predictivo de Subestaciones Eléctricas

Proyecto final del curso de **Programación** — Maestría en Automatización Industrial,
Universidad Tecnológica de Honduras (UTH), periodo **2026.4**.

Agente que analiza las mediciones de las visitas de mantenimiento a interruptores y
restauradores de subestaciones, **clasifica** los parámetros críticos, lleva un
**histórico** para detectar tendencias (mantenimiento *predictivo*) y **notifica**
automáticamente los hallazgos por **correo** y **Telegram**.

---

## 📋 Descripción

En las visitas de mantenimiento se toman mediciones de cada equipo (voltajes, resistencia
de puesta a tierra, corrientes, potencias, número de operaciones). Este agente:

1. **Lee** esas mediciones desde un archivo Excel (la base de datos del proyecto).
2. **Clasifica** los parámetros que tienen un límite definido en `NORMAL` o `REVISAR`.
3. **Acumula un histórico** por equipo y calcula la **tendencia** de la resistencia de
   tierra para anticipar problemas antes de que ocurran.
4. **Genera un reporte** legible y **notifica** los hallazgos por correo (informe completo)
   y Telegram (alerta corta), indicando qué equipo revisar y con qué cuadrilla consultar.

---

## ✨ Características

- **Programación Orientada a Objetos (POO):** 4 clases en composición.
- **Dos tipos de parámetro:** con límite (se clasifican) y de solo monitoreo (se registran
  para el histórico). El comportamiento se decide editando una tabla, **sin tocar código**.
- **Histórico y tendencia:** compara las lecturas a lo largo del tiempo.
- **Tres canales de salida:** correo (Gmail, HTML con tablas y gráficos), Telegram (alerta
  con formato) y **dashboard web** (HTML estático con navegación por desplegables).
- **Múltiples destinatarios:** la alerta puede enviarse a varios correos y varios chats de
  Telegram a la vez.
- **Correo enfocado:** cuando hay una alerta, el correo se concentra **solo en el equipo
  afectado**, con el motivo y los datos de su última visita (los gráficos quedan en el
  dashboard para no recargar el correo).
- **Ficha técnica por equipo:** cada equipo guarda sus datos de placa (serie, marca, modelo,
  tipo, I nominal, tensión, año, ICC), que se muestran en el dashboard.
- **Base de datos en Excel** con **una hoja por equipo** (datos separados, sin mezclar) y
  menús desplegables, fácil de llenar en campo.

---

## 🧱 Arquitectura (clases POO)

Las clases se relacionan por **composición** (unas contienen a otras):

```
Subestacion          (Ceiba Termica, San Isidro, Isletas)
   └── Equipo         (CTE-32L8 interruptor, SIS-32L30 restaurador, ...)
          └── Visita  (una por fecha: cuadrilla, observación)
                 └── Medicion  (Voltaje VDC, Resistencia tierra, Corrientes, ...)
```

| Clase | Responsabilidad |
|-------|-----------------|
| `Medicion` | Guarda un valor medido y sabe **clasificarse** (`clasificar()`). |
| `Visita` | Una inspección en una fecha; agrupa sus mediciones y conoce sus **hallazgos**. |
| `Equipo` | Un interruptor/restaurador; guarda su **historial** de visitas y calcula **tendencias**. |
| `Subestacion` | Agrupa los equipos de una misma subestación. |

---

## 📂 Estructura de archivos

```
agente-mantenimiento/
├── modelos.py                       # Las 4 clases POO
├── agente.py                        # Programa principal: lee, clasifica, reporta, notifica
├── graficos.py                      # Genera los gráficos de línea (matplotlib)
├── dashboard.py                     # Genera el dashboard web estático (dashboard.html)
├── agregar_equipo.py                # Agrega un equipo (fila índice + hoja) sin desajustes
├── revisar.py                       # Revisión automática: refresca el dashboard (sin notificar)
├── notificaciones.py                # Envío por correo (Gmail) y Telegram
├── config.py                        # Credenciales privadas (NO se comparte)
├── config.example.py                # Plantilla de credenciales (sí se comparte)
├── Mantenimiento_Subestaciones.xlsx # Base de datos (índice + 1 hoja por equipo + Limites)
├── requirements.txt                 # Librerías necesarias
├── .gitignore                       # Protege config.py
└── README.md                        # Este archivo
```

### 🗂️ Estructura del Excel (base de datos)

Para mantener el orden y que **los datos de cada equipo no se mezclen**, el libro tiene:

| Hoja | Contenido |
|------|-----------|
| `Equipos` | Índice: lista de subestaciones, código de cada interruptor y su tipo. |
| `CTE-32L08`, `SIS-32L30`, ... | **Una hoja por equipo**: arriba su **ficha técnica** (serie, marca, modelo, tipo, I nominal, tensión, año, ICC) y debajo la tabla de visitas (fecha, cuadrilla, mediciones, observación). |
| `Limites` | Parámetros con límite y su criterio (`minimo`/`maximo`). |

> **Para agregar un equipo nuevo**, lo más seguro es usar el script `agregar_equipo.py`
> (ver abajo): crea la fila del índice y la hoja de datos juntas, evitando desajustes.
> Si lo haces a mano, recuerda que cada equipo necesita **dos cosas**: su fila en `Equipos`
> y una hoja con el **mismo código exacto** (mismas columnas que las demás).

#### Agregar un equipo con el script

```bash
python agregar_equipo.py "Coyoles Central" "CCE-32L43" "Restaurador"
```

O sin argumentos (`python agregar_equipo.py`) para que te lo pregunte paso a paso. Luego
abres el Excel y llenas las visitas en la hoja del equipo.

---

## ⚙️ Requisitos

- **Python 3.12** o superior
- **openpyxl** (para leer Excel)
- **matplotlib** (para los gráficos de línea del correo)

---

## 🚀 Instalación

1. Instala las dependencias:

   ```bash
   pip install -r requirements.txt
   ```

2. Crea tu archivo de configuración a partir de la plantilla:

   ```bash
   copy config.example.py config.py      # Windows
   ```

3. Abre `config.py` y completa tus credenciales (ver siguiente sección).

---

## 🔐 Configuración (`config.py`)

### Correo (Gmail)
Gmail exige una **Contraseña de aplicación** (no la contraseña normal):
1. Activa la Verificación en 2 pasos en tu cuenta de Google.
2. Genera una contraseña de aplicación en `myaccount.google.com/apppasswords`.
3. Colócala en `CORREO_APP_PASSWORD`, junto con el remitente y el destinatario.

### Telegram
1. En Telegram, crea un bot con **@BotFather** (`/newbot`) y copia el **token**.
2. Envíale un mensaje a tu bot.
3. El `chat_id` se obtiene consultando
   `https://api.telegram.org/bot<TOKEN>/getUpdates`.

> ⚠️ **Seguridad:** `config.py` contiene secretos y está listado en `.gitignore`
> para que **nunca** se suba a internet. Comparte solo `config.example.py`.

---

## ▶️ Uso

1. Llena la **hoja del equipo** correspondiente en el Excel con los datos de cada visita
   (fecha, cuadrilla, valores medidos y observación). Cada equipo tiene su propia hoja.
2. Ejecuta el agente:

   ```bash
   python agente.py
   ```

El agente imprime el reporte en pantalla y, **si encuentra hallazgos**, los envía por
correo y Telegram automáticamente.

### 🌐 Dashboard web (estático)

Además del correo y Telegram, hay un **tablero web** que muestra el estado de los equipos:

```bash
python dashboard.py
```

Esto genera el archivo **`dashboard.html`** y lo abre en el navegador. No levanta ningún
servidor (no queda nada corriendo en segundo plano). Muestra **un equipo a la vez**, con
dos menús desplegables para navegar: uno por **subestación** y otro por **equipo**. Para
ver datos actualizados, vuelve a ejecutar `python dashboard.py`.

### 🔄 Revisión automática (cada 50 min)

`revisar.py` relee el Excel y **regenera el dashboard** sin abrir el navegador y **sin
enviar notificaciones** (pensado para correr en segundo plano). Deja un registro en
`revisiones_log.txt`.

```bash
python revisar.py
```

En Windows se programó con el **Programador de tareas** para que corra solo cada 50 minutos
(tarea `AgenteMantenimiento_Revision`). Comandos útiles (en CMD):

```bat
schtasks /Query  /TN "AgenteMantenimiento_Revision"            :: ver estado
schtasks /Change /TN "AgenteMantenimiento_Revision" /DISABLE   :: pausar
schtasks /Change /TN "AgenteMantenimiento_Revision" /ENABLE    :: reanudar
schtasks /Delete /TN "AgenteMantenimiento_Revision" /F         :: eliminar
```

> La revisión automática solo notifica **cuando algo cambia** (ver abajo); si nada cambia,
> no manda nada. Una notificación "a la fuerza" (siempre que haya hallazgos) se hace con
> `python agente.py`.

### 🔔 Notificar solo cuando algo cambia

Para no saturar, la revisión automática recuerda qué alertas ya conocía (en
`estado_alertas.json`) y **solo envía correo + Telegram cuando hay un cambio**: una alerta
**nueva** (un equipo entra en REVISAR) o una **resuelta** (deja de estarlo). Si las alertas
son las mismas que la última vez, no envía nada.

- La identidad de una alerta es `EQUIPO|parámetro` (p. ej. `ISL-32L44|Resistencia tierra`).
- Lo gestiona `agente.notificar_si_cambio(...)`, que usa `revisar.py` en cada corrida.
- `python agente.py` (manual) **siempre** notifica si hay hallazgos, sin importar el estado.

---

## 📏 Lógica de clasificación

Los límites viven en la hoja **Limites** del Excel. Cada parámetro se evalúa según su criterio:

| Parámetro | Criterio | Regla |
|-----------|----------|-------|
| Voltaje VDC | `minimo` | 🔴 REVISAR si el valor **≤ 20 V** (desgaste de baterías) |
| Voltaje AC | `minimo` | 🔴 REVISAR si el valor **≤ 119 V** (alimentación AC) |
| Resistencia de tierra | `maximo` | 🔴 REVISAR si el valor **≥ 5 Ω** |

Los demás parámetros (corriente de fuga, corrientes por fase, potencias, número de
operaciones) **no tienen límite**: son de **solo monitoreo** y sirven para el análisis
histórico. La **corriente de fuga (mA)** acompaña a la resistencia de tierra en la sección
de puesta a tierra; si más adelante quieres clasificarla, basta agregar una fila con su
criterio `maximo` en la hoja **Limites**.

> 💡 Para cambiar un límite o agregar un parámetro, basta con editar la hoja **Limites**
> del Excel; el código no se modifica.

---

## 📊 Ejemplo de salida

```
 RESUMEN
   Subestaciones : 3
   Equipos       : 3
   Visitas       : 10
   Hallazgos     : 1  (requieren accion)

 1) HALLAZGOS QUE REQUIEREN ACCION
   [!] ISL-32L44  (restaurador)
       Subestacion : Isletas
       Fecha       : 2026-06-02        Cuadrilla: TCO1
       Medido      : 5.5 Ohm   (limite maximo: 5)

 3) HISTORICO Y TENDENCIA - Resistencia de tierra (Ohm)
   ISL-32L44 (restaurador) - Isletas
        2025-11-12     2.6 Ohm   NORMAL
        2026-03-08     2.9 Ohm   NORMAL
        2026-06-02     5.5 Ohm   REVISAR  <-- REVISAR
        Tendencia: subiendo (2.9 -> 5.5)
```

---

## 👤 Autor

**Jose David Perez Munguia**
Maestría en Automatización Industrial — UTH 2026.4
Curso de Programación — Prof. PhD(c) Luis Loo

---

## 📝 Notas

- Los datos incluidos en el Excel son **sintéticos** (de ejemplo) para demostrar el
  funcionamiento; deben reemplazarse por las mediciones reales de campo.
- No incluir en el repositorio datos confidenciales reales de la empresa ni credenciales
  de telecontrol.
- El envío de notificaciones desactiva la verificación de certificado SSL porque la red
  corporativa inspecciona el tráfico HTTPS con su propio certificado.
- Los gráficos muestran solo las **últimas 5 visitas** de cada equipo, para que sigan siendo
  legibles a medida que la base de datos crece (configurable en `graficos.py`,
  `MAX_REGISTROS`).

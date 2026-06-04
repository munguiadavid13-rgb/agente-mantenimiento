"""
publicar_dashboard.py
---------------------
Genera el dashboard en docs/index.html para publicarlo en GitHub Pages.

GitHub Pages sirve el contenido de la carpeta doc/ de la rama main en:
  https://munguiadavid13-rgb.github.io/agente-mantenimiento/

Este script NO abre el navegador ni hace git. Solo regenera docs/index.html.
Para que los cambios aparezcan en internet, despues haz:

    python publicar_dashboard.py
    git add docs
    git commit -m "Actualiza dashboard publicado"
    git push

OJO: el dashboard queda PUBLICO. No publiques datos confidenciales reales.
"""

import os

import agente
import dashboard

DESTINO_DIR = os.path.join(agente.CARPETA, "docs")
DESTINO = os.path.join(DESTINO_DIR, "index.html")


def main():
    os.makedirs(DESTINO_DIR, exist_ok=True)
    contenido = dashboard.construir_html()
    with open(DESTINO, "w", encoding="utf-8") as f:
        f.write(contenido)
    print("Dashboard publicable generado en:", DESTINO)
    print("Ahora sube los cambios:")
    print('  git add docs && git commit -m "Actualiza dashboard" && git push')


if __name__ == "__main__":
    main()

# dashboard-ofsc

Tablero de seguimiento — **Dashboard Avance Operativo (OFSC)**.

## Publicación

El sitio se publica con **GitHub Pages desde la rama `main`**, sirviendo el
archivo **`index.html`**. Esa es la única fuente de verdad del reporte.

> ⚠️ **Importante para el script de generación**
>
> El script local debe escribir el reporte en `index.html` y publicarlo en la
> rama `main`. No usar la rama `principal` (historia independiente que causaba
> conflictos y dejaba el sitio sin actualizar).

### Flujo recomendado del script

```bash
git checkout main
git pull --ff-only origin main

# ...generar el reporte sobre index.html...

git add index.html
git commit -m "Dashboard actualizado $(date '+%d/%m/%Y %H:%M')"
git push origin main
```

El script (`.py`), los datos fuente (`.xlsx`, `.csv`) y `iniciar_dashboard.bat`
son locales y están excluidos en `.gitignore`.

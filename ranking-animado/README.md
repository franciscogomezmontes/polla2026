# Ranking animado — La Mejor Polla Mundialista 2026

Página estática (bar chart race) que muestra la evolución del ranking de
todos los participantes a lo largo del torneo, con escala fija en puntos
y modo de comparación entre participantes seleccionados.

## Estructura

```
ranking-animado/
  index.html                      <- página final (no tiene datos hardcodeados)
  data/
    ranking_historico.json        <- datos exportados desde el Admin
  scripts/
    exportar_ranking.py           <- script que genera el JSON
```

## Cómo actualizar los datos después de cada jornada

Cada vez que actualices el Admin con los resultados reales de una jornada,
vuelve a correr el script de exportación para refrescar el JSON:

```bash
python scripts/exportar_ranking.py "Admin_POLLA_FIFA_WORLD_CUP_2026_r00.xlsx" "data/ranking_historico.json"
```

El script:
- Lee la hoja `Posiciones` del Admin (columnas `J0`, `J1`, `J2`, ... una por
  día calendario del torneo).
- Detecta automáticamente cuál es la última jornada con datos reales
  (no necesitas indicarle un número fijo de jornadas).
- No recalcula nada — solo copia los valores que el Admin ya calculó con
  sus propias fórmulas.
- Genera `slug` único por participante (ej. `diego-cadena`), que se usará
  más adelante para las páginas individuales de predicciones.

Después de correr el script, solo necesitas hacer commit y push del
archivo `data/ranking_historico.json` actualizado. El `index.html` no
necesita tocarse nunca para actualizaciones normales de datos.

## Cómo probar localmente

Abrir `index.html` directamente con doble clic en el navegador **no
funciona** para cargar el JSON (los navegadores bloquean `fetch()` sobre
`file://` por seguridad). Para probar localmente, levanta un servidor
simple desde la carpeta `ranking-animado/`:

```bash
python -m http.server 8000
```

Y abre `http://localhost:8000/` en el navegador.

## Ajustar la escala fija del eje X

Actualmente el techo de la escala está fijado en 500 puntos
(`SCALE_MAX` en `index.html`). Conforme avance el torneo y entremos a
fase eliminatoria (donde los puntos por partido son mucho más altos),
es probable que haya que subir este valor. Para cambiarlo, editar la
línea:

```js
var SCALE_MAX = 500;
```

## Publicación

Mientras este código viva en el branch `feature/ranking-animado`, no es
visible para los participantes (GitHub Pages solo publica `main`). Antes
de fusionar a `main`, validar:

- Que los 43 participantes aparezcan correctamente.
- Que el modo "Comparar" funcione con cualquier cantidad de selecciones.
- Que el ranking refleje empates reales (mismos puntos = mismo número de
  ranking).
- Que la velocidad de animación se pueda cambiar en cualquier momento.

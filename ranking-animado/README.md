# Ranking animado y predicciones individuales — La Mejor Polla Mundialista 2026

Dos páginas estáticas que se alimentan del Admin:

1. **`index.html`** — bar chart race con la evolución del ranking de todos
   los participantes, escala fija en puntos y modo de comparación.
2. **`participante.html`** — predicciones completas de un participante
   (fase de grupos, eliminatoria, premios FIFA), accesible haciendo clic
   en cualquier nombre desde el ranking.

## Estructura

```
ranking-animado/
  index.html                          <- ranking animado
  participante.html                   <- predicciones individuales (recibe ?id=<slug>)
  data/
    ranking_historico.json            <- histórico de puntos por jornada
    predicciones/
      _indice.json                    <- lista de todos los slugs disponibles
      <slug-participante>.json        <- un archivo por participante (43 en total)
  scripts/
    exportar_ranking.py               <- genera ranking_historico.json
    exportar_predicciones.py          <- genera data/predicciones/*.json
```

## Cómo actualizar los datos después de cada jornada

Cada vez que actualices el Admin con resultados reales, corre ambos
scripts para refrescar todo:

```bash
python scripts/exportar_ranking.py "Admin_POLLA_FIFA_WORLD_CUP_2026_r00.xlsx" "data/ranking_historico.json"
python scripts/exportar_predicciones.py "Admin_POLLA_FIFA_WORLD_CUP_2026_r00.xlsx" "data/predicciones"
```

Ambos scripts:
- Solo leen valores que el Admin ya calculó con sus propias fórmulas — no
  recalculan nada.
- Detectan automáticamente los datos disponibles (jornadas jugadas,
  partidos con resultado real, equipos aún "por definir" en eliminatoria).
- Usan el mismo `slug` por participante en ambos JSON, para que los links
  desde el ranking lleven a la página de predicciones correcta.

Después de correr los scripts, solo necesitas hacer commit y push de la
carpeta `data/` actualizada. Ninguno de los dos `.html` necesita tocarse
para actualizaciones normales de datos.

## Qué muestra participante.html

- **Fase de grupos**: cada partido con predicción, resultado real (si ya
  se jugó) y puntos obtenidos. Los partidos pendientes se muestran
  atenuados. También incluye las predicciones de clasificados (1ro, 2do,
  mejor 3ro) por grupo.
- **Eliminatoria**: las 32 llaves completas (Dieciseisavos a Final) con
  la predicción de marcador y ganador. Como los rivales de varias llaves
  dependen de resultados futuros, se muestran como "Por definir" hasta
  que el Admin los resuelva.
- **Premios FIFA**: las 8 predicciones de premios individuales y de
  equipo.

## Cómo probar localmente

Igual que con el ranking: abrir los `.html` con doble clic **no
funciona** porque bloquea `fetch()` sobre `file://`. Levanta un servidor
simple desde la carpeta `ranking-animado/`:

```bash
python -m http.server 8000
```

Y abre `http://localhost:8000/` en el navegador. Desde ahí, hacer clic en
cualquier nombre del ranking debe llevarte a su página de predicciones.

## Ajustar la escala fija del eje X (ranking animado)

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

- Que los 43 participantes aparezcan correctamente en el ranking.
- Que el modo "Comparar" funcione con cualquier cantidad de selecciones.
- Que el ranking refleje empates reales (mismos puntos = mismo número de
  ranking).
- Que la velocidad de animación se pueda cambiar en cualquier momento.
- Que cada nombre del ranking lleve a la página de predicciones correcta.
- Que las páginas de predicciones muestren correctamente partidos
  jugados vs. pendientes, y las llaves de eliminatoria aún sin definir.


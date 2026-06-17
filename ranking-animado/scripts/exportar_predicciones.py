"""
exportar_predicciones.py

Lee la hoja 'Puntos' del Admin de la Polla Mundialista y genera un archivo
JSON por participante con todas sus predicciones: partidos de grupo,
clasificados de grupo, eliminatoria completa, y premios FIFA.

NO recalcula nada: copia los valores que el Admin ya tiene (predicciones,
resultados reales, puntos obtenidos). El Admin sigue siendo la unica
fuente de verdad.

Uso:
    python exportar_predicciones.py "Admin_POLLA_FIFA_WORLD_CUP_2026_r00.xlsx" "salida/predicciones/"

Genera un archivo <slug>.json por cada participante dentro de la carpeta
de salida indicada, usando el mismo slug que exportar_ranking.py para que
las paginas de ranking y de predicciones queden enlazadas.
"""

import sys
import os
import json
import re
from datetime import datetime, time as dtime
import openpyxl


SHEET = "Puntos"
FILA_NOMBRE_PART = 3
FILA_TIPO_COL = 4
PRIMER_COL_PARTICIPANTE = 11
COLS_POR_BLOQUE = 4

COL_NUM_PARTIDO = 1
COL_GRUPO = 2
COL_FECHA = 3
COL_HORA = 4
COL_EQUIPO1 = 5
COL_VS = 6
COL_EQUIPO2 = 7
COL_REAL_LOCAL = 8
COL_REAL_VISITANTE = 9
COL_REAL_GANADOR = 10

FILA_INICIO_GRUPOS = 5
FILA_FIN_GRUPOS = 112

FASES_ELIMINATORIA = {
    "DIECISEISAVOS": "Dieciseisavos",
    "OCTAVOS": "Octavos",
    "CUARTOS": "Cuartos",
    "SEMIS": "Semifinales",
    "3ro": "3er y 4to puesto",
    "1ro": "Final",
}

PREMIOS_LABELS = {
    "Goleador del Torneo": "Goleador del Torneo (Bota de Oro)",
    "Mejor Jugador": "Mejor Jugador (Balon de Oro)",
    "Mejor Jugador Joven": "Mejor Jugador Joven",
    "Mejor Arquero": "Mejor Arquero (Guante de Oro)",
    "Equipo Fair Play": "Equipo Fair Play",
    "Equipo más goles anotados": "Equipo que anoto mas goles",
    "Equipo más goles recibidos": "Equipo al que le anotaron mas goles",
    "Equipo sorpresa (menor ranking)": "Equipo sorpresa",
}

RIVAL_MEJOR_TERCERO = {
    "1E": "3ABCDF",
    "1I": "3CDFGH",
    "1A": "3CEFHI",
    "1L": "3EHIJK",
    "1D": "3BEFIJ",
    "1G": "3AEHIJ",
    "1B": "3EFGIJ",
    "1K": "3DEIJL",
}


def slugify(nombre):
    s = nombre.strip().lower()
    s = (s.replace("á", "a").replace("é", "e").replace("í", "i")
           .replace("ó", "o").replace("ú", "u").replace("ñ", "n"))
    s = re.sub(r"[^a-z0-9]+", "-", s)
    return s.strip("-")


def fmt_fecha_hora(fecha, hora):
    out = ""
    if isinstance(fecha, datetime):
        out += fecha.strftime("%d/%m/%Y")
    if isinstance(hora, dtime):
        out += " " + hora.strftime("%H:%M")
    return out.strip() or None


def detectar_participantes(ws):
    """
    Devuelve lista de dicts {nombre, col_inicio} para cada bloque de
    4 columnas a partir de PRIMER_COL_PARTICIPANTE. No asume un numero
    fijo de participantes: recorre hasta que no encuentre mas nombres.
    """
    participantes = []
    c = PRIMER_COL_PARTICIPANTE
    while c <= ws.max_column:
        nombre = ws.cell(row=FILA_NOMBRE_PART, column=c).value
        if not nombre:
            break
        participantes.append({"nombre": str(nombre).strip(), "col_inicio": c})
        c += COLS_POR_BLOQUE
    return participantes


def extraer_partidos_grupo(ws, col_inicio):
    partidos = []
    grupo_actual = None
    r = FILA_INICIO_GRUPOS
    while r <= FILA_FIN_GRUPOS:
        grupo_cell = ws.cell(row=r, column=COL_GRUPO).value
        if grupo_cell:
            grupo_actual = grupo_cell

        num_partido = ws.cell(row=r, column=COL_NUM_PARTIDO).value
        eq1 = ws.cell(row=r, column=COL_EQUIPO1).value

        if num_partido is not None and eq1 is not None:
            eq2 = ws.cell(row=r, column=COL_EQUIPO2).value
            fecha = ws.cell(row=r, column=COL_FECHA).value
            hora = ws.cell(row=r, column=COL_HORA).value
            real_local = ws.cell(row=r, column=COL_REAL_LOCAL).value
            real_visit = ws.cell(row=r, column=COL_REAL_VISITANTE).value

            pred_local = ws.cell(row=r, column=col_inicio).value
            pred_visit = ws.cell(row=r, column=col_inicio + 1).value
            pred_ganador = ws.cell(row=r, column=col_inicio + 2).value
            pred_puntos = ws.cell(row=r, column=col_inicio + 3).value

            jugado = real_local is not None and real_visit is not None

            partidos.append({
                "grupo": grupo_actual,
                "equipo_local": eq1,
                "equipo_visitante": eq2,
                "fecha": fmt_fecha_hora(fecha, hora),
                "jugado": jugado,
                "resultado_real": {"local": real_local, "visitante": real_visit} if jugado else None,
                "prediccion": {"local": pred_local, "visitante": pred_visit, "ganador": pred_ganador},
                "puntos": pred_puntos if jugado else None,
            })
        elif grupo_cell is None:
            texto_clasif = ws.cell(row=r, column=COL_EQUIPO1).value
            if texto_clasif and re.match(r"^\d[A-L]$", str(texto_clasif)):
                pred_ganador = ws.cell(row=r, column=col_inicio + 2).value
                pred_puntos = ws.cell(row=r, column=col_inicio + 3).value
                partidos.append({
                    "tipo": "clasificado_grupo",
                    "posicion": texto_clasif,
                    "prediccion": pred_ganador,
                    "puntos": pred_puntos,
                })
        r += 1
    return partidos


def extraer_eliminatoria(ws, col_inicio):
    rondas = []
    ronda_actual = None
    r = FILA_FIN_GRUPOS + 1
    fin_eliminatoria = None

    for rr in range(r, ws.max_row + 1):
        b = ws.cell(row=rr, column=COL_GRUPO).value
        if b == "Extra":
            fin_eliminatoria = rr
            break
    if fin_eliminatoria is None:
        fin_eliminatoria = ws.max_row

    rr = r
    while rr < fin_eliminatoria:
        fase_cell = ws.cell(row=rr, column=COL_GRUPO).value
        if fase_cell and fase_cell in FASES_ELIMINATORIA:
            ronda_actual = FASES_ELIMINATORIA[fase_cell]

        eq1 = ws.cell(row=rr, column=COL_EQUIPO1).value
        num_partido = ws.cell(row=rr, column=COL_NUM_PARTIDO).value

        if num_partido is not None and ronda_actual is not None:
            eq2 = ws.cell(row=rr, column=COL_EQUIPO2).value
            real_local = ws.cell(row=rr, column=COL_REAL_LOCAL).value
            real_visit = ws.cell(row=rr, column=COL_REAL_VISITANTE).value
            pred_local = ws.cell(row=rr, column=col_inicio).value
            pred_visit = ws.cell(row=rr, column=col_inicio + 1).value
            pred_ganador = ws.cell(row=rr, column=col_inicio + 2).value

            jugado = real_local is not None and real_visit is not None

            eq2_invalido = (eq2 is None) or (isinstance(eq2, str) and eq2.strip().upper() == "#N/A")
            if eq2_invalido and eq1 in RIVAL_MEJOR_TERCERO:
                eq2_mostrar = RIVAL_MEJOR_TERCERO[eq1]
            elif eq2_invalido:
                eq2_mostrar = "Por definir"
            else:
                eq2_mostrar = eq2

            rondas.append({
                "ronda": ronda_actual,
                "equipo_local": eq1 if eq1 else "Por definir",
                "equipo_visitante": eq2_mostrar,
                "jugado": jugado,
                "resultado_real": {"local": real_local, "visitante": real_visit} if jugado else None,
                "prediccion": {"local": pred_local, "visitante": pred_visit, "ganador": pred_ganador},
            })
        rr += 1
    return rondas


def extraer_premios(ws, col_inicio):
    premios = []
    r_start = None
    for r in range(FILA_FIN_GRUPOS, ws.max_row + 1):
        if ws.cell(row=r, column=COL_GRUPO).value == "Extra":
            r_start = r
            break
    if r_start is None:
        return premios

    r = r_start
    while r <= ws.max_row:
        label = ws.cell(row=r, column=COL_EQUIPO1).value
        if not label:
            r += 1
            if r - r_start > 12:
                break
            continue
        label_norm = PREMIOS_LABELS.get(label, label)
        prediccion = ws.cell(row=r, column=col_inicio + 2).value
        premios.append({"premio": label_norm, "prediccion": prediccion})
        r += 1
        if len(premios) >= len(PREMIOS_LABELS):
            break
    return premios


def exportar(path_admin, dir_salida):
    wb = openpyxl.load_workbook(path_admin, data_only=True)
    if SHEET not in wb.sheetnames:
        raise ValueError(f"No se encontro la hoja '{SHEET}' en {path_admin}")
    ws = wb[SHEET]

    os.makedirs(dir_salida, exist_ok=True)

    participantes = detectar_participantes(ws)
    if not participantes:
        raise ValueError("No se encontraron bloques de participantes en la hoja Puntos")

    nombres_usados = {}
    indice = []

    for p in participantes:
        nombre = p["nombre"]
        col_inicio = p["col_inicio"]

        partidos_grupo = extraer_partidos_grupo(ws, col_inicio)
        eliminatoria = extraer_eliminatoria(ws, col_inicio)
        premios = extraer_premios(ws, col_inicio)

        slug = slugify(nombre)
        if slug in nombres_usados:
            nombres_usados[slug] += 1
            slug = f"{slug}-{nombres_usados[slug]}"
        else:
            nombres_usados[slug] = 0

        data = {
            "nombre": nombre,
            "slug": slug,
            "generado": datetime.now().isoformat(timespec="seconds"),
            "fase_grupos": partidos_grupo,
            "eliminatoria": eliminatoria,
            "premios_fifa": premios,
        }

        path_out = os.path.join(dir_salida, f"{slug}.json")
        with open(path_out, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

        indice.append({"nombre": nombre, "slug": slug})

    path_indice = os.path.join(dir_salida, "_indice.json")
    with open(path_indice, "w", encoding="utf-8") as f:
        json.dump({"generado": datetime.now().isoformat(timespec="seconds"),
                    "participantes": indice}, f, ensure_ascii=False, separators=(",", ":"))

    return indice


def main():
    if len(sys.argv) < 2:
        print("Uso: python exportar_predicciones.py <admin.xlsx> [dir_salida]")
        sys.exit(1)

    path_admin = sys.argv[1]
    dir_salida = sys.argv[2] if len(sys.argv) > 2 else "predicciones"

    indice = exportar(path_admin, dir_salida)

    print(f"Exportado: {len(indice)} archivos JSON en {dir_salida}/")
    print(f"Indice generado en {dir_salida}/_indice.json")


if __name__ == "__main__":
    main()

"""
actualizar_web.py

Corre los tres pasos de actualizacion despues de que actualices el Admin
con resultados reales de una nueva jornada:

    1. exportar_ranking.py        -> ranking-animado/data/ranking_historico.json
    2. exportar_predicciones.py   -> ranking-animado/data/predicciones/
    3. generar_web.py             -> index.html (tabla principal, en la raiz)

Para el paso 3, el script pregunta interactivamente el numero de jornada
y la fecha, ya que no todos los dias hay partidos y no es seguro
calcularlo automaticamente. Si prefieres no generar index.html en esta
corrida (por ejemplo, si solo quieres actualizar la grafica y las
predicciones), puedes responder vacio cuando se te pregunte.

Uso normal (usa la ruta por defecto del Admin):
    python actualizar_web.py

Uso indicando una ruta distinta del Admin:
    python actualizar_web.py "C:\\otra\\ruta\\Admin_POLLA_FIFA_WORLD_CUP_2026_r00.xlsx"

No recalcula nada del Admin: solo lee lo que ya esta calculado y lo
convierte a JSON / HTML para la web.
"""

import sys
import os
import subprocess

RUTA_ADMIN_DEFAULT = r"C:\Users\franc\Desktop\Polla Mundial 2026\Admin Original\Admin_POLLA_FIFA_WORLD_CUP_2026_r00.xlsx"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
RAIZ_REPO = os.path.join(SCRIPT_DIR, "..", "..")
DATA_DIR = os.path.join(SCRIPT_DIR, "..", "data")

RUTA_RANKING_JSON = os.path.join(DATA_DIR, "ranking_historico.json")
RUTA_PREDICCIONES_DIR = os.path.join(DATA_DIR, "predicciones")
RUTA_INDICE_PREDICCIONES = os.path.join(RUTA_PREDICCIONES_DIR, "_indice.json")
RUTA_INDEX_HTML = os.path.join(RAIZ_REPO, "index.html")
RUTA_GENERAR_WEB = os.path.join(RAIZ_REPO, "generar_web.py")


def run(cmd):
    print(f"\n>>> {' '.join(cmd)}")
    resultado = subprocess.run(cmd, capture_output=True, text=True)
    if resultado.stdout:
        print(resultado.stdout.strip())
    if resultado.returncode != 0:
        print(f"ERROR (codigo {resultado.returncode}):")
        print(resultado.stderr.strip())
        return False
    return True


def main():
    ruta_admin = sys.argv[1] if len(sys.argv) > 1 else RUTA_ADMIN_DEFAULT

    if not os.path.isfile(ruta_admin):
        print(f"No se encontro el archivo Admin en:\n  {ruta_admin}")
        print("\nVerifica la ruta, o pasa la ruta correcta como argumento:")
        print('  python actualizar_web.py "C:\\ruta\\a\\tu\\Admin.xlsx"')
        sys.exit(1)

    print(f"Usando Admin: {ruta_admin}")

    script_ranking = os.path.join(SCRIPT_DIR, "exportar_ranking.py")
    script_predicciones = os.path.join(SCRIPT_DIR, "exportar_predicciones.py")

    ok_ranking = run([sys.executable, script_ranking, ruta_admin, RUTA_RANKING_JSON])
    ok_predicciones = run([sys.executable, script_predicciones, ruta_admin, RUTA_PREDICCIONES_DIR])

    if not (ok_ranking and ok_predicciones):
        print("\nHubo un error en alguno de los pasos. Revisa los mensajes arriba.")
        sys.exit(1)

    print("\n" + "-" * 50)
    print("Paso 3: actualizar la tabla principal (index.html)")
    print("Dejalo en blanco y presiona Enter si NO quieres regenerar")
    print("index.html en esta corrida.")
    numero_jornada = input("\nNumero de jornada (ej. 7): ").strip()

    ok_generar = True
    if numero_jornada:
        fecha_jornada = input("Fecha (ej. 17 Jun 2026): ").strip()
        label = f"Jornada {numero_jornada} - {fecha_jornada}" if fecha_jornada else f"Jornada {numero_jornada}"

        if not os.path.isfile(RUTA_GENERAR_WEB):
            print(f"\nAviso: no se encontro generar_web.py en {RUTA_GENERAR_WEB}")
            print("Se omite el paso 3. Los pasos 1 y 2 ya se completaron.")
        else:
            ok_generar = run([
                sys.executable, RUTA_GENERAR_WEB,
                "--admin", ruta_admin,
                "--jornada", label,
                "--output", RUTA_INDEX_HTML,
                "--indice-predicciones", RUTA_INDICE_PREDICCIONES,
            ])
    else:
        print("Se omite la actualizacion de index.html.")

    print("\n" + "=" * 50)
    if ok_ranking and ok_predicciones and ok_generar:
        print("LISTO. Archivos actualizados:")
        print(f"  - {os.path.relpath(RUTA_RANKING_JSON, SCRIPT_DIR)}")
        print(f"  - {os.path.relpath(RUTA_PREDICCIONES_DIR, SCRIPT_DIR)}\\*.json")
        if numero_jornada:
            print(f"  - {os.path.relpath(RUTA_INDEX_HTML, SCRIPT_DIR)}")
        print("\nSiguiente paso: revisa los cambios en GitHub Desktop,")
        print("haz commit y push de todo lo actualizado.")
    else:
        print("Hubo un error en alguno de los pasos. Revisa los mensajes arriba.")
        sys.exit(1)


if __name__ == "__main__":
    main()

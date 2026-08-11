"""
Utilidad para etiquetar rápido el ground truth de una hoja real: qué marcó
FÍSICAMENTE el estudiante en cada pregunta (no la pauta académica). Es el
cuello de botella real para todo lo demás en tests/data/omr/ -- sin más fotos
reales etiquetadas no se puede medir generalización, comparar algoritmos ni
evitar el overfitting a la única foto de calibración que había hasta ahora.

Flujo:
  1. Corre el motor OMR sobre la imagen (mismo pipeline que produccion, sin
     llamar a la API) para pre-rellenar una propuesta de respuestas.
  2. Guarda el diagnóstico visual (círculos + etiquetas P/letra/confianza)
     como imagen aparte, para poder mirarlo mientras se corrige.
  3. Muestra P1..Pn en bloques de 20 y pide correcciones puntuales en vez de
     tener que confirmar pregunta por pregunta -- eficiente para etiquetar
     muchas hojas.
  4. Guarda un JSON en el mismo formato que ya usa
     tests/data/omr/hoja_calibracion.ground_truth.json.

Uso:
    py ground_truth_tool.py --imagen ruta/foto.jpg --n 80 [--solo-respuestas]
                             [--salida ruta/foto.ground_truth.json]
                             [--agregar-a-dataset tests/data/omr/dataset.json]

IMPORTANTE (privacidad -- ver item 21 del plan OMR): las fotos reales de
estudiantes y los JSON de ground truth que generes acá son para uso LOCAL.
No los subas al repo (que puede ser público) -- tests/data/omr/ debe seguir
conteniendo solo fixtures ya recortadas/anonimizadas, nunca el dataset real
completo de una aplicación.
"""
import argparse
import json
import os
import sys

import cv2
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import omr_metrics

OPCIONES_VALIDAS = {"A", "B", "C", "D", "E", "-", ""}


def _cargar_respuestas_previas(salida_path: str, n: int):
    """Si ya existe un ground truth para esta imagen, lo usa como punto de partida
    en vez de repartir desde cero -- útil para revisar/corregir una etiquetación
    anterior sin perder lo que ya estaba bien."""
    if not os.path.exists(salida_path):
        return None
    try:
        data = json.load(open(salida_path, encoding="utf-8"))
        return [data["respuestas"].get(str(i)) for i in range(1, n + 1)]
    except Exception:
        return None


def _imprimir_bloque(respuestas, n):
    n_bloques = -(-n // 20)
    for b in range(n_bloques):
        ini = b * 20 + 1
        fin = min((b + 1) * 20, n)
        piezas = [f"P{i}:{respuestas[i - 1] or '-'}" for i in range(ini, fin + 1)]
        print("  " + "  ".join(piezas))


def _aplicar_correcciones(respuestas, texto):
    """Parsea correcciones tipo '12=C 45=- 67=B' (case-insensitive, '-' = en blanco).
    Devuelve (respuestas_corregidas, errores) -- errores son entradas que no se
    pudieron interpretar, para que la persona las vea y las reintente."""
    errores = []
    for token in texto.split():
        if "=" not in token:
            errores.append(token)
            continue
        num_s, letra = token.split("=", 1)
        letra = letra.strip().upper()
        try:
            num = int(num_s)
        except ValueError:
            errores.append(token)
            continue
        if not (1 <= num <= len(respuestas)) or letra not in OPCIONES_VALIDAS:
            errores.append(token)
            continue
        respuestas[num - 1] = None if letra in ("-", "") else letra
    return respuestas, errores


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--imagen", required=True, help="ruta a la foto real de la hoja")
    ap.add_argument("--n", type=int, default=80, help="número de preguntas de la plantilla (default 80)")
    ap.add_argument("--solo-respuestas", action="store_true",
                     help="la imagen ya viene recortada mostrando solo el bloque RESPUESTAS")
    ap.add_argument("--salida", default=None,
                     help="ruta del .ground_truth.json a escribir (default: <imagen>.ground_truth.json)")
    ap.add_argument("--agregar-a-dataset", default=None,
                     help="si se indica, agrega/actualiza la entrada de esta imagen en ese dataset.json")
    args = ap.parse_args()

    salida_path = args.salida or (os.path.splitext(args.imagen)[0] + ".ground_truth.json")
    diag_path = os.path.splitext(salida_path)[0] + ".diagnostico.jpg"

    app = omr_metrics._cargar_app_module()
    datos_bytes = open(args.imagen, "rb").read()

    propuesta = [None] * args.n
    try:
        img_bgr = cv2.imdecode(np.frombuffer(datos_bytes, np.uint8), cv2.IMREAD_COLOR)
        salida = app.omr_analizar_imagen(img_bgr, es_recorte=args.solo_respuestas, n_preguntas=args.n)
        resultados = (salida["resultados"] + [{"letra": None}] * args.n)[:args.n]
        propuesta = [r["letra"] for r in resultados]
        diag = app.omr_anotar_diagnostico(salida["body_bgr"], salida["y_centers"], salida["band_x_centers"],
                                           salida["radio"], resultados)
        ok, buf = cv2.imencode(".jpg", diag, [cv2.IMWRITE_JPEG_QUALITY, 90])
        if ok:
            open(diag_path, "wb").write(buf.tobytes())
            print(f"Diagnóstico visual guardado en: {diag_path}  (ábrelo para comparar contra la foto)")
    except Exception as e:
        print(f"AVISO: el motor OMR no pudo leer esta hoja ({e}). Partiendo de todo en blanco -- "
              "corrige a mano desde cero.", file=sys.stderr)

    previas = _cargar_respuestas_previas(salida_path, args.n)
    respuestas = previas if previas is not None else propuesta
    if previas is not None:
        print(f"Ya existía un ground truth en {salida_path} -- partiendo de esa etiquetación previa.")

    print(f"\nPropuesta del motor OMR para {args.n} preguntas (revisa contra la foto/diagnóstico):")
    _imprimir_bloque(respuestas, args.n)

    while True:
        print("\nEscribe correcciones como 'N=LETRA' separadas por espacio (LETRA en A-E, o '-' para en blanco).")
        print("Ejemplos: '34=D 66=-'   |   enter vacío = aceptar todo tal cual   |   'q' = salir sin guardar")
        texto = input("> ").strip()
        if texto.lower() == "q":
            print("Cancelado, no se guardó nada.")
            return
        if not texto:
            break
        respuestas, errores = _aplicar_correcciones(respuestas, texto)
        if errores:
            print(f"No se pudieron interpretar: {errores} -- revísalas e intenta de nuevo.")
        _imprimir_bloque(respuestas, args.n)

    gt = {
        "imagen": os.path.basename(args.imagen),
        "modo": "solo_respuestas" if args.solo_respuestas else "completa",
        "n_preguntas": args.n,
        "nota": "Ground truth = lo que el estudiante marco fisicamente en la hoja (verificado a mano), "
                "no la respuesta academicamente correcta. Preguntas en null quedaron genuinamente en "
                "blanco o ilegibles en la foto original.",
        "respuestas": {str(i): respuestas[i - 1] for i in range(1, args.n + 1)},
    }
    json.dump(gt, open(salida_path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"\nGuardado: {salida_path}")

    if args.agregar_a_dataset:
        _agregar_a_dataset(args.agregar_a_dataset, args.imagen, args.n, args.solo_respuestas, respuestas)


def _agregar_a_dataset(dataset_path: str, imagen_path: str, n: int, solo_respuestas: bool, respuestas: list):
    casos = json.load(open(dataset_path, encoding="utf-8")) if os.path.exists(dataset_path) else []
    entrada = {
        "imagen": imagen_path.replace("\\", "/"),
        "n_preguntas": n,
        "solo_respuestas": solo_respuestas,
        "respuestas_correctas": respuestas,
    }
    casos = [c for c in casos if c.get("imagen") != entrada["imagen"]] + [entrada]
    json.dump(casos, open(dataset_path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"Actualizado {dataset_path} ({len(casos)} imagen(es) en total).")


if __name__ == "__main__":
    main()

"""
Harness de métricas para comparar 3 métodos de lectura de hojas de respuestas:
  A) "ia"      — flujo 100% Claude Vision, Claude lee las burbujas (app_revisor.procesar_imagen)
  B) "omr"     — motor OMR puro, sin ninguna llamada a la API (omr_analizar_imagen)
  C) "hibrido" — respuestas 100% del motor OMR; Claude solo transcribe nombre/RUT de la
                 cabecera (nunca lee burbujas) (app_revisor.procesar_imagen_hibrido)

El motor OMR vive inline dentro de app_revisor.py (no en un módulo aparte) a
propósito, para eliminar la posibilidad de que ese archivo y un omr.py separado
queden desincronizados en un redeploy. Este script extrae las funciones que
necesita directamente del código fuente de app_revisor.py, sin ejecutar su UI.

Uso (requiere una API key de Anthropic solo para los métodos "ia" e "hibrido";
"omr" no llama a la API y se puede correr sin key):

    py omr_metrics.py --dataset dataset.json [--api-key sk-ant-...] [--metodos omr,hibrido,ia]

dataset.json tiene esta forma:
[
  {"imagen": "ruta/a/foto1.jpg", "n_preguntas": 80, "solo_respuestas": false,
   "respuestas_correctas": ["C","D","A", ..., null, null]},
  ...
]

"respuestas_correctas" es la respuesta REAL que marcó el estudiante en cada
pregunta (ground truth transcrita a mano por una persona mirando la foto), NO
la pauta de corrección del examen — este harness mide qué tan bien cada método
LEE la hoja, no si el alumno acertó la pregunta.

Para cada método calcula: accuracy total, accuracy por pregunta (posición en la
hoja), accuracy por alternativa (A/B/C/D/E), falsos positivos/negativos sobre
"sin respuesta", % resuelto sin IA, cantidad de llamadas a Claude, tiempo
promedio por imagen, y una matriz de confusión A-E (más "—" = sin respuesta).
"""

import argparse
import json
import sys
import time
from collections import defaultdict

import cv2
import numpy as np

LETRAS_CM = ["A", "B", "C", "D", "E", "—"]  # "—" representa None/sin respuesta


def matriz_confusion_vacia():
    return {real: {pred: 0 for pred in LETRAS_CM} for real in LETRAS_CM}


def actualizar_matriz(cm, real, pred):
    r = real if real in LETRAS_CM[:5] else "—"
    p = pred if pred in LETRAS_CM[:5] else "—"
    cm[r][p] += 1


def imprimir_matriz(cm):
    header = "      " + "  ".join(f"{l:>4}" for l in LETRAS_CM)
    print(header)
    for real in LETRAS_CM:
        fila = "  ".join(f"{cm[real][pred]:>4}" for pred in LETRAS_CM)
        print(f"{real:>4}  {fila}")


def evaluar_respuestas(pred: list, real: list, estados: list = None, etiquetas_confiables: set = None) -> dict:
    """Si se pasan `estados` (el status/metodo por pregunta que devolvió el motor) y
    `etiquetas_confiables` (qué estados cuentan como "el motor dice que esto es
    confiable"), además calcula high-confidence precision/coverage: de las
    preguntas que el motor entregó como confiables, qué fracción realmente
    coincide con el ground truth. Esta es la métrica más importante del sistema
    -- una respuesta "confiable" incorrecta es la falla más peligrosa posible,
    mucho peor que dejar una pregunta dudosa para revisión manual."""
    n = len(real)
    aciertos = 0
    fp_sin_respuesta = 0  # dijo letra, era None
    fn_sin_respuesta = 0  # dijo None, había letra
    cm = matriz_confusion_vacia()
    por_pregunta = []
    n_alta_confianza = 0
    aciertos_alta_confianza = 0
    for i in range(n):
        p, r = pred[i] if i < len(pred) else None, real[i]
        ok = (p == r)
        aciertos += ok
        por_pregunta.append(ok)
        actualizar_matriz(cm, r, p)
        if r is None and p is not None:
            fp_sin_respuesta += 1
        if r is not None and p is None:
            fn_sin_respuesta += 1
        if estados and etiquetas_confiables and i < len(estados) and estados[i] in etiquetas_confiables:
            n_alta_confianza += 1
            aciertos_alta_confianza += ok
    return {
        "accuracy": aciertos / n if n else 0.0,
        "aciertos": aciertos, "n": n,
        "falsos_positivos_sin_respuesta": fp_sin_respuesta,
        "falsos_negativos_sin_respuesta": fn_sin_respuesta,
        "confusion": cm,
        "por_pregunta": por_pregunta,
        "n_alta_confianza": n_alta_confianza,
        "aciertos_alta_confianza": aciertos_alta_confianza,
    }


ESTADOS_CONFIABLES_OMR = {"alta_confianza"}
ESTADOS_CONFIABLES_HIBRIDO = {"confiable"}


def correr_metodo_omr(app_module, datos_bytes: bytes, solo_respuestas: bool, n: int) -> dict:
    t0 = time.time()
    img_bgr = cv2.imdecode(np.frombuffer(datos_bytes, np.uint8), cv2.IMREAD_COLOR)
    salida = app_module.omr_analizar_imagen(img_bgr, es_recorte=solo_respuestas, n_preguntas=n)
    resultados = (salida["resultados"] + [{"letra": None, "status": "sin_marca", "omr_confidence": 0.0}] * n)[:n]
    respuestas = [r["letra"] for r in resultados]
    estados = [r["status"] for r in resultados]
    elapsed = time.time() - t0
    n_directo = sum(1 for r in resultados if r["status"] in ("alta_confianza", "confianza_media"))
    n_geo_error = sum(1 for r in resultados if r["status"] == "geometry_error")
    return {"respuestas": respuestas, "estados": estados, "etiquetas_confiables": ESTADOS_CONFIABLES_OMR,
            "tiempo_s": elapsed, "llamadas_api": 0,
            "pct_resuelto_sin_ia": n_directo / n if n else 0.0,
            "geometry_confidence_por_banda": salida.get("geometry_confidence_por_banda", []),
            "n_geometry_error": n_geo_error}


def correr_metodo_ia(app_module, cliente, datos_bytes: bytes, solo_respuestas: bool, n: int) -> dict:
    t0 = time.time()
    res = app_module.procesar_imagen(cliente, "eval", datos_bytes, "image/jpeg", n, solo_respuestas)
    elapsed = time.time() - t0
    return {"respuestas": res["respuestas"], "estados": [], "etiquetas_confiables": set(),
            "tiempo_s": elapsed, "llamadas_api": res.get("intentos", 1), "pct_resuelto_sin_ia": 0.0,
            "geometry_confidence_por_banda": [], "n_geometry_error": 0}


def correr_metodo_hibrido(app_module, cliente, datos_bytes: bytes, solo_respuestas: bool, n: int) -> dict:
    """Las respuestas son 100% del motor OMR; la única llamada a la API (si la hay) es
    para transcribir nombre/RUT de la cabecera -- nunca para leer burbujas."""
    t0 = time.time()
    res = app_module.procesar_imagen_hibrido(cliente, "eval", datos_bytes, "image/jpeg", n, solo_respuestas)
    elapsed = time.time() - t0
    meta = res.get("omr_meta", {})
    n_llamadas = 0 if (res.get("solo_respuestas") or not meta.get("usado")) else 1  # solo identificación
    return {"respuestas": res["respuestas"], "estados": meta.get("metodo_por_pregunta", []),
            "etiquetas_confiables": ESTADOS_CONFIABLES_HIBRIDO,
            "tiempo_s": elapsed, "llamadas_api": n_llamadas,
            "pct_resuelto_sin_ia": meta.get("n_confiable", 0) / n if n and meta.get("usado") else 0.0,
            "geometry_confidence_por_banda": meta.get("geometry_confidence_por_banda", []),
            "n_geometry_error": meta.get("n_geometry_error", 0)}


def _cargar_app_module():
    """Carga app_revisor.py como módulo de funciones puras, sin ejecutar su UI de Streamlit,
    igual que se hizo durante el desarrollo/pruebas de este pipeline. El bloque completo del
    motor OMR (entre los comentarios "MOTOR OMR" y "fin motor OMR") se extrae de una sola vez;
    el resto de funciones se extrae una por una como antes."""
    import re as _re
    import types
    ruta = __file__.replace("omr_metrics.py", "app_revisor.py")
    src = open(ruta, encoding="utf-8").read()

    def extraer(nombre):
        start = src.index(f"\ndef {nombre}") + 1
        rest = src[start + len(f"def {nombre}"):]
        m = _re.search(r'\ndef |\n\n\n#', rest)
        end = start + len(f"def {nombre}") + (m.start() if m else len(rest))
        return src[start:end]

    bloque_omr = src[src.index("OMR_THRESHOLDS = {"):src.index("# ═══════════════════════ fin motor OMR")]

    nombres = ["prompt_dinamico", "prompt_identificacion", "prompt_revision_dudosas",
               "abrir_imagen_corregida", "preparar_imagenes", "mejorar_contraste_burbujas",
               "_img_a_b64_jpeg", "evaluar_sospecha", "_llamar_claude", "procesar_imagen",
               "_bgr_a_jpeg_b64", "llamar_claude_identificacion", "llamar_claude_revision_dudosas",
               "_crop_dudosa_b64_jpeg",
               "analizar_hoja_omr", "_fallback_no_leido", "_construir_resultado_omr",
               "procesar_imagen_hibrido"]
    import io, base64, hashlib, json as _json
    from collections import Counter
    from PIL import Image, ImageEnhance, ImageOps
    import anthropic
    mod = types.ModuleType("app_revisor_funcs")
    ns = mod.__dict__
    ns.update({
        "io": io, "json": _json, "re": _re, "base64": base64, "hashlib": hashlib, "Counter": Counter,
        "np": np, "cv2": cv2, "Image": Image, "ImageEnhance": ImageEnhance, "ImageOps": ImageOps,
        "anthropic": anthropic, "LETRAS_VALIDAS": {"A", "B", "C", "D", "E"},
        "REFUERZO_REINTENTO": _re.search(r'REFUERZO_REINTENTO = \(([\s\S]*?)\)\n', src).group(0).split("=", 1)[1].strip().rstrip(")").strip("(").strip(),
    })
    exec(bloque_omr, ns)
    for n in nombres:
        exec(extraer(n), ns)
    return mod


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dataset", required=True, help="ruta a dataset.json")
    ap.add_argument("--api-key", default=None, help="API key de Anthropic (solo necesaria para omr/hibrido)")
    ap.add_argument("--metodos", default="omr", help="coma-separado: omr,ia,hibrido (default: solo omr, no requiere API key)")
    args = ap.parse_args()

    metodos = [m.strip() for m in args.metodos.split(",") if m.strip()]
    casos = json.load(open(args.dataset, encoding="utf-8"))

    cliente = None
    if "ia" in metodos or "hibrido" in metodos:
        if not args.api_key:
            print("Error: --api-key es obligatorio para los métodos 'ia' o 'hibrido'.", file=sys.stderr)
            sys.exit(1)
        import anthropic
        cliente = anthropic.Anthropic(api_key=args.api_key, timeout=90.0, max_retries=1)
    app_module = _cargar_app_module()  # siempre hace falta: "omr" también vive en app_revisor.py

    resultados_por_metodo = defaultdict(list)
    tiempos = defaultdict(list)
    llamadas = defaultdict(list)
    pct_sin_ia = defaultdict(list)

    for caso in casos:
        datos_bytes = open(caso["imagen"], "rb").read()
        n = caso["n_preguntas"]
        solo_resp = caso.get("solo_respuestas", False)
        real = caso["respuestas_correctas"]

        for metodo in metodos:
            try:
                if metodo == "omr":
                    salida = correr_metodo_omr(app_module, datos_bytes, solo_resp, n)
                elif metodo == "ia":
                    salida = correr_metodo_ia(app_module, cliente, datos_bytes, solo_resp, n)
                elif metodo == "hibrido":
                    salida = correr_metodo_hibrido(app_module, cliente, datos_bytes, solo_resp, n)
                else:
                    continue
            except Exception as e:
                print(f"[{metodo}] {caso['imagen']}: ERROR {e}", file=sys.stderr)
                continue
            ev = evaluar_respuestas(salida["respuestas"], real, salida.get("estados"), salida.get("etiquetas_confiables"))
            ev["n_geometry_error"] = salida.get("n_geometry_error", 0)
            resultados_por_metodo[metodo].append(ev)
            tiempos[metodo].append(salida["tiempo_s"])
            llamadas[metodo].append(salida["llamadas_api"])
            pct_sin_ia[metodo].append(salida["pct_resuelto_sin_ia"])
            geo_txt = ""
            if salida.get("geometry_confidence_por_banda"):
                geo_txt = f"  geometry_confidence={[round(g,2) for g in salida['geometry_confidence_por_banda']]}"
            print(f"[{metodo}] {caso['imagen']}: accuracy={ev['accuracy']*100:.1f}% "
                  f"({ev['aciertos']}/{ev['n']})  tiempo={salida['tiempo_s']*1000:.0f}ms  "
                  f"llamadas_api={salida['llamadas_api']}{geo_txt}")

    print("\n" + "=" * 70)
    for metodo in metodos:
        evs = resultados_por_metodo[metodo]
        if not evs:
            continue
        acc_total = sum(e["aciertos"] for e in evs) / sum(e["n"] for e in evs)
        cm = matriz_confusion_vacia()
        for e in evs:
            for real_l in LETRAS_CM:
                for pred_l in LETRAS_CM:
                    cm[real_l][pred_l] += e["confusion"][real_l][pred_l]
        print(f"\n### Método: {metodo}  ({len(evs)} imagen(es))")
        print(f"Accuracy total: {acc_total*100:.2f}%")
        n_ac = sum(e["n_alta_confianza"] for e in evs)
        aciertos_ac = sum(e["aciertos_alta_confianza"] for e in evs)
        n_total = sum(e["n"] for e in evs)
        if n_ac:
            precision_ac = aciertos_ac / n_ac
            print(f"HIGH-CONFIDENCE precision: {precision_ac*100:.2f}%  ({aciertos_ac}/{n_ac})  "
                  f"<- métrica principal: una respuesta 'confiable' incorrecta es la falla más grave")
            print(f"HIGH-CONFIDENCE coverage: {n_ac/n_total*100:.1f}%  ({n_ac}/{n_total} preguntas)")
        n_geo_error = sum(e.get("n_geometry_error", 0) for e in evs)
        if n_geo_error:
            print(f"GEOMETRY_ERROR (sin evidencia de grilla real, forzadas a revisión manual): "
                  f"{n_geo_error}/{n_total}")
        print(f"Tiempo promedio por imagen: {sum(tiempos[metodo])/len(tiempos[metodo])*1000:.0f} ms")
        print(f"Llamadas a Claude promedio por imagen: {sum(llamadas[metodo])/len(llamadas[metodo]):.2f}")
        print(f"% resuelto sin llamar a IA: {sum(pct_sin_ia[metodo])/len(pct_sin_ia[metodo])*100:.1f}%")
        print(f"Falsos positivos 'sin respuesta' (dijo letra, era blanco): "
              f"{sum(e['falsos_positivos_sin_respuesta'] for e in evs)}")
        print(f"Falsos negativos 'sin respuesta' (dijo blanco, había letra): "
              f"{sum(e['falsos_negativos_sin_respuesta'] for e in evs)}")
        print("Matriz de confusión (fila=real, columna=predicho):")
        imprimir_matriz(cm)


if __name__ == "__main__":
    main()

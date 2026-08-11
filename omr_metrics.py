"""
Harness de métricas para comparar 3 métodos de lectura de hojas de respuestas.
Terminología exacta (no intercambiable -- ver item 6 del cierre de esta rama):

  A) "ia"      — flujo 100% Claude Vision: Claude LEE las burbujas directamente,
                 sin ningún paso de visión clásica (app_revisor.procesar_imagen).
  B) "omr"     — motor OMR puro: SOLO Computer Vision (OpenCV/Hough/kmeans),
                 cero llamadas a la API, cero intervención de Claude en nada,
                 ni siquiera identificación (omr_analizar_imagen).
  C) "hibrido" — OMR es la fuente PRIMARIA de A-E, siempre. Claude interviene
                 en dos frentes acotados y nunca reemplaza una respuesta que
                 el OMR ya dio como confiable:
                   1. identificación (nombre/RUT/folleto) -- siempre disponible;
                   2. arbitraje de preguntas ambiguas/doble marca -- SOLO si
                      `ia_arbitraje_habilitado=True` (default False) y solo
                      sobre bandas con geometry_state==GEOMETRY_OK.
                 Una respuesta resuelta por el arbitraje de IA NUNCA es una
                 "respuesta OMR" -- se cuenta aparte (n_answers_ai vs.
                 n_answers_omr, ver evaluar_respuestas/correr_metodo_hibrido).
                 (app_revisor.procesar_imagen_hibrido)

El motor OMR vive inline dentro de app_revisor.py (no en un módulo aparte) a
propósito, para eliminar la posibilidad de que ese archivo y un omr.py separado
queden desincronizados en un redeploy. Este script extrae las funciones que
necesita directamente del código fuente de app_revisor.py, sin ejecutar su UI.

Uso (requiere una API key de Anthropic solo para los métodos "ia" e "hibrido";
"omr" no llama a la API y se puede correr sin key):

    py omr_metrics.py --dataset dataset.json [--api-key sk-ant-...] \
        [--metodos omr,hibrido,ia] [--ia-arbitraje]

dataset.json tiene esta forma:
[
  {"imagen": "ruta/a/foto1.jpg", "n_preguntas": 80, "solo_respuestas": false,
   "respuestas_correctas": ["C","D","A", ..., null, null]},
  ...
]

El campo se acepta también como "ground_truth" (alias compatible, mismo
formato) -- es la respuesta REAL que marcó el estudiante en cada pregunta
(ground truth transcrita a mano por una persona mirando la foto). NO ES LA
PAUTA ACADÉMICA DEL EXAMEN: este harness mide qué tan bien cada método LEE la
hoja (qué burbuja está rellena), no si el alumno acertó la pregunta.

Para cada método calcula: accuracy total, accuracy por pregunta (posición en la
hoja), accuracy por alternativa (A/B/C/D/E), falsos positivos/negativos sobre
"sin respuesta", % resuelto sin IA, cantidad de llamadas a Claude (desglosadas
por identificación vs. arbitraje), tiempo promedio por imagen, y una matriz de
confusión A-E (más "—" = sin respuesta).
"""

import argparse
import json
import sys
import time
from collections import defaultdict, Counter

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


def evaluar_respuestas(pred: list, real: list, estados: list = None,
                        etiquetas_confiables: set = None, etiquetas_media: set = None) -> dict:
    """Si se pasan `estados` (el status/metodo por pregunta que devolvió el motor),
    `etiquetas_confiables` (qué estados cuentan como "el motor dice que esto es
    confiable") y `etiquetas_media` (estados de confianza media), además calcula:

    - HIGH-CONFIDENCE precision/coverage: de las preguntas que el motor entregó
      como confiables, qué fracción realmente coincide con el ground truth --
      la métrica más importante del sistema, una respuesta "confiable"
      incorrecta es la falla más peligrosa posible, mucho peor que dejar una
      pregunta dudosa para revisión manual.
    - MEDIUM-CONFIDENCE precision: lo mismo pero sobre las de confianza media.
    - blank precision/recall: de las preguntas que el motor dice "en blanco",
      cuántas realmente lo estaban (precision), y de las que realmente estaban
      en blanco, cuántas el motor efectivamente marcó como tal (recall).
    - un Counter() de los estados crudos, para que quien llama derive tasas
      específicas del método (ambiguo, doble_marca, geometry_error, etc.) sin
      que esta función tenga que conocer el vocabulario de cada método."""
    n = len(real)
    aciertos = 0
    fp_sin_respuesta = 0  # dijo letra, era None
    fn_sin_respuesta = 0  # dijo None, había letra
    cm = matriz_confusion_vacia()
    por_pregunta = []
    n_alta_confianza = aciertos_alta_confianza = 0
    n_media = aciertos_media = 0
    pred_blank = real_blank = blank_correcto = 0
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
        if p is None:
            pred_blank += 1
        if r is None:
            real_blank += 1
        if p is None and r is None:
            blank_correcto += 1
        if estados and i < len(estados):
            if etiquetas_confiables and estados[i] in etiquetas_confiables:
                n_alta_confianza += 1
                aciertos_alta_confianza += ok
            if etiquetas_media and estados[i] in etiquetas_media:
                n_media += 1
                aciertos_media += ok
    return {
        "accuracy": aciertos / n if n else 0.0,
        "aciertos": aciertos, "n": n,
        "falsos_positivos_sin_respuesta": fp_sin_respuesta,
        "falsos_negativos_sin_respuesta": fn_sin_respuesta,
        "confusion": cm,
        "por_pregunta": por_pregunta,
        "n_alta_confianza": n_alta_confianza,
        "aciertos_alta_confianza": aciertos_alta_confianza,
        "n_media": n_media,
        "aciertos_media": aciertos_media,
        "pred_blank": pred_blank, "real_blank": real_blank, "blank_correcto": blank_correcto,
        "estados_counter": Counter(estados) if estados else Counter(),
    }


ESTADOS_CONFIABLES_OMR = {"alta_confianza"}
ESTADOS_MEDIA_OMR = {"confianza_media"}
ESTADOS_AMBIGUOS_OMR = {"ambiguo"}
ESTADOS_DOBLE_MARCA_OMR = {"doble_marca"}
ESTADOS_GEOMETRY_ERROR_OMR = {"geometry_error"}

ESTADOS_CONFIABLES_HIBRIDO = {"confiable"}
ESTADOS_MEDIA_HIBRIDO = {"revisar_media"}
ESTADOS_MANUAL_REVIEW_HIBRIDO = {"revisar_dudoso", "revisar_geometria"}
ESTADOS_AI_HIBRIDO = {"revisada_ia"}
ESTADOS_GEOMETRY_ERROR_HIBRIDO = {"revisar_geometria"}

# Qué tasas derivar del Counter de estados crudos de cada método -- el
# vocabulario de estados NO es el mismo entre "omr" puro (alta_confianza,
# confianza_media, ambiguo, doble_marca, geometry_error, sin_marca) e
# "hibrido" (confiable, revisar_media, revisar_dudoso, revisar_geometria,
# revisada_ia, blanco), así que cada método declara aparte qué estados
# cuentan para cada tasa reportada.
METODO_TASAS = {
    "omr": {
        "ambiguous_rate": ESTADOS_AMBIGUOS_OMR,
        "double_mark_rate": ESTADOS_DOBLE_MARCA_OMR,
        "geometry_error_rate": ESTADOS_GEOMETRY_ERROR_OMR,
    },
    "hibrido": {
        "manual_review_rate": ESTADOS_MANUAL_REVIEW_HIBRIDO,
        "AI_arbitration_rate": ESTADOS_AI_HIBRIDO,
        "geometry_error_rate": ESTADOS_GEOMETRY_ERROR_HIBRIDO,
    },
}


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
    return {"respuestas": respuestas, "estados": estados,
            "etiquetas_confiables": ESTADOS_CONFIABLES_OMR, "etiquetas_media": ESTADOS_MEDIA_OMR,
            "tiempo_s": elapsed, "llamadas_api": 0, "n_api_calls_identification": 0, "n_api_calls_answer_arbitration": 0,
            "pct_resuelto_sin_ia": n_directo / n if n else 0.0,
            "geometry_confidence_por_banda": salida.get("geometry_confidence_por_banda", []),
            "row_alignment_confidence_por_banda": salida.get("row_alignment_confidence_por_banda", []),
            "n_geometry_error": n_geo_error}


def correr_metodo_ia(app_module, cliente, datos_bytes: bytes, solo_respuestas: bool, n: int) -> dict:
    t0 = time.time()
    res = app_module.procesar_imagen(cliente, "eval", datos_bytes, "image/jpeg", n, solo_respuestas)
    elapsed = time.time() - t0
    return {"respuestas": res["respuestas"], "estados": [], "etiquetas_confiables": set(), "etiquetas_media": set(),
            "tiempo_s": elapsed, "llamadas_api": res.get("intentos", 1),
            "n_api_calls_identification": res.get("intentos", 1), "n_api_calls_answer_arbitration": 0,
            "pct_resuelto_sin_ia": 0.0, "geometry_confidence_por_banda": [], "n_geometry_error": 0}


def correr_metodo_hibrido(app_module, cliente, datos_bytes: bytes, solo_respuestas: bool, n: int,
                           ia_arbitraje_habilitado: bool = False) -> dict:
    """OMR es la fuente PRIMARIA de A-E, siempre. Hay DOS llamadas a la API
    posibles, ninguna para leer una burbuja directamente: identificación de
    cabecera (nombre/RUT, siempre disponible) y arbitraje de las preguntas que
    el OMR deja genuinamente ambiguas (solo si `ia_arbitraje_habilitado=True`,
    default False -- igual que en producción). Antes este harness asumía que
    la única llamada posible era la de identificación (n_llamadas = 0 o 1), lo
    que subestimaba el costo real de una hoja con ambiguas -- bug real,
    corregido leyendo los dos conteos que ahora expone omr_meta por separado.
    Una respuesta arbitrada por IA se cuenta en n_answers_ai, NUNCA en
    n_answers_omr -- no es lo mismo, aunque ambas terminen en el mismo array
    `respuestas` que devuelve el híbrido."""
    t0 = time.time()
    res = app_module.procesar_imagen_hibrido(cliente, "eval", datos_bytes, "image/jpeg", n, solo_respuestas,
                                               ia_arbitraje_habilitado=ia_arbitraje_habilitado)
    elapsed = time.time() - t0
    meta = res.get("omr_meta", {})
    n_id = meta.get("n_api_calls_identification", 0)
    n_arb = meta.get("n_api_calls_answer_arbitration", 0)
    return {"respuestas": res["respuestas"], "estados": meta.get("metodo_por_pregunta", []),
            "etiquetas_confiables": ESTADOS_CONFIABLES_HIBRIDO, "etiquetas_media": ESTADOS_MEDIA_HIBRIDO,
            "tiempo_s": elapsed, "llamadas_api": n_id + n_arb,
            "n_api_calls_identification": n_id, "n_api_calls_answer_arbitration": n_arb,
            "n_answers_omr": meta.get("n_answers_omr", 0), "n_answers_ai": meta.get("n_answers_ai", 0),
            "n_answers_manual": meta.get("n_answers_manual", 0),
            "n_answers_unresolved": meta.get("n_answers_unresolved", 0),
            "pct_resuelto_sin_ia": meta.get("n_confiable", 0) / n if n and meta.get("usado") else 0.0,
            "geometry_confidence_por_banda": meta.get("geometry_confidence_por_banda", []),
            "row_alignment_confidence_por_banda": meta.get("row_alignment_confidence_por_banda", []),
            "n_geometry_error": meta.get("n_geometry_error", 0),
            "n_geometry_warning": meta.get("n_geometry_warning", 0)}


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
    ap.add_argument("--ia-arbitraje", action="store_true",
                     help="habilita el arbitraje de IA para ambiguas en 'hibrido' (default: apagado, igual que en producción)")
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
        # "ground_truth" es el nombre conceptualmente correcto (alias
        # compatible): lo que el estudiante marcó físicamente, NO la pauta
        # académica. Se acepta también el nombre histórico del campo para no
        # romper los fixtures ya existentes en tests/data/omr/.
        real = caso.get("ground_truth", caso.get("respuestas_correctas"))
        if real is None:
            print(f"[{caso['imagen']}] sin 'ground_truth' ni 'respuestas_correctas' en el dataset -- se omite.",
                  file=sys.stderr)
            continue

        for metodo in metodos:
            try:
                if metodo == "omr":
                    salida = correr_metodo_omr(app_module, datos_bytes, solo_resp, n)
                elif metodo == "ia":
                    salida = correr_metodo_ia(app_module, cliente, datos_bytes, solo_resp, n)
                elif metodo == "hibrido":
                    salida = correr_metodo_hibrido(app_module, cliente, datos_bytes, solo_resp, n,
                                                    ia_arbitraje_habilitado=args.ia_arbitraje)
                else:
                    continue
            except Exception as e:
                print(f"[{metodo}] {caso['imagen']}: ERROR {e}", file=sys.stderr)
                continue
            ev = evaluar_respuestas(salida["respuestas"], real, salida.get("estados"),
                                     salida.get("etiquetas_confiables"), salida.get("etiquetas_media"))
            ev["n_geometry_error"] = salida.get("n_geometry_error", 0)
            ev["n_geometry_warning"] = salida.get("n_geometry_warning", 0)
            ev["n_api_calls_identification"] = salida.get("n_api_calls_identification", 0)
            ev["n_api_calls_answer_arbitration"] = salida.get("n_api_calls_answer_arbitration", 0)
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
        n_med = sum(e["n_media"] for e in evs)
        aciertos_med = sum(e["aciertos_media"] for e in evs)
        if n_med:
            precision_med = aciertos_med / n_med
            print(f"MEDIUM-CONFIDENCE precision: {precision_med*100:.2f}%  ({aciertos_med}/{n_med})")

        n_geo_error = sum(e.get("n_geometry_error", 0) for e in evs)
        n_geo_warning = sum(e.get("n_geometry_warning", 0) for e in evs)
        # geometry_success_rate: fracción de preguntas cuya geometría fue
        # demostrable (no forzada a revisión manual por falta de evidencia) --
        # separado de accuracy: mide si el motor pudo UBICAR la grilla, no si
        # leyó bien la marca una vez ubicada.
        geometry_success_rate = 1 - (n_geo_error / n_total) if n_total else 0.0
        print(f"geometry_success_rate: {geometry_success_rate*100:.2f}%  "
              f"(GEOMETRY_ERROR: {n_geo_error}/{n_total}, GEOMETRY_WARNING: {n_geo_warning}/{n_total})")

        # blank precision/recall: de lo que el método dijo "en blanco", cuánto
        # realmente lo estaba (precision); de lo que realmente estaba en
        # blanco, cuánto se detectó como tal (recall).
        pred_blank = sum(e["pred_blank"] for e in evs)
        real_blank = sum(e["real_blank"] for e in evs)
        blank_correcto = sum(e["blank_correcto"] for e in evs)
        if pred_blank:
            print(f"blank_precision: {blank_correcto/pred_blank*100:.2f}%  ({blank_correcto}/{pred_blank})")
        if real_blank:
            print(f"blank_recall: {blank_correcto/real_blank*100:.2f}%  ({blank_correcto}/{real_blank})")

        # Tasas derivadas del Counter de estados crudos -- vocabulario propio
        # de cada método (ver METODO_TASAS).
        estados_totales = Counter()
        for e in evs:
            estados_totales.update(e["estados_counter"])
        for nombre_tasa, etiquetas in METODO_TASAS.get(metodo, {}).items():
            count = sum(estados_totales.get(e, 0) for e in etiquetas)
            if count:
                print(f"{nombre_tasa}: {count/n_total*100:.2f}%  ({count}/{n_total})")

        # accuracy_by_letter: de las preguntas donde el ground truth es esa
        # letra, qué fracción el método acertó -- deja ver si el error se
        # concentra en una alternativa puntual (p.ej. una columna corrida).
        partes_letra = []
        for letra in LETRAS_CM[:5]:
            total_letra = sum(cm[letra][p] for p in LETRAS_CM)
            if total_letra:
                partes_letra.append(f"{letra}={cm[letra][letra]/total_letra*100:.0f}%({cm[letra][letra]}/{total_letra})")
        if partes_letra:
            print("accuracy_by_letter: " + "  ".join(partes_letra))

        # Costo real de API: identificación y arbitraje de dudosas por separado
        # -- antes se contaban como una sola llamada indistinta (bug real, ver
        # correr_metodo_hibrido), lo que subestimaba el costo de hojas con
        # preguntas ambiguas.
        n_id = sum(e.get("n_api_calls_identification", 0) for e in evs)
        n_arb = sum(e.get("n_api_calls_answer_arbitration", 0) for e in evs)
        if n_id or n_arb:
            print(f"n_api_calls_identification: {n_id}  |  n_api_calls_answer_arbitration: {n_arb}")

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

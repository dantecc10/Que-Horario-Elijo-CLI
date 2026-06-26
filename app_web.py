import json
import os
import uuid
from datetime import datetime, time

from flask import Flask, flash, make_response, redirect, render_template, request, session, url_for
from werkzeug.utils import secure_filename

from extract_pdf import extraer_cursos_desde_archivo

try:
    from weasyprint import HTML

    HAS_WEASYPRINT = True
except Exception:
    HAS_WEASYPRINT = False


UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"pdf", "xlsx", "xls"}
MAX_RESULTADOS = 50000
DEFAULT_LIMITE_MOSTRAR = 1000
MAX_LIMITE_MOSTRAR = 5000

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB
# Clave de desarrollo; en produccion define FLASK_SECRET_KEY.
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

# Estado en memoria para la sesion actual.
DATASTORE = {}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def convertir_hora(valor):
    if isinstance(valor, time):
        return valor
    if isinstance(valor, (float, int)):
        valor = str(int(valor)).zfill(4)
    if isinstance(valor, str):
        valor = valor.strip()
        if valor.isdigit() and len(valor) in (3, 4):
            valor = valor.zfill(4)
            return time(int(valor[:2]), int(valor[2:]))
        for format_time in ("%H:%M", "%H:%M:%S", "%H:%M:%S.%f", "%I:%M%p"):
            try:
                return datetime.strptime(valor, format_time).time()
            except ValueError:
                continue
    return None


def expandir_dias(dias_texto):
    dias_validos = "LMAJVSD"
    dias = []
    for char in dias_texto:
        if char in dias_validos and char not in dias:
            dias.append(char)
    return dias


def normalizar_cursos_a_materias(cursos):
    materias = {}
    for curso in cursos:
        materia = str(curso.get("Materia", "")).strip()
        nrc = str(curso.get("NRC", "")).strip()
        profesor = str(curso.get("Profesor", "")).strip()
        inicio = convertir_hora(curso.get("Hora de inicio"))
        fin = convertir_hora(curso.get("Hora de fin"))
        salon = str(curso.get("Salon", "")).strip()
        dias = expandir_dias(str(curso.get("Dia", "")).upper())

        if not materia or not nrc or not profesor or not inicio or not fin or not dias:
            continue

        opciones_por_materia = materias.setdefault(materia, {})
        key = (nrc, profesor)
        opcion = opciones_por_materia.setdefault(
            key,
            {
                "nrc": nrc,
                "profesor": profesor,
                "horarios": [],
            },
        )

        for dia in dias:
            bloque = {
                "inicio": inicio,
                "fin": fin,
                "dia": dia,
                "salon": salon,
            }
            if bloque not in opcion["horarios"]:
                opcion["horarios"].append(bloque)

    return {materia: list(opciones.values()) for materia, opciones in materias.items()}


def horas_entre(t1, t2):
    dt1 = datetime.combine(datetime.today(), t1)
    dt2 = datetime.combine(datetime.today(), t2)
    return (dt2 - dt1).total_seconds() / 3600


def horarios_chocan(horarios):
    por_dia = {}
    for h in horarios:
        inicio = h.get("inicio")
        fin = h.get("fin")
        dia = h.get("dia")
        if not (isinstance(inicio, time) and isinstance(fin, time) and dia):
            continue
        por_dia.setdefault(dia, []).append((inicio, fin))

    for bloques in por_dia.values():
        bloques.sort()
        for i in range(len(bloques) - 1):
            if bloques[i][1] > bloques[i + 1][0]:
                return True
    return False


def calcular_horas(combinacion):
    por_dia = {}
    horas_clase = 0.0
    hora_min = None
    hora_max = None
    for opcion in combinacion:
        for h in opcion["horarios"]:
            inicio = h["inicio"]
            fin = h["fin"]
            dia = h["dia"]
            if not (isinstance(inicio, time) and isinstance(fin, time)):
                continue
            horas_clase += horas_entre(inicio, fin)
            por_dia.setdefault(dia, []).append((inicio, fin))
            if hora_min is None or inicio < hora_min:
                hora_min = inicio
            if hora_max is None or fin > hora_max:
                hora_max = fin

    horas_permanencia = 0.0
    for bloques in por_dia.values():
        bloques.sort()
        primero = min(b[0] for b in bloques)
        ultimo = max(b[1] for b in bloques)
        horas_permanencia += horas_entre(primero, ultimo)

    hora_min_str = hora_min.strftime("%H:%M") if hora_min else ""
    hora_max_str = hora_max.strftime("%H:%M") if hora_max else ""

    return horas_clase, horas_permanencia, hora_min_str, hora_max_str


def construir_vista_semanal(resultado):
    dias_orden = ["L", "M", "A", "J", "V", "S", "D"]
    dias_label = {
        "L": "Lunes",
        "M": "Martes",
        "A": "Miercoles",
        "J": "Jueves",
        "V": "Viernes",
        "S": "Sabado",
        "D": "Domingo",
    }

    eventos = []
    marcas_tiempo = set()
    for idx, opcion in enumerate(resultado["combinacion"]):
        materia = resultado["materias"][idx]
        for h in opcion["horarios"]:
            inicio = h.get("inicio")
            fin = h.get("fin")
            dia = h.get("dia")
            if not (isinstance(inicio, time) and isinstance(fin, time) and dia in dias_label):
                continue
            marcas_tiempo.add(inicio)
            marcas_tiempo.add(fin)
            eventos.append(
                {
                    "dia": dia,
                    "inicio": inicio,
                    "fin": fin,
                    "materia": materia,
                    "nrc": opcion.get("nrc", ""),
                    "profesor": opcion.get("profesor", ""),
                    "salon": h.get("salon", ""),
                }
            )

    cortes = sorted(marcas_tiempo)
    if len(cortes) < 2:
        return {"dias": dias_orden, "labels": dias_label, "rows": []}

    rows = []
    for i in range(len(cortes) - 1):
        ini = cortes[i]
        fin = cortes[i + 1]
        cells = {d: [] for d in dias_orden}

        for ev in eventos:
            if ev["inicio"] < fin and ev["fin"] > ini:
                cells[ev["dia"]].append(ev)

        if any(cells[d] for d in dias_orden):
            rows.append(
                {
                    "slot": f"{ini.strftime('%H:%M')} - {fin.strftime('%H:%M')}",
                    "cells": cells,
                }
            )

    return {"dias": dias_orden, "labels": dias_label, "rows": rows}


def generar_horarios(materias_seleccionadas, min_materias=3, max_materias=None):
    import itertools

    materia_keys = list(materias_seleccionadas.keys())
    if not materia_keys:
        return [], False, False, 0

    resultados = []
    truncado = False
    uso_subconjuntos = False

    start = min(max_materias if max_materias is not None else len(materia_keys), len(materia_keys))
    end = max(min_materias, 1)

    for size in range(start, end - 1, -1):
        for subset in itertools.combinations(materia_keys, size):
            materia_opciones = [materias_seleccionadas[m] for m in subset]
            for combinacion in itertools.product(*materia_opciones):
                todos_horarios = []
                for opcion in combinacion:
                    todos_horarios.extend(opcion["horarios"])
                if horarios_chocan(todos_horarios):
                    continue

                horas_clase, horas_permanencia, hora_min_inicio, hora_max_fin = calcular_horas(combinacion)
                resultados.append(
                    {
                        "materias": list(subset),
                        "combinacion": combinacion,
                        "horas_clase": horas_clase,
                        "horas_permanencia": horas_permanencia,
                        "hora_min_inicio": hora_min_inicio,
                        "hora_max_fin": hora_max_fin,
                        "es_individual": False,
                    }
                )

                if len(resultados) >= MAX_RESULTADOS:
                    truncado = True
                    break
            if truncado:
                break
        if truncado:
            break

    individuales = []
    for materia in materia_keys:
        for opcion in materias_seleccionadas[materia]:
            horas_clase, horas_permanencia, hora_min_inicio, hora_max_fin = calcular_horas([opcion])
            individuales.append(
                {
                    "materias": [materia],
                    "combinacion": [opcion],
                    "horas_clase": horas_clase,
                    "horas_permanencia": horas_permanencia,
                    "hora_min_inicio": hora_min_inicio,
                    "hora_max_fin": hora_max_fin,
                    "es_individual": True,
                }
            )

    resultados.extend(individuales)

    resultados.sort(key=lambda x: (-len(x["materias"]), x["horas_permanencia"], x["horas_clase"]))
    for idx, r in enumerate(resultados, start=1):
        r["id"] = idx
        r["vista_semanal"] = construir_vista_semanal(r)

    max_materias_logradas = max((len(r["materias"]) for r in resultados if not r["es_individual"]), default=0)
    if max_materias_logradas < start:
        uso_subconjuntos = True
    return resultados, truncado, uso_subconjuntos, max_materias_logradas


def json_safe(obj):
    if isinstance(obj, time):
        return obj.strftime("%H:%M")
    if isinstance(obj, dict):
        return {k: json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [json_safe(item) for item in obj]
    return obj


def _get_state():
    state_id = session.get("state_id")
    if not state_id:
        return None
    return DATASTORE.get(state_id)


def _save_state(state):
    state_id = session.get("state_id")
    if not state_id:
        state_id = str(uuid.uuid4())
        session["state_id"] = state_id
    DATASTORE[state_id] = state


def _get_result_by_id(state, result_id):
    if not state:
        return None
    for r in state.get("resultados", []):
        if r.get("id") == result_id:
            return r
    return None


@app.route("/", methods=["GET"])
def index():
    state = _get_state() or {}
    resultados = state.get("resultados", [])
    return render_template(
        "index.html",
        materias_disponibles=state.get("materias_disponibles", []),
        resultados=resultados,
        resultados_json=json.dumps(json_safe(resultados)),
        resultados_totales=state.get("resultados_totales", 0),
        truncado=state.get("truncado", False),
        uso_subconjuntos=state.get("uso_subconjuntos", False),
        max_materias_logradas=state.get("max_materias_logradas", 0),
        total_seleccionadas=state.get("total_seleccionadas", 0),
        pdf_name=state.get("pdf_name"),
        seleccionadas=state.get("seleccionadas", []),
        limite_mostrar=state.get("limite_mostrar", DEFAULT_LIMITE_MOSTRAR),
    )


@app.route("/upload", methods=["POST"])
def upload_file():
    if "pdf" not in request.files:
        flash("No se recibio ningun archivo.", "error")
        return redirect(url_for("index"))

    file = request.files["pdf"]
    if file.filename == "":
        flash("Debes seleccionar un archivo.", "error")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash("Formato invalido. Solo se permite PDF, XLSX o XLS.", "error")
        return redirect(url_for("index"))

    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    safe_name = secure_filename(file.filename)
    save_name = f"{uuid.uuid4()}_{safe_name}"
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], save_name)
    file.save(file_path)

    try:
        cursos = extraer_cursos_desde_archivo(file_path)
    except Exception as exc:
        flash(f"No se pudo procesar el archivo: {exc}", "error")
        return redirect(url_for("index"))

    if not cursos:
        flash(
            "Se abrio el archivo, pero no se detectaron cursos. Verifica que tenga columnas/filas de materias válidas.",
            "error",
        )
        return redirect(url_for("index"))

    materias = normalizar_cursos_a_materias(cursos)
    if not materias:
        flash(
            f"Se extrajeron {len(cursos)} cursos, pero ninguno paso la validacion de horas/dias.",
            "error",
        )
        return redirect(url_for("index"))

    materias_disponibles = []
    for materia, opciones in sorted(materias.items()):
        materias_disponibles.append(
            {
                "nombre": materia,
                "opciones": len(opciones),
            }
        )

    _save_state(
        {
            "pdf_name": safe_name,
            "materias": materias,
            "materias_disponibles": materias_disponibles,
            "resultados": [],
            "resultados_totales": 0,
            "seleccionadas": [],
            "truncado": False,
            "uso_subconjuntos": False,
            "max_materias_logradas": 0,
            "total_seleccionadas": 0,
        }
    )

    flash(f"Archivo procesado. Materias detectadas: {len(materias_disponibles)}", "success")
    return redirect(url_for("index"))


@app.route("/generate", methods=["POST"])
def generate():
    state = _get_state()
    if not state or "materias" not in state:
        flash("Primero sube y procesa un PDF.", "error")
        return redirect(url_for("index"))

    seleccionadas = request.form.getlist("materias")
    if not seleccionadas:
        flash("Selecciona al menos una materia.", "error")
        return redirect(url_for("index"))

    materias_filtradas = {
        materia: state["materias"][materia]
        for materia in seleccionadas
        if materia in state["materias"]
    }
    if not materias_filtradas:
        flash("No se encontraron materias seleccionadas validas.", "error")
        return redirect(url_for("index"))

    min_materias = int(request.form.get("min_materias", 3))
    max_materias = int(request.form.get("max_materias", len(seleccionadas)))
    limite_mostrar = int(request.form.get("limite_mostrar", DEFAULT_LIMITE_MOSTRAR))

    min_materias = max(1, min(min_materias, len(materias_filtradas)))
    max_materias = max(min_materias, min(max_materias, len(materias_filtradas)))
    limite_mostrar = max(1, min(limite_mostrar, MAX_LIMITE_MOSTRAR))

    resultados, truncado, uso_subconjuntos, max_materias_logradas = generar_horarios(
        materias_filtradas, min_materias, max_materias
    )
    state["resultados_totales"] = len(resultados)
    state["resultados"] = resultados[:limite_mostrar]
    state["seleccionadas"] = seleccionadas
    state["truncado"] = truncado
    state["uso_subconjuntos"] = uso_subconjuntos
    state["max_materias_logradas"] = max_materias_logradas
    state["total_seleccionadas"] = len(materias_filtradas)
    state["limite_mostrar"] = limite_mostrar
    state["min_materias"] = min_materias
    state["max_materias"] = max_materias
    _save_state(state)

    if not resultados:
        flash("No se encontraron combinaciones sin choques.", "error")
    elif uso_subconjuntos:
        flash(
            f"No se lograron combinaciones con {max_materias} materias. Mostrando alternativas con hasta {max_materias_logradas} materias sin choque.",
            "success",
        )
    else:
        flash(f"Combinaciones generadas: {len(resultados)}", "success")

    return redirect(url_for("index"))


@app.route("/export/<int:result_id>", methods=["GET"])
def export_schedule_pdf(result_id):
    state = _get_state()
    if not state:
        flash("No hay sesion activa. Genera horarios primero.", "error")
        return redirect(url_for("index"))

    resultado = _get_result_by_id(state, result_id)
    if not resultado:
        flash("No se encontro ese horario para exportar.", "error")
        return redirect(url_for("index"))

    html = render_template(
        "export_schedule.html",
        r=resultado,
        pdf_name=state.get("pdf_name"),
        generated_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        has_weasyprint=HAS_WEASYPRINT,
    )

    if HAS_WEASYPRINT:
        try:
            pdf_bytes = HTML(string=html, base_url=request.url_root).write_pdf()
            response = make_response(pdf_bytes)
            response.headers["Content-Type"] = "application/pdf"
            response.headers["Content-Disposition"] = f'attachment; filename="horario_{result_id}.pdf"'
            return response
        except Exception as exc:
            flash(f"No se pudo generar el PDF en servidor: {exc}", "error")

    # Fallback visual: permite imprimir/guardar como PDF desde el navegador.
    return html


if __name__ == "__main__":
    PORT = int(os.environ.get("FLASK_PORT", 5000))
    app.run(debug=True, port=PORT)

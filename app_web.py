import os
import uuid
from datetime import datetime, time

from flask import Flask, flash, redirect, render_template, request, session, url_for
from werkzeug.utils import secure_filename

from extract_pdf import extraer_cursos_desde_pdf


UPLOAD_FOLDER = "uploads"
ALLOWED_EXTENSIONS = {"pdf"}
MAX_RESULTADOS = 5000
MAX_RESULTADOS_MOSTRADOS = 200

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
    for opcion in combinacion:
        for h in opcion["horarios"]:
            inicio = h["inicio"]
            fin = h["fin"]
            dia = h["dia"]
            if not (isinstance(inicio, time) and isinstance(fin, time)):
                continue
            horas_clase += horas_entre(inicio, fin)
            por_dia.setdefault(dia, []).append((inicio, fin))

    horas_permanencia = 0.0
    for bloques in por_dia.values():
        bloques.sort()
        primero = min(b[0] for b in bloques)
        ultimo = max(b[1] for b in bloques)
        horas_permanencia += horas_entre(primero, ultimo)

    return horas_clase, horas_permanencia


def generar_horarios(materias_seleccionadas):
    import itertools

    materia_keys = list(materias_seleccionadas.keys())
    if not materia_keys:
        return [], False, False, 0

    resultados = []
    truncado = False
    uso_subconjuntos = False

    def agregar_resultados_para_materias(keys):
        nonlocal truncado
        materia_opciones = [materias_seleccionadas[m] for m in keys]
        for combinacion in itertools.product(*materia_opciones):
            todos_horarios = []
            for opcion in combinacion:
                todos_horarios.extend(opcion["horarios"])
            if horarios_chocan(todos_horarios):
                continue

            horas_clase, horas_permanencia = calcular_horas(combinacion)
            resultados.append(
                {
                    "materias": list(keys),
                    "combinacion": combinacion,
                    "horas_clase": horas_clase,
                    "horas_permanencia": horas_permanencia,
                }
            )

            if len(resultados) >= MAX_RESULTADOS:
                truncado = True
                return

    # Intento principal: todas las materias seleccionadas.
    agregar_resultados_para_materias(materia_keys)

    # Fallback: si no hubo combinaciones completas, intenta subconjuntos.
    if not resultados and len(materia_keys) > 1:
        uso_subconjuntos = True
        for size in range(len(materia_keys) - 1, 0, -1):
            for subset in itertools.combinations(materia_keys, size):
                agregar_resultados_para_materias(subset)
                if truncado:
                    break
            if resultados or truncado:
                # Si ya hay resultados para el mayor tamaño posible, no bajar más.
                break

    resultados.sort(key=lambda x: (-len(x["materias"]), x["horas_permanencia"], x["horas_clase"]))
    for idx, r in enumerate(resultados, start=1):
        r["id"] = idx

    max_materias_logradas = max((len(r["materias"]) for r in resultados), default=0)
    return resultados, truncado, uso_subconjuntos, max_materias_logradas


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


@app.route("/", methods=["GET"])
def index():
    state = _get_state() or {}
    return render_template(
        "index.html",
        materias_disponibles=state.get("materias_disponibles", []),
        resultados=state.get("resultados", []),
        resultados_totales=state.get("resultados_totales", 0),
        truncado=state.get("truncado", False),
        uso_subconjuntos=state.get("uso_subconjuntos", False),
        max_materias_logradas=state.get("max_materias_logradas", 0),
        total_seleccionadas=state.get("total_seleccionadas", 0),
        pdf_name=state.get("pdf_name"),
        seleccionadas=state.get("seleccionadas", []),
        max_resultados_mostrados=MAX_RESULTADOS_MOSTRADOS,
    )


@app.route("/upload", methods=["POST"])
def upload_pdf():
    if "pdf" not in request.files:
        flash("No se recibio ningun archivo.", "error")
        return redirect(url_for("index"))

    file = request.files["pdf"]
    if file.filename == "":
        flash("Debes seleccionar un archivo PDF.", "error")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash("Formato invalido. Solo se permite PDF.", "error")
        return redirect(url_for("index"))

    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    safe_name = secure_filename(file.filename)
    save_name = f"{uuid.uuid4()}_{safe_name}"
    pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], save_name)
    file.save(pdf_path)

    try:
        cursos = extraer_cursos_desde_pdf(pdf_path)
    except Exception as exc:
        flash(f"No se pudo procesar el PDF: {exc}", "error")
        return redirect(url_for("index"))

    if not cursos:
        flash(
            "El PDF se abrio, pero no se detectaron cursos. Revisa que no sea escaneado y que tenga NRC visibles.",
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

    flash(f"PDF procesado. Materias detectadas: {len(materias_disponibles)}", "success")
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

    resultados, truncado, uso_subconjuntos, max_materias_logradas = generar_horarios(materias_filtradas)
    state["resultados_totales"] = len(resultados)
    state["resultados"] = resultados[:MAX_RESULTADOS_MOSTRADOS]
    state["seleccionadas"] = seleccionadas
    state["truncado"] = truncado
    state["uso_subconjuntos"] = uso_subconjuntos
    state["max_materias_logradas"] = max_materias_logradas
    state["total_seleccionadas"] = len(materias_filtradas)
    _save_state(state)

    if not resultados:
        flash("No se encontraron combinaciones sin choques ni siquiera usando subconjuntos de materias.", "error")
    elif uso_subconjuntos:
        flash(
            f"No hubo horario con las {len(materias_filtradas)} materias. Mostrando alternativas con hasta {max_materias_logradas} materias sin choque.",
            "success",
        )
    else:
        flash(f"Combinaciones generadas: {len(resultados)}", "success")

    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)

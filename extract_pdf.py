import pymupdf  # PyMuPDF
import pandas as pd
import re
import os
import unicodedata
import hashlib


FOOTER_PATTERNS = [
    r"P[áa]gina\s+\d+",
    r"Programaci[óo]n Acad[ée]mica\s+de\s+.*",
    r"Las secciones sombreadas se impartir[áa]n en CU2",
    r"Licenciatura en.*",
    r"BENEM[EÉ]RITA UNIVERSIDAD AUT[ÓO]NOMA DE PUEBLA",
    r"SIGNIFICADO",
    r"Lunes y mi[ée]rcoles",
    r"Martes y viernes",
    r"Lunes",
    r"Martes",
    r"Mi[ée]rcoles",
    r"Jueves",
    r"Viernes",
    r"S[áa]bado",
    r"EJEMPLO:",
    r"Horario X de.*",
    r"Horario Y de.*",
    r"Horario XX de.*",
    r"Horario YY de.*",
    r"CLAVE\s*\n\s*MATERIA\s*\n\s*NRC\s*\n\s*CUPO\s*\n\s*DIAS\s*\n\s*HORARIO UBICACI[OÓ]N\s*\n\s*DOCENTE\s*\n\s*BLOQUES NUEVO\s*\n\s*INGRESO",
    r"CLAVE\s*\n\s*NRC LISTA\s*\n\s*CRUZADA",
    r"^\s*L\s*\n\s*A\s*\n\s*M\s*\n\s*J\s*\n\s*V\s*\n\s*S\s*$",
]


def _limpiar_pies_pagina(texto):
    for patron in FOOTER_PATTERNS:
        texto = re.sub(patron, "", texto, flags=re.IGNORECASE | re.MULTILINE | re.DOTALL)
    return texto


def calcular_hash_archivo(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def limpiar_encabezados(texto):
    """
    Quita las primeras 3 líneas (encabezado general) y el encabezado de la tabla,
    aunque esté fragmentado en varias líneas.
    """
    # Buscar el primer NRC (5 dígitos consecutivos)
    match = re.search(r"\d{5}", texto)
    if match:
        # Cortar el texto desde el primer NRC encontrado
        return texto[match.start():]
    else:
        # Si no se encuentra, regresar el texto tal cual
        return texto
    
def separar_lineas_por_nrc(texto_sin_encabeza):
    # Separa el texto en líneas, cada una iniciando con un NRC (5 dígitos)
    lineas = []
    patron = re.compile(r'(?=\d{5}\b)')
    for linea in patron.split(texto_sin_encabeza):
        linea = linea.strip()
        if linea:
            # Limpiar dobles espacios que pueden interferir con el regex
            linea_limpia = re.sub(r'\s+', ' ', linea)
            lineas.append(linea_limpia)
    return lineas


def _normalizar_hhmm(valor):
    valor = valor.strip()
    h, m = valor.split(":")
    return f"{int(h):02d}{int(m):02d}"


def _es_rango_hora(texto):
    return bool(re.match(r"^\d{1,2}:\d{2}/\d{1,2}:\d{2}$", texto.strip()))


def _es_linea_id_profesor(texto):
    return bool(re.match(r"^\d{9}\s+", texto.strip()))


def _es_linea_seccion(texto):
    return bool(re.match(r"^P\d{2}-\d{3}$", texto.strip(), re.IGNORECASE))


def _es_linea_codigo_materia(texto):
    return bool(re.match(r"^[A-Za-z]{3,6}-\d{3}\b", texto.strip()))


def _es_linea_ruido(texto):
    upper = texto.upper().strip()
    if not upper:
        return True

    prefijos_ruido = (
        "PA ",
        "PROGRAMACION ACADEMICA",
        "LICENCIATURA",
        "TERMINAL EN ",
        "DERECHO ",
        "ID",
        "CLAVE",
        "NOMB.",
        "HRS.",
        "EDIF-SAL",
        "SEC.",
        "CODIGO",
        "MATERIA",
        "NRC CRE",
        "LUNES",
        "MARTES",
        "MIERCOLES",
        "JUEVES",
        "VIERNES",
        "SABADO",
        "OBS",
        "PÁGINA",
        "PAGINA",
    )
    return upper.startswith(prefijos_ruido)


def _normalizar_nombre_columna(valor):
    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = texto.replace("_", " ")
    texto = re.sub(r"[^a-z0-9\s]", "", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def _normalizar_rango_hora_excel(valor):
    texto = str(valor).strip()
    if not texto or texto.lower() == "nan" or "-" not in texto:
        return None, None
    ini, fin = [x.strip() for x in texto.split("-", 1)]
    if ":" in ini and ":" in fin:
        return _normalizar_hhmm(ini), _normalizar_hhmm(fin)
    ini = re.sub(r"\D", "", ini)
    fin = re.sub(r"\D", "", fin)
    if len(ini) in (3, 4) and len(fin) in (3, 4):
        return ini.zfill(4), fin.zfill(4)
    return None, None


def extraer_cursos_desde_excel(excel_path):
    """
    Extrae cursos desde un archivo Excel con layout tipo carga académica.
    Soporta dos formatos:
      1. Con encabezados específicos (nrc, clave, materia, dias, hora, profesor, salon).
      2. Con encabezados genéricos (Column1..ColumnN) o sin encabezados, usando posición fija:
         Col 0=NRC, 1=Clave, 2=Materia, 3=Sección, 4=Día, 5=Hora(HHMM-HHMM), 6=Profesor, 7=Salón.
    Devuelve una lista de diccionarios con el mismo esquema que los parsers PDF.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"No se encontro el archivo: {excel_path}")

    required_fields = {"nrc", "clave", "materia", "dias", "hora", "profesor", "salon"}
    all_rows = []

    xls = pd.ExcelFile(excel_path)
    for sheet_name in xls.sheet_names:
        raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        header_row_idx = None
        header_map = None
        formato_posicional = False

        # Intentar detectar encabezados con nombres requeridos
        for i in range(min(25, len(raw))):
            row_vals = [_normalizar_nombre_columna(v) for v in raw.iloc[i].tolist()]
            current_map = {}
            for j, name in enumerate(row_vals):
                if name in ("nrc", "clave", "materia", "dias", "hora", "profesor", "salon"):
                    current_map[name] = j
            if required_fields.issubset(set(current_map.keys())):
                header_row_idx = i
                header_map = current_map
                break

        # Si no se encontraron encabezados, detectar formato posicional
        if header_row_idx is None:
            for i in range(min(10, len(raw))):
                row_vals = [_normalizar_nombre_columna(v) for v in raw.iloc[i].tolist()]
                if any("column" in v for v in row_vals if v):
                    formato_posicional = True
                    header_row_idx = i + 1
                    header_map = {
                        "nrc": 0, "clave": 1, "materia": 2,
                        "dias": 4, "hora": 5, "profesor": 6, "salon": 7,
                    }
                    break

        if header_row_idx is None or header_map is None:
            continue

        if formato_posicional:
            data = raw.iloc[header_row_idx:].copy()
            data.columns = [f"col{i}" for i in range(data.shape[1])]
            rename_map = {}
            for field, col_idx in header_map.items():
                rename_map[f"col{col_idx}"] = field
            data = data.rename(columns=rename_map)
            for req in required_fields:
                if req not in data.columns:
                    data = None
                    break
            if data is None:
                continue
        else:
            data = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row_idx)
            renamed = {}
            for col in data.columns:
                norm = _normalizar_nombre_columna(col)
                if norm in required_fields:
                    renamed[col] = norm
            data = data.rename(columns=renamed)
            for req in required_fields:
                if req not in data.columns:
                    data = None
                    break
            if data is None:
                continue

        for _, row in data.iterrows():
            nrc = str(row.get("nrc", "")).strip()
            clave = str(row.get("clave", "")).strip()
            materia = str(row.get("materia", "")).strip()
            dias = str(row.get("dias", "")).strip().upper()
            profesor = str(row.get("profesor", "")).strip()
            salon = str(row.get("salon", "")).strip()
            hora_ini, hora_fin = _normalizar_rango_hora_excel(row.get("hora", ""))

            if nrc.endswith(".0"):
                nrc = nrc[:-2]
            nrc = re.sub(r"\D", "", nrc)

            if not nrc or not materia or not profesor:
                continue
            if not dias or dias == "-":
                continue
            if not hora_ini or not hora_fin:
                continue

            all_rows.append(
                {
                    "NRC": nrc,
                    "Clave": clave,
                    "Materia": materia,
                    "Profesor": profesor,
                    "Hora de inicio": hora_ini,
                    "Hora de fin": hora_fin,
                    "Dia": dias,
                    "Salon": salon,
                    "Aclaraciones": "",
                }
            )

    # Deduplicar conservando orden.
    seen = set()
    cursos = []
    for c in all_rows:
        key = (
            c["NRC"],
            c["Clave"],
            c["Materia"],
            c["Profesor"],
            c["Hora de inicio"],
            c["Hora de fin"],
            c["Dia"],
            c["Salon"],
        )
        if key in seen:
            continue
        seen.add(key)
        cursos.append(c)

    return cursos


def extraer_cursos_desde_archivo(file_path):
    """
    Detecta automáticamente el tipo de archivo soportado y ejecuta el parser adecuado.
    Soporta: PDF, XLSX, XLS.
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        return extraer_cursos_desde_pdf(file_path)
    if ext in {".xlsx", ".xls"}:
        return extraer_cursos_desde_excel(file_path)
    raise ValueError(f"Formato no soportado: {ext}")


def parsear_linea_horario(linea_texto):
    """
    Utiliza una expresión regular para extraer los datos de una línea de texto.
    Está optimizada para la estructura de 9 columnas del PDF.
    """
    # Expresión regular ajustada para ser más robusta.
    # Captura: NRC, Clave, Materia, Sección, Días, Hora, Profesor, Salón, Aclaraciones
    patron = re.compile(
        r"(\d{5})\s+"                      # 1. NRC
        r"([A-Za-z\d]+\s+[A-Za-z\d]+)\s+" # 2. Clave
        r"(.+?)\s+"                         # 3. Materia
        r"([A-Za-z\d]{2,4})\s+"            # 4. Sección (más flexible)
        r"([LMAJVSD]+)\s+"                  # 5. Días
        r"(\d{4}-\d{4})\s+"                # 6. Hora inicio-fin
        r"([A-Za-zÑÁÉÍÓÚñáéíóú\s\-.]+?)\s+" # 7. Profesor
        r"(\S+)\s*"                         # 8. Salón
        r"(.*)$",                             # 9. Aclaraciones
        re.IGNORECASE,
    )

    coincidencia = patron.search(linea_texto)
    if not coincidencia or len(coincidencia.groups()) < 9:
        return None

    grupos = coincidencia.groups()
    hora_inicio, hora_fin = grupos[5].split("-")

    datos = {
        "NRC": grupos[0].strip(),
        "Clave": grupos[1].strip(),
        "Materia": grupos[2].strip(),
        "Profesor": grupos[6].strip(),
        "Hora de inicio": hora_inicio.strip(),
        "Hora de fin": hora_fin.strip(),
        "Dia": grupos[4].strip().upper(),
        "Salon": grupos[7].strip(),
        "Aclaraciones": grupos[8].strip(),
    }

    # En algunos PDFs, el token de salón queda al final de aclaraciones,
    # o el regex capturó el salón real como "-" y el verdadero salón está
    # en medio de aclaraciones (contaminado por pies de página).
    # Buscar cualquier token con patrón de salón (edificio/sala) en aclaraciones.
    if (datos["Salon"] == "-" or _es_patron_salon(datos.get("Salon", "")) is False) and datos.get("Aclaraciones"):
        tokens = datos["Aclaraciones"].split()
        idx_salon = None
        for i, tok in enumerate(tokens):
            if re.match(r"^[A-Za-z0-9]+/[A-Za-z0-9]+$", tok):
                idx_salon = i
                break
        if idx_salon is not None:
            posible_salon = tokens[idx_salon]
            partes_antes = tokens[:idx_salon]
            partes_despues = tokens[idx_salon + 1:]
            if partes_antes:
                profesor_extra = " ".join(partes_antes).strip()
                if datos["Profesor"] == "-":
                    datos["Profesor"] = profesor_extra
                else:
                    datos["Profesor"] = f"{datos['Profesor']} {profesor_extra}".strip()
            datos["Salon"] = posible_salon
            datos["Aclaraciones"] = " ".join(partes_despues).strip()

    return datos


def _es_patron_salon(valor):
    return bool(re.match(r"^[A-Za-z0-9]+/[A-Za-z0-9]+$", valor))


def _parsear_formato_buap_page(texto_pagina):
    cursos_encontrados = []
    text_pag_sin_pies = _limpiar_pies_pagina(texto_pagina)
    text_pag_sin_encabeza = limpiar_encabezados(text_pag_sin_pies)
    lineas_utiles = separar_lineas_por_nrc(text_pag_sin_encabeza)

    for linea in lineas_utiles:
        if re.match(r"^\d{5}", linea):
            datos_curso = parsear_linea_horario(linea)
            if datos_curso:
                cursos_encontrados.append(datos_curso)

    return cursos_encontrados


def _parsear_formato_pal_page(texto_pagina):
    """
    Parser alterno para documentos tipo PAL (cuatrimestral), donde el layout
    está en columnas por día y la extracción de texto sale por renglones sueltos.
    """
    texto_pagina = _limpiar_pies_pagina(texto_pagina)
    dias = ["L", "M", "A", "J", "V", "S"]
    cursos = []

    lineas = [l.strip() for l in texto_pagina.splitlines() if l.strip()]
    i = 0
    while i < len(lineas):
        linea = lineas[i]

        if _es_linea_ruido(linea):
            i += 1
            continue

        profesor = "SIN CATEDRATICO"
        if _es_linea_id_profesor(linea):
            profesor = re.sub(r"^\d{9}\s+", "", linea).strip() or profesor
            i += 1
            if i >= len(lineas):
                break
            linea = lineas[i]

        # En este formato, la siguiente línea útil suele ser EDIF-SAL.
        salon = linea
        i += 1
        if i >= len(lineas):
            break

        # Sección (ej: P26-001)
        if _es_linea_seccion(lineas[i]):
            i += 1
            if i >= len(lineas):
                break

        codigo_materia_line = lineas[i]
        if not _es_linea_codigo_materia(codigo_materia_line):
            continue

        m_codigo = re.match(r"^([A-Za-z]{3,6}-\d{3})\s+(.+)$", codigo_materia_line)
        if not m_codigo:
            i += 1
            continue

        clave = m_codigo.group(1).strip()
        materia = m_codigo.group(2).strip()
        i += 1
        if i >= len(lineas):
            break

        # Línea 'lista' y NRC (a veces juntos, a veces separados)
        nrc = None
        linea_lista = lineas[i]
        m_nrc_inline = re.search(r"\blista\s+(\d{5})\b", linea_lista, re.IGNORECASE)
        if m_nrc_inline:
            nrc = m_nrc_inline.group(1)
            i += 1
        elif re.match(r"^lista$", linea_lista, re.IGNORECASE):
            i += 1
            if i < len(lineas) and re.match(r"^\d{5}$", lineas[i]):
                nrc = lineas[i]
                i += 1
        else:
            # No parece ser un bloque de materia válido.
            continue

        if not nrc:
            continue

        # Saltar CRE y CUP si vienen como enteros
        for _ in range(2):
            if i < len(lineas) and re.match(r"^\d+$", lineas[i]):
                i += 1

        # Tomar rangos de hora consecutivos
        rangos = []
        while i < len(lineas) and _es_rango_hora(lineas[i]):
            rangos.append(lineas[i])
            i += 1

        # Si no hay rangos, aún así registra una fila mínima para no perder materia.
        if not rangos:
            cursos.append(
                {
                    "NRC": nrc,
                    "Clave": clave,
                    "Materia": materia,
                    "Profesor": profesor,
                    "Hora de inicio": "0000",
                    "Hora de fin": "0001",
                    "Dia": "L",
                    "Salon": salon,
                    "Aclaraciones": "HORARIO_NO_DETECTADO",
                }
            )
            continue

        # Mapeo por orden de columnas visibles (L, M, A, J, V, S).
        for idx_rango, rango in enumerate(rangos):
            ini, fin = [x.strip() for x in rango.split("/")]
            dia = dias[min(idx_rango, len(dias) - 1)]
            cursos.append(
                {
                    "NRC": nrc,
                    "Clave": clave,
                    "Materia": materia,
                    "Profesor": profesor,
                    "Hora de inicio": _normalizar_hhmm(ini),
                    "Hora de fin": _normalizar_hhmm(fin),
                    "Dia": dia,
                    "Salon": salon,
                    "Aclaraciones": "",
                }
            )

    return cursos


# Mapa de bloques FCFM a días de la semana
BLOQUE_FCFM_MAP = {
    "X": ["L", "M", "J"],
    "Y": ["A", "J", "V"],
    "XX": ["L", "M"],
    "YY": ["A", "V"],
}


def _expandir_dias_fcfm(codigo_dias, hora_ini, hora_fin):
    """
    Expande el código de días del formato FCFM.
    Puede ser un bloque (X/Y/XX/YY) o días explícitos (L,M,V, A,J, etc.).
    Para bloques X/Y, J recibe hora ajustada:
      - X: J tiene solo la primera hora (fin = ini + 1h en HHMM)
      - Y: J tiene solo la segunda hora (ini = ini + 1h en HHMM)
    """
    codigo = codigo_dias.replace(" ", "").replace(",", "").upper()

    if codigo in BLOQUE_FCFM_MAP:
        dias_base = BLOQUE_FCFM_MAP[codigo]
        ini_int = int(hora_ini)
        fin_int = int(hora_fin)
        resultados = []
        for d in dias_base:
            if codigo == "X" and d == "J":
                # J recibe solo la primera hora
                fin_j = ini_int + 100
                resultados.append((d, hora_ini, f"{min(fin_j, fin_int):04d}"))
            elif codigo == "Y" and d == "J":
                # J recibe solo la segunda hora
                ini_j = ini_int + 100
                resultados.append((d, f"{min(ini_j, fin_int):04d}", hora_fin))
            else:
                resultados.append((d, hora_ini, hora_fin))
        return resultados
    else:
        # Días explícitos: expandir caracter individual
        dias_validos = set("LMAJVSD")
        resultados = []
        for ch in codigo:
            if ch in dias_validos:
                resultados.append((ch, hora_ini, hora_fin))
        return resultados


def _es_clave_fcfm(linea):
    """Reconoce líneas como 'FGMA 004', 'ACTS 001', 'FISS 003', etc."""
    return bool(re.match(r"^[A-Za-z]{3,6}\s+\d{2,4}\b", linea.strip()))


def _es_nrc_cupo(linea):
    """Reconoce líneas como '30013 50' (NRC + cupo)"""
    return bool(re.match(r"^\d{5}\s+\d{2,3}$", linea.strip()))


def _es_tiempo_salon(linea):
    """Reconoce líneas como '1000-1159 1FM4/101'"""
    return bool(re.match(r"^\d{4}-\d{4}\s+\S+/?\S*$", linea.strip()))


def _es_bloque_nuevo_ingreso(linea):
    """Reconoce códigos como 'ACT0126OT', 'FIS0126OT', 'LFA0126OT' o referencias como 'MATS 001 27693'"""
    limpia = linea.strip()
    if re.match(r"^[A-Za-z]{3,6}\s+\d{2,4}\s+\d{5}$", limpia):
        return True
    if re.match(r"^[A-Za-z]{3,6}\d{2,4}OT\b", limpia):
        return True
    return False


def _parsear_formato_fcfm_page(texto_pagina):
    """
    Parser para el formato FCFM (Facultad de Ciencias Físico Matemáticas).
    Tiene un sistema de 'bloques' (X, Y, XX, YY) para indicar días y
    un layout distinto al BUAP estándar.
    """
    texto_pagina = _limpiar_pies_pagina(texto_pagina)
    lineas_raw = [l.strip() for l in texto_pagina.split("\n") if l.strip()]

    # Filtrar ruido obvio
    lineas = []
    for l in lineas_raw:
        upper = l.upper()
        if upper.startswith(("CLAVE", "MATERIA", "NRC", "CUPO", "DIAS", "HORARIO", "DOCENTE", "BLOQUES", "NUEVO", "INGRESO", "LUNES", "MARTES", "MIÉRCOLES", "MIERCOLES", "JUEVES", "VIERNES", "SÁBADO", "SABADO", "HORARIO POR MATERIA", "SIGNIFICADO", "EJEMPLO:", "L ", "A ", "M ", "J ", "V ", "S ")):
            continue
        if re.match(r"^(L|A|M|J|V|S)\s*$", l) and len(l.strip()) <= 2:
            continue
        if re.match(r"^Horario\s+[XY]+\s+de\s+\d+", l, re.IGNORECASE):
            continue
        if "BENEMÉRITA UNIVERSIDAD" in upper or "FACULTAD DE CIENCIAS" in upper:
            continue
        if re.match(r"^[LMAJVSD,\s]+$", l) and l.strip().replace(",", "").replace(" ", "").upper() in ("L", "A", "M", "J", "V", "S", "LA", "LM", "LAM", "LMA", "AMJ", "AJV", "LAMJ", "LAMJV", "LMV", "AMJV", "LAJV", "LAJ", "MJ", "AV", "L-V", "LMJV", "AMJ", "AMV"):
            # Estas son líneas de solo días: podrían ser parte de una entrada o
            # ruido, las dejamos pasar porque pueden ser días explícitos
            lineas.append(l)
            continue
        lineas.append(l)

    cursos = []
    i = 0
    while i < len(lineas):
        linea = lineas[i]

        # Buscar inicio de bloque: una clave FCFM
        if not _es_clave_fcfm(linea):
            i += 1
            continue

        clave = linea
        materia_partes = []
        i += 1

        # Acumular líneas de materia hasta encontrar NRC+Cupo
        nrc = None
        cupo = None
        while i < len(lineas):
            if _es_nrc_cupo(lineas[i]):
                partes = lineas[i].split()
                nrc = partes[0]
                cupo = partes[1]
                i += 1
                break
            elif _es_clave_fcfm(lineas[i]):
                # Llegó otra clave sin encontrar NRC - la entrada anterior
                # no tenía NRC, reiniciar
                materia_partes = []
                break
            else:
                materia_partes.append(lineas[i])
                i += 1

        if nrc is None:
            continue

        materia = " ".join(materia_partes).strip() if materia_partes else ""
        # Si la materia está vacía, extraerla de la línea de clave
        if not materia:
            m = re.match(r"^([A-Za-z]{3,6}\s+\d{2,4})\s+(.+)$", clave)
            if m:
                clave = m.group(1).strip()
                materia = m.group(2).strip()
            else:
                # Intentar siguiente línea
                if i < len(lineas) and not _es_nrc_cupo(lineas[i]) and not _es_tiempo_salon(lineas[i]):
                    materia = lineas[i]
                    i += 1
        else:
            # Extraer clave limpia (sin materia)
            m = re.match(r"^([A-Za-z]{3,6}\s+\d{2,4})", clave)
            if m:
                clave = m.group(1).strip()

        if not materia:
            continue

        # Días/bloque
        if i >= len(lineas):
            break
        codigo_dias = lineas[i]
        i += 1

        # Tiempo y salón
        if i >= len(lineas):
            break
        if not _es_tiempo_salon(lineas[i]):
            continue
        tiempo_salon = lineas[i]
        i += 1

        m_tiempo = re.match(r"^(\d{4})-(\d{4})\s+(.+)$", tiempo_salon)
        if not m_tiempo:
            continue
        hora_ini_str = m_tiempo.group(1)
        hora_fin_str = m_tiempo.group(2)
        salon = m_tiempo.group(3).strip()

        # Docente: acumular líneas hasta encontrar
        # - un bloque nuevo ingreso
        # - otra clave FCFM
        # - o fin de página
        partes_docente = []
        bloque_code = None
        while i < len(lineas):
            if _es_bloque_nuevo_ingreso(lineas[i]):
                bloque_code = lineas[i]
                i += 1
                break
            if _es_clave_fcfm(lineas[i]):
                break
            if _es_nrc_cupo(lineas[i]):
                break
            if _es_tiempo_salon(lineas[i]):
                break

            # La línea puede contener el código de bloque al final (ej: "GARCIA - VILCHIS ANA LLUVIA ACT0126OT")
            linea_actual = lineas[i]
            m_code = re.search(r"\b([A-Za-z]{3,6}\d{2,4}OT)\b", linea_actual)
            if m_code:
                antes = linea_actual[:m_code.start()].strip()
                if antes:
                    partes_docente.append(antes)
                bloque_code = m_code.group(1)
                i += 1
                break
            partes_docente.append(linea_actual)
            i += 1

        profesor = " ".join(partes_docente).strip() if partes_docente else "SIN DOCENTE"
        # Limpiar guiones sueltos y espacios extras
        profesor = re.sub(r"\s+", " ", profesor).strip()
        profesor = re.sub(r"^\s*-\s*", "", profesor).strip()
        # Limpiar cualquier código OT residual (ej: ACT0126OT, FIS0226OT)
        profesor = re.sub(r"\b[A-Za-z]{3,6}\d{2,4}OT\b", "", profesor).strip()
        profesor = re.sub(r"\s+", " ", profesor).strip()

        # Expandir días
        horarios = _expandir_dias_fcfm(codigo_dias, hora_ini_str, hora_fin_str)
        for dia, ini, fin in horarios:
            cursos.append({
                "NRC": nrc,
                "Clave": clave,
                "Materia": materia,
                "Profesor": profesor,
                "Hora de inicio": ini,
                "Hora de fin": fin,
                "Dia": dia,
                "Salon": salon,
                "Aclaraciones": "",
            })

    return cursos


def _detectar_tipo_pdf(pdf_path):
    """
    Detecta el tipo de PDF basado en el contenido de la primera página.
    Retorna: 'buap', 'pal', 'fcfm', o None
    """
    doc = pymupdf.open(pdf_path)
    try:
        page = doc.load_page(0)
        texto = page.get_textpage().extractText().upper()
    finally:
        doc.close()

    if "FACULTAD DE CIENCIAS FÍSICO MATEMÁTICAS" in texto:
        return "fcfm"
    if "PA " in texto or "PROGRAMACION ACADEMICA" in texto:
        return "pal"
    if "NRC" in texto and "CLAVE" in texto:
        return "buap"
    return None


def extraer_cursos_desde_pdf(pdf_path):
    """
    Extrae y devuelve una lista de cursos detectados en un PDF.

    Retorna una lista de diccionarios con el mismo formato usado
    en el flujo CLI. Si no se encuentran cursos, regresa lista vacia.
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"No se encontro el archivo: {pdf_path}")

    tipo = _detectar_tipo_pdf(pdf_path)
    if tipo == "fcfm":
        return _extraer_fcfm_completo(pdf_path)

    cursos_encontrados = []
    doc = pymupdf.open(pdf_path)
    try:
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            texto_pagina = page.get_textpage().extractText()

            # Detección automática por página: elegimos el parser con más filas válidas.
            cursos_buap = _parsear_formato_buap_page(texto_pagina)
            cursos_pal = _parsear_formato_pal_page(texto_pagina)

            elegidos = cursos_buap if len(cursos_buap) >= len(cursos_pal) else cursos_pal
            for c in elegidos:
                if c not in cursos_encontrados:
                    cursos_encontrados.append(c)
    finally:
        doc.close()

    # Merge post-parseo: unificar nombres de profesor partidos
    cursos_encontrados = _merge_por_nrc(cursos_encontrados)

    return cursos_encontrados


def _extraer_fcfm_completo(pdf_path):
    """
    Extrae cursos de un PDF en formato FCFM leyendo todas las páginas
    con el parser FCFM.
    """
    doc = pymupdf.open(pdf_path)
    cursos_encontrados = []
    try:
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            texto_pagina = page.get_textpage().extractText()
            cursos = _parsear_formato_fcfm_page(texto_pagina)
            for c in cursos:
                if c not in cursos_encontrados:
                    cursos_encontrados.append(c)
    finally:
        doc.close()
    return cursos_encontrados


def _merge_por_nrc(cursos):
    """
    Agrupa cursos por NRC y unifica nombres de profesor incompletos.
    Cuando un mismo NRC tiene el mismo día/hora/salón pero profesor
    diferente (uno incompleto por error de parsing), se queda con el
    nombre más largo/completo.
    """
    from collections import defaultdict

    grupos = defaultdict(list)
    for c in cursos:
        grupos[c["NRC"]].append(c)

    result = []
    for nrc, entries in grupos.items():
        if len(entries) <= 1:
            result.extend(entries)
            continue

        # Construir un mapa: (dia, inicio, fin, salon) -> mejor profesor
        slot_prof = {}
        for c in entries:
            key = (c["Dia"], c["Hora de inicio"], c["Hora de fin"], c["Salon"])
            prof = c["Profesor"]
            if key in slot_prof:
                existente = slot_prof[key]
                if len(prof) > len(existente):
                    slot_prof[key] = prof
            else:
                slot_prof[key] = prof

        # Reconstruir entries con el mejor profesor por slot
        vistos = set()
        for c in entries:
            key = (c["Dia"], c["Hora de inicio"], c["Hora de fin"], c["Salon"])
            if key in vistos:
                continue
            vistos.add(key)
            mejor_prof = slot_prof.get(key, c["Profesor"])
            c["Profesor"] = mejor_prof
            if c not in result:
                result.append(c)

    return result


def extraer_pdf_a_excel(pdf_path, excel_path):
    """
    Extrae los cursos de un PDF y los exporta a un archivo Excel.
    """
    if not os.path.exists(pdf_path):
        print(f"❌ Error: El archivo '{pdf_path}' no se encontró.")
        return False

    try:
        cursos_encontrados = extraer_cursos_desde_pdf(pdf_path)
    except Exception as e:
        print(f"❌ Error al abrir o procesar el PDF: {e}")
        return False

    if not cursos_encontrados:
        print("❌ No se encontraron cursos con el formato esperado en el PDF.")
        print("Verifica que el PDF no sea una imagen escaneada y que la estructura sea la correcta.")
        return False

    print(f"\n✅ ¡Se encontraron {len(cursos_encontrados)} clases únicas en el PDF!")

    # --- 3. SELECCIÓN DE NRC's POR EL USUARIO ---
    nrcs_seleccionados = []
    print("\n--- Selección de Cursos ---")
    print("Ingresa el NRC de cada curso que quieras añadir a tu horario.")
    print("Cuando termines, escribe 'listo' y presiona Enter.\n")

    cursos_para_mostrar = {}
    for curso in cursos_encontrados:
        cursos_para_mostrar.setdefault(curso['NRC'], curso)
            
    for nrc, curso in cursos_para_mostrar.items():
        print(f"➡️  NRC: {curso['NRC']}, Materia: {curso['Materia']}, Profesor: {curso['Profesor']}")

    while True:
        entrada = input("\nIngresa un NRC para agregarlo (o escribe 'listo' para terminar): ").strip().lower()

        if entrada == 'listo':
            if not nrcs_seleccionados:
                print("⚠️ No seleccionaste ningún NRC. El programa terminará.")
                return
            print("\n✅ Selección finalizada. Generando el archivo de Excel...")
            break
        elif entrada.isdigit() and len(entrada) == 5:
            if entrada in nrcs_seleccionados:
                print(f"✔️  El NRC {entrada} ya había sido agregado.")
            elif entrada in cursos_para_mostrar:
                nrcs_seleccionados.append(entrada)
                print(f"👍 NRC {entrada} agregado. Seleccionados hasta ahora: {', '.join(nrcs_seleccionados)}")
            else:
                print(f"❌ El NRC {entrada} no se encontró en la lista de cursos. Intenta de nuevo.")
        else:
            print("❌ Entrada no válida. Por favor, ingresa un NRC de 5 dígitos o la palabra 'listo'.")

    # --- 4. EXPORTACIÓN A EXCEL ---
    clases_a_exportar = [curso for curso in cursos_encontrados if curso['NRC'] in nrcs_seleccionados]
    df = pd.DataFrame(clases_a_exportar)
    df = df[['NRC', 'Materia', 'Profesor', 'Hora de inicio', 'Hora de fin', 'Dia', 'Salon']]

    try:
        df.to_excel(excel_path, index=False, header=False)
        print(f"\n🎉 ¡Éxito! Se ha creado el archivo '{excel_path}' con todas las clases de los NRCs que seleccionaste.")
        return True
    except Exception as e:
        print(f"\n❌ Ocurrió un error al guardar el archivo de Excel: {e}")

    # if not cursos_encontrados:
    #     print("❌ No se encontraron cursos para exportar.")
    #     return False

    # columnas = ['NRC', 'Materia', 'Profesor', 'Hora de inicio', 'Hora de fin', 'Dia', 'Salon']
    # df = pd.DataFrame(cursos_encontrados)
    # df = df[columnas]
    # try:
    #     df.to_excel(excel_path, index=False, header=False)
    #     print(f"\n🎉 ¡Éxito! Se ha creado el archivo '{excel_path}' con todas las clases encontradas.")
    #     return True
    # except Exception as e:
    #     print(f"\n❌ Ocurrió un error al guardar el archivo de Excel: {e}")
    #     return False


def _parsear_horario_personal(pdf_path):
    """
    Parsea un PDF de Horario de Clases personal de BUAP.

    Este formato tiene una tabla con columnas de ancho fijo generada
    por la plataforma de la universidad. Cada materia puede ocupar
    1-3 filas (continuaciones con '-').

    Retorna una lista de cursos en formato estándar:
      {NRC, Clave, Materia, Profesor, Hora de inicio, Hora de fin, Dia, Salon}

    El parser usa pdftotext -layout para preservar el alineamiento
    de columnas y mapea los rangos de hora a días por posición
    horizontal de la columna.
    """
    import subprocess as _sp

    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"No se encontro el archivo: {pdf_path}")

    try:
        result = _sp.run(
            ["pdftotext", "-layout", pdf_path, "-"],
            capture_output=True, text=True, timeout=30,
        )
        texto = result.stdout
    except FileNotFoundError:
        raise RuntimeError(
            "pdftotext no esta instalado. Instala poppler-utils: "
            "sudo apt install poppler-utils"
        )
    except _sp.TimeoutExpired:
        raise RuntimeError("pdftotext tardo demasiado al procesar el PDF.")

    if not texto.strip():
        raise RuntimeError("El PDF no contiene texto extraible.")

    texto_norm = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode().upper()
    if "HORARIO DE CURSOS" not in texto_norm:
        raise RuntimeError(
            "El PDF no parece ser un Horario de Clases de BUAP. "
            "Se esperaba encontrar 'HORARIO DE CURSOS'."
        )

    lineas = texto.split("\n")

    hdr_idx = next(
        (i for i, l in enumerate(lineas)
         if "CÓDIGO" in unicodedata.normalize("NFKD", l).encode("ascii", "ignore").decode().upper()
         and "PROFESOR" in l.upper()),
        None,
    )
    if hdr_idx is None:
        raise RuntimeError("No se encontro la fila de encabezado de la tabla.")

    data_end = next(
        (i for i in range(hdr_idx + 1, len(lineas))
         if lineas[i].strip().startswith("NOTAS")),
        len(lineas),
    )
    data_lines = [
        l for l in lineas[hdr_idx + 1:data_end]
        if l.strip() and "TOTAL" not in l
    ]

    DAY_ZONES = [
        ("L", 0, 65), ("A", 65, 74), ("M", 74, 83),
        ("J", 83, 97), ("V", 97, 111), ("S", 111, 118),
        ("D", 118, 300),
    ]

    cursos = []
    curso_actual = None

    for line in data_lines:
        body = line[16:] if len(line) > 16 else line
        tokens = body.split()
        if not tokens:
            continue

        code_tok = tokens[0]
        es_nuevo = bool(re.match(r"^[A-Z]{3,6}-\d{3}$", code_tok, re.IGNORECASE))

        if es_nuevo:
            curso_actual = {
                "code": code_tok, "sec": "", "name": "",
                "nrc": "", "prof": "",
            }

            pos_after_code = body.index(code_tok) + len(code_tok)
            remainder = body[pos_after_code:]
            sec_m = re.search(r"\b(\d{3})\b", remainder)
            if sec_m:
                curso_actual["sec"] = sec_m.group(1)

            name_start = pos_after_code
            if sec_m:
                name_start = pos_after_code + sec_m.end()
            name_part = body[name_start:]
            name_m = re.match(
                r"\s*(.+?)(?:\s+\d{4}-\d{4}|\s{10,}|$)", name_part
            )
            if name_m:
                curso_actual["name"] = name_m.group(1).strip()

            cursos.append(curso_actual)

        if not curso_actual:
            continue

        if len(line) > 158:
            nrc_v = line[154:159].strip()
            if nrc_v and nrc_v.isdigit() and len(nrc_v) == 5:
                curso_actual["nrc"] = nrc_v
            prof_v = line[159:].strip()
            if prof_v:
                curso_actual["prof"] = prof_v

    resultados = []
    for c in cursos:
        nrc = c.get("nrc", "")
        if not nrc:
            continue
        materia = c["name"]
        clave = c["code"]
        profesor = c.get("prof", "") or "SIN PROFESOR"

        curso_lineas = [
            l for l in data_lines
            if c["code"] in l or (
                l.strip().startswith("-") and
                nrc and nrc in l
            )
        ]

        bloques = []
        for line in curso_lineas:
            for tm in re.finditer(r"(\d{4})-(\d{4})", line[:120]):
                pos = tm.start()
                for day, lo, hi in DAY_ZONES:
                    if lo <= pos < hi:
                        bloques.append({
                            "Dia": day,
                            "Hora de inicio": tm.group(1),
                            "Hora de fin": tm.group(2),
                        })
                        break

        salon = ""
        for line in curso_lineas:
            if len(line) > 146:
                s = line[141:147].strip()
                if s:
                    salon = s
                    break

        for b in bloques:
            resultados.append({
                "NRC": nrc,
                "Clave": clave,
                "Materia": materia,
                "Profesor": profesor,
                "Hora de inicio": b["Hora de inicio"],
                "Hora de fin": b["Hora de fin"],
                "Dia": b["Dia"],
                "Salon": salon,
            })

    return resultados


# def main():
#     """
#     Función principal que orquesta la extracción, selección y exportación usando PyMuPDF.
#     """
#     # --- 1. CONFIGURACIÓN INICIAL ---
#     pdf_path = 'PA_PRIMAVERA.pdf'
#     excel_path = 'Horarios_Seleccionados.xlsx'

#     if not os.path.exists(pdf_path):
#         print(f"❌ Error: El archivo '{pdf_path}' no se encontró.")
#         print("Asegúrate de que el PDF esté en la misma carpeta que este script.")
#         return

#     # --- 2. EXTRACCIÓN Y PROCESAMIENTO CON PyMuPDF ---
#     print(f"📄 Leyendo el PDF con PyMuPDF: '{pdf_path}'...")
#     cursos_encontrados = []
    
#     try:
#         doc = pymupdf.open(pdf_path)
#     except Exception as e:
#         print(f"❌ Error al abrir el PDF. Puede que esté dañado o protegido. Error: {e}")
#         return

#     # Extraer texto de todas las páginas y aplicar las reglas de limpieza
#     for page_num in range(len(doc)):
#         page = doc.load_page(page_num)
#         texto_pagina = page.get_textpage().extractText()
#         text_pag_sin_encabeza = limpiar_encabezados(texto_pagina)
#         lineas_utiles = separar_lineas_por_nrc(text_pag_sin_encabeza)
        
#         # Procesar y parsear cada línea útil
#         for linea in lineas_utiles:
            
#             # Solo procesar líneas que empiecen con un NRC
#             if re.match(r"^\d{5}", linea):
#                 datos_curso = parsear_linea_horario(linea)
#                 if datos_curso:
#                     if datos_curso not in cursos_encontrados:
#                         cursos_encontrados.append(datos_curso)

#     doc.close()

#     if not cursos_encontrados:
#         print("❌ No se encontraron cursos con el formato esperado en el PDF.")
#         print("Verifica que el PDF no sea una imagen escaneada y que la estructura sea la correcta.")
#         return

#     print(f"\n✅ ¡Se encontraron {len(cursos_encontrados)} clases únicas en el PDF!")

#     # --- 3. SELECCIÓN DE NRC's POR EL USUARIO ---
#     nrcs_seleccionados = []
#     print("\n--- Selección de Cursos ---")
#     print("Ingresa el NRC de cada curso que quieras añadir a tu horario.")
#     print("Cuando termines, escribe 'listo' y presiona Enter.\n")

#     cursos_para_mostrar = {}
#     for curso in cursos_encontrados:
#         cursos_para_mostrar.setdefault(curso['NRC'], curso)
            
#     for nrc, curso in cursos_para_mostrar.items():
#         print(f"➡️  NRC: {curso['NRC']}, Materia: {curso['Materia']}, Profesor: {curso['Profesor']}")

#     while True:
#         entrada = input("\nIngresa un NRC para agregarlo (o escribe 'listo' para terminar): ").strip().lower()

#         if entrada == 'listo':
#             if not nrcs_seleccionados:
#                 print("⚠️ No seleccionaste ningún NRC. El programa terminará.")
#                 return
#             print("\n✅ Selección finalizada. Generando el archivo de Excel...")
#             break
#         elif entrada.isdigit() and len(entrada) == 5:
#             if entrada in nrcs_seleccionados:
#                 print(f"✔️  El NRC {entrada} ya había sido agregado.")
#             elif entrada in cursos_para_mostrar:
#                 nrcs_seleccionados.append(entrada)
#                 print(f"👍 NRC {entrada} agregado. Seleccionados hasta ahora: {', '.join(nrcs_seleccionados)}")
#             else:
#                 print(f"❌ El NRC {entrada} no se encontró en la lista de cursos. Intenta de nuevo.")
#         else:
#             print("❌ Entrada no válida. Por favor, ingresa un NRC de 5 dígitos o la palabra 'listo'.")

#     # --- 4. EXPORTACIÓN A EXCEL ---
#     clases_a_exportar = [curso for curso in cursos_encontrados if curso['NRC'] in nrcs_seleccionados]
#     df = pd.DataFrame(clases_a_exportar)
#     df = df[['NRC', 'Materia', 'Profesor', 'Hora de inicio', 'Hora de fin', 'Dia', 'Salon']]

#     try:
#         df.to_excel(excel_path, index=False, header=False)
#         print(f"\n🎉 ¡Éxito! Se ha creado el archivo '{excel_path}' con todas las clases de los NRCs que seleccionaste.")
#     except Exception as e:
#         print(f"\n❌ Ocurrió un error al guardar el archivo de Excel: {e}")

# if __name__ == '__main__':
#     main()
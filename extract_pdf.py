import pymupdf  # PyMuPDF
import pandas as pd
import re
import os
import unicodedata


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

    # En algunos PDFs, el token de salón queda al final de aclaraciones.
    # Ejemplo real: Profesor="AMBROSIO", Salon="-", Aclaraciones="VAZQUEZ ALMA DELIA 1CCO5/107"
    if datos["Salon"] == "-" and datos["Aclaraciones"]:
        tokens = datos["Aclaraciones"].split()
        if tokens:
            posible_salon = tokens[-1]
            if re.match(r"^[A-Za-z0-9]+/[A-Za-z0-9]+$", posible_salon):
                profesor_extra = " ".join(tokens[:-1]).strip()
                if profesor_extra:
                    datos["Profesor"] = f"{datos['Profesor']} {profesor_extra}".strip()
                datos["Salon"] = posible_salon
                datos["Aclaraciones"] = ""

    return datos


def _parsear_formato_buap_page(texto_pagina):
    cursos_encontrados = []
    text_pag_sin_encabeza = limpiar_encabezados(texto_pagina)
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



def extraer_cursos_desde_pdf(pdf_path):
    """
    Extrae y devuelve una lista de cursos detectados en un PDF.

    Retorna una lista de diccionarios con el mismo formato usado
    en el flujo CLI. Si no se encuentran cursos, regresa lista vacia.
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"No se encontro el archivo: {pdf_path}")

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

    return cursos_encontrados

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
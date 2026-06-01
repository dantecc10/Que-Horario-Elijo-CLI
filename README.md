# Que Horario Elijo CLI

CLI en Python para:
- Cargar materias desde Excel.
- Generar combinaciones de horarios sin choques.
- Exportar horarios a Excel.
- (Ruta opcional) conectar cuenta de Google para flujo OAuth.

Tambien incluye interfaz web para:
- Cargar PDF desde navegador.
- Revisar materias detectadas.
- Seleccionar materias objetivo.
- Generar todas las combinaciones de horario sin choques.

## Requisitos

- Python 3.10 o superior.
- pip.

## Instalacion rapida

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Ejecutar version actual (estable)

```bash
python main.py
```

## Ejecutar interfaz web

```bash
python app_web.py
```

Luego abre en navegador:

```text
http://127.0.0.1:5000
```

Flujo web:
1. Subir PDF.
2. Marcar materias deseadas.
3. Generar combinaciones de horario.

Notas:
- Esta version usa check_dependencies.py para verificar paquetes base en el arranque.
- Si usaras import-pdf, coloca el PDF en una carpeta llamada AcademicSchedule en la raiz del proyecto.

## Ejecutar version con OAuth (experimental)

```bash
python mainNext/main_with_oAuth_unimplemented.py
```

Ademas necesitas:
- Archivo credentials.json en la raiz del proyecto (credenciales OAuth de Google).
- token.json se genera automaticamente al conectar cuenta por primera vez.

## Estructura de carpetas esperada

- SchoolSubjectList: excels de materias.
- Schedules: excels generados de horarios.
- AcademicSchedule: PDFs para import-pdf (crear manualmente si no existe).
- uploads: PDFs subidos en el flujo web (se crea automaticamente).

## Comando unico para instalar todo

Si ya tienes un entorno virtual activo:

```bash
python -m pip install -r requirements.txt
```

## Problemas comunes

- Error: python3: can't open file .../main
  - Causa: falta indicar la extension del archivo.
  - Solucion: usa python3 main.py

- No se encuentra AcademicSchedule al usar import-pdf
  - Causa: la carpeta no existe por defecto.
  - Solucion: crea la carpeta AcademicSchedule en la raiz y coloca ahi el PDF.

- Faltan dependencias
  - Solucion: ejecuta python -m pip install -r requirements.txt

## Plan recomendado de ejecucion

1. Preparar entorno
	- Crear y activar .venv.
	- Instalar dependencias desde requirements.txt.

2. Probar arranque base
	- Ejecutar python3 main.py.
	- Escribir help y luego exit para validar menu.

3. Cargar materias
	- Colocar un .xlsx en SchoolSubjectList.
	- Entrar a classes y procesar el archivo.

4. Generar y guardar horarios
	- Entrar a calendars.
	- Ejecutar generateCalendars.
	- Revisar con pushNviewCals y guardar con g.

5. Flujo PDF opcional
	- Crear AcademicSchedule y colocar el PDF.
	- Ejecutar import-pdf para generar un Excel en SchoolSubjectList.

6. Flujo OAuth opcional (experimental)
	- Colocar credentials.json en la raiz.
	- Ejecutar mainNext/main_with_oAuth_unimplemented.py.

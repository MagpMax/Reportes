from logging import config
from pathlib import Path
import unicodedata

import pandas as pd
from flask import Flask, render_template, request, jsonify

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
BITACORA = BASE_DIR / "Bitacora.xlsx"
EXCEL_PATH = BITACORA  # Alias conservado por compatibilidad.


# ============================================================
# CONFIGURACIÓN DE PROCESOS Y EQUIPOS
# ============================================================
CONFIGURACIONES = {
    "infraestructura": {
        "MAP_TITULOS": {
            "Administración de Servidores, Dominios, Sistemas Operativos":
                "Administración de Servidores, Dominios, Sistemas Operativos",
            "Administración Infraestructura Cloud":
                "Administración Infraestructura Cloud",
            "Administración Middleware": "Administración Middleware",
        },
        "EQUIPOS_ORDEN": [
            "Administración de Servidores, Dominios, Sistemas Operativos",
            "Administración Infraestructura Cloud",
            "Administración Middleware",
        ],
        "EQUIPOS": {
            "Administración de Servidores, Dominios, Sistemas Operativos": [
                "RodrigoLorca",
                "PabloSalazar",
            ],
            "Administración Infraestructura Cloud": [
                "GonzaloRodriguez",
                "CesarGonzalez",
            ],
            "Administración Middleware": [
                "OscarMora",
                "AmericoSaravia",
                "GabrielNuñez",
            ],
        },
        "NOMBRES_PROFESIONALES": {
            "RodrigoLorca": "Rodrigo Lorca",
            "PabloSalazar": "Pablo Salazar",
            "GonzaloRodriguez": "Gonzalo Rodríguez",
            "CesarGonzalez": "César González",
            "OscarMora": "Oscar Mora",
            "AmericoSaravia": "Américo Saravia",
            "GabrielNuñez": "Gabriel Núñez",
        },
    },
    "monitoreo": {
        "MAP_TITULOS": {
            "Monitoreo y Respaldos": "Monitoreo y Respaldos",
        },
        "EQUIPOS_ORDEN": ["Monitoreo y Respaldos"],
        "EQUIPOS": {
            "Monitoreo y Respaldos": [
                "JPBuzeta",
                "CarlosJara",
                "CristiánVásquez",
                "JCarlos"
            ],
        },
        "NOMBRES_PROFESIONALES": {
            "JPBuzeta": "Juan Pablo Buzeta",
            "CarlosJara": "Carlos Jara",
            "CristiánVásquez": "Cristián Vásquez",
            "JCarlos": "Juan Carlos Fuentes",
        },
    },
}


def preparar_configuracion(clave_proceso):
    configuracion = CONFIGURACIONES[clave_proceso]

    profesionales_normalizados = set()
    mapeo_equipos = {}
    mapeo_nombres = {}

    for equipo, profesionales in configuracion["EQUIPOS"].items():
        for profesional in profesionales:
            clave = normalizar(profesional)

            profesionales_normalizados.add(clave)
            mapeo_equipos[clave] = equipo
            mapeo_nombres[clave] = (
                configuracion["NOMBRES_PROFESIONALES"].get(
                    profesional, profesional
                )
            )

    return {
        **configuracion,
        "PROCESO": clave_proceso,
        "PROFESIONALES_NORMALIZADOS": profesionales_normalizados,
        "MAPEO_EQUIPOS": mapeo_equipos,
        "MAPEO_NOMBRES": mapeo_nombres,
    }

# ============================================================
# UTILIDADES DE NORMALIZACIÓN Y CONFIGURACIÓN
# ============================================================

def normalizar(texto):
    """Normaliza nombres para comparar datos de Excel y configuración."""
    texto = str(texto or "").strip().lower()
    texto = unicodedata.normalize("NFD", texto)

    texto = "".join(
        caracter
        for caracter in texto
        if unicodedata.category(caracter) != "Mn"
    )

    # Elimina todos los espacios para unificar nombres concatenados.
    return "".join(
        caracter
        for caracter in texto
        if caracter.isalnum()
    )



def obtener_configuracion(proceso=None):
    """Valida el proceso y construye sus mapas de profesionales y equipos."""
    proceso = normalizar(proceso or "infraestructura")
    if proceso not in CONFIGURACIONES:
        raise ValueError(
            "Proceso no válido. Utilice 'infraestructura' o 'monitoreo'."
        )

    config = CONFIGURACIONES[proceso]
    profesionales = [
        profesional
        for lista in config["EQUIPOS"].values()
        for profesional in lista
    ]
    mapeo_equipos = {
        normalizar(profesional): equipo
        for equipo, lista in config["EQUIPOS"].items()
        for profesional in lista
    }
    mapeo_nombres = {
        normalizar(clave): nombre
        for clave, nombre in config["NOMBRES_PROFESIONALES"].items()
    }

    return {
        **config,
        "PROCESO": proceso,
        "PROFESIONALES": profesionales,
        "PROFESIONALES_NORMALIZADOS": {
            normalizar(p) for p in profesionales
        },
        "MAPEO_EQUIPOS": mapeo_equipos,
        "MAPEO_NOMBRES": mapeo_nombres,
    }


def obtener_hojas_excel():
    """Devuelve las hojas disponibles en la bitácora interna."""
    if not BITACORA.exists():
        raise FileNotFoundError(f"No existe el archivo Excel: {BITACORA}")
    return pd.ExcelFile(BITACORA, engine="openpyxl").sheet_names



def diagnosticar_bitacora():

    print("\n========== DIAGNÓSTICO GENERAL ==========")
    print("Archivo utilizado:", BITACORA.resolve())

    excel = pd.ExcelFile(BITACORA, engine="openpyxl")

    print("Hojas encontradas:", excel.sheet_names)

    for hoja in excel.sheet_names:

        df = pd.read_excel(
            BITACORA,
            sheet_name=hoja,
            engine="openpyxl",
            header=0
        )

        print(f"\n--- HOJA: {hoja} ---")
        print("Columnas:", list(df.columns))

        if "FECHA SOLICITUD" in df.columns:
            fechas = pd.to_datetime(
                df["FECHA SOLICITUD"],
                errors="coerce",
                dayfirst=True
            )

            filas = df[fechas.dt.strftime("%Y-%m-%d") == "2026-10-23"]

            if not filas.empty:
                print("REGISTROS DEL 23-10-2026:")
                print(filas.to_string(index=False))

    print("\n========== FIN DIAGNÓSTICO ==========\n")



def enriquecer_registros(registros, config):
    """Filtra por profesional, agrega equipo y nombre visible, y ordena."""

    profesionales_validos = config["PROFESIONALES_NORMALIZADOS"]
    mapeo_equipos = config["MAPEO_EQUIPOS"]
    mapeo_nombres = config["MAPEO_NOMBRES"]

    resultado = []
    descartados_sin_nombre = 0
    descartados_no_validos = 0
    sin_equipo = 0



    for registro_original in registros:
        registro = dict(registro_original)

        profesional = str(
            registro.get("profesional")
            or registro.get("hoja_origen")
            or ""
        ).strip()

        clave_profesional = normalizar(profesional)

        print(f"Procesando registro: {registro}")


        if not clave_profesional:
            descartados_sin_nombre += 1
            continue


        equipo = mapeo_equipos.get(clave_profesional, "Sin equipo")


        if equipo == "Sin equipo":
            sin_equipo += 1

        registro["equipo"] = equipo
        registro["profesional_nombre"] = mapeo_nombres.get(
            clave_profesional, profesional
        )

        resultado.append(registro)

        for clave, valor in registro.items():
            if isinstance(valor, str):
                registro[clave] = valor.strip()


    orden = {
        equipo: indice
        for indice, equipo in enumerate(config["EQUIPOS_ORDEN"])
    }

    resultado.sort(
        key=lambda r: (
            orden.get(r.get("equipo", ""), 999),
            normalizar(r.get("profesional_nombre", "")),
            str(r.get("fecha", "")),
        )
    )


    print(f"Registros recibidos: {len(registros)}")
    print(f"Registros válidos: {len(resultado)}")
    print(f"Descartados sin profesional: {descartados_sin_nombre}")
    print(f"Descartados por filtro: {descartados_no_validos}")
    print(f"Registros sin equipo asignado: {sin_equipo}")

    return resultado


# ============================================================
# LECTURA DE BITÁCORA EXCEL
# ============================================================
def leer_bitacora(hoja="Todas", fecha_desde=None, fecha_hasta=None):

    """Lee una hoja específica o consolida todas las hojas del Excel."""
    if not BITACORA.exists():
        raise FileNotFoundError(f"No existe el archivo: {BITACORA}")

    hojas_disponibles = obtener_hojas_excel()
    if hoja == "Todas":
        hojas_a_leer = hojas_disponibles
    elif hoja in hojas_disponibles:
        hojas_a_leer = [hoja]
    else:
        raise ValueError(f"La hoja '{hoja}' no existe en el archivo Excel.")

    
    alias = {
        "fecha": [
            "fecha solicitud",
            "fecha",
            "dia"
        ],
        "periodo": [
            "periodo",
            "periodo informe mensual"
        ],
        "profesional": [
            "profesional",
            "recurso",
            "responsable"
        ],
        "tipo": [
            "inc req warning",
            "inc / req / warning",
            "tipo",
            "tipo registro"
        ],
        "actividad": [
            "actividad",
            "tarea",
            "resumen"
        ],
        "detalle": [
            "detalle observacion",
            "detalle / observacion",
            "descripcion",
            "resultado"
        ],
        "estado": [
            "estado",
            "estado en proceso pendiente finalizado",
            "status"
        ],
        "tiempo": [
            "tiempo estimado actividad",
            "tiempo",
            "horas",
            "duracion"
        ],
        "solicitante": [
            "solicitante",
            "personal solicitante",
            "area personal solicitante",
            "area / personal solicitante"
        ],
        "area": [
            "area",
            "servicio"
        ]
    }


    registros = []
    for nombre_hoja in hojas_a_leer:
        df = pd.read_excel(
            BITACORA,
            sheet_name=nombre_hoja,
            engine="openpyxl",
        )
        if df.empty:
            continue

        # Se normalizan encabezados, conservando la relación con la columna real.
        columnas_originales = list(df.columns)
        df.columns = [normalizar(c) for c in df.columns]
        columnas = {}



        for destino, alternativas in alias.items():
            alternativas_normalizadas = {normalizar(a) for a in alternativas}
            for columna in df.columns:
                if columna in alternativas_normalizadas:
                    columnas[destino] = columna
                    break

   
        # if "fecha" not in columnas:
        #     app.logger.warning(
        #         "Se omite la hoja '%s': no se encontró una columna de fecha. "
        #         "Columnas disponibles: %s",
        #         nombre_hoja,
        #         columnas_originales,
        #     )
        #     continue

        resultado = pd.DataFrame(index=df.index)
        for destino in alias:
            if destino in columnas:
                resultado[destino] = df[columnas[destino]]
            else:
                resultado[destino] = ""


        
        resultado["fecha"] = pd.to_datetime(
            resultado["fecha"], errors="coerce", dayfirst=True
        )

        # Eliminar registros sin fecha válida.
        resultado = resultado.dropna(subset=["fecha"]).copy()

        # Filtrar desde la fecha de inicio.
        if fecha_desde:
            desde = pd.to_datetime(fecha_desde)
            resultado = resultado[resultado["fecha"] >= desde].copy()

        # Filtrar hasta la fecha de término, si se especifica.
        if fecha_hasta:
            hasta = pd.to_datetime(fecha_hasta)
            resultado = resultado[resultado["fecha"] <= hasta].copy()

        # Convertir las fechas al formato utilizado por el informe.
        resultado["fecha"] = resultado["fecha"].dt.strftime("%Y-%m-%d")


        for columna in alias:
            if columna != "fecha":
                resultado[columna] = (
                    resultado[columna]
                    .fillna("")
                    .astype(str)
                    .replace({"nan": "None", "NaT": ""})
                )
                resultado[columna] = resultado[columna].replace("None", "")


        # Eliminar filas que no contienen una actividad real.
        # Una fecha y un horario por sí solos no constituyen una actividad.

        resultado["actividad"] = (
            resultado["actividad"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        antes = len(resultado)

        resultado = resultado[
            resultado["actividad"] != ""
        ].copy()

        descartados = antes - len(resultado)

        if descartados:
            print(
                f"Hoja {nombre_hoja}: "
                f"{descartados} registros descartados sin actividad."
            )

        resultado["hoja_origen"] = nombre_hoja

        registros.extend(
            resultado.to_dict(orient="records")
        )





    
    return registros


# ============================================================
# CÁLCULO DE MÉTRICAS
# ============================================================
def calcular_informe(registros, inicio, fin):
    """Filtra por período y calcula métricas compatibles con el informe actual."""
    if not inicio or not fin:
        raise ValueError("Debe indicar las fechas de inicio y término.")
    if inicio > fin:
        raise ValueError("La fecha de inicio no puede ser posterior a la fecha de término.")



    desde = pd.to_datetime(inicio, format="%Y-%m-%d").normalize()
    hasta = pd.to_datetime(fin, format="%Y-%m-%d").normalize()

    filas = []

    for r in registros:
        if not r.get("fecha"):
            continue

        fecha_registro = pd.to_datetime(
            r["fecha"],
            errors="coerce",
            format="%Y-%m-%d"
        )

        if pd.isna(fecha_registro):
            continue

        fecha_registro = fecha_registro.normalize()

        if desde <= fecha_registro <= hasta:
            filas.append(r)

        

    profesionales = {
        str(r.get("profesional", "")).strip()
        for r in filas
        if str(r.get("profesional", "")).strip()
        and normalizar(r.get("profesional", "")) != "n/a"
    }

    finalizados = sum(
        "finaliz" in normalizar(r.get("estado", ""))
        for r in filas
    )

    tipos = {}
    for r in filas:
        tipo = str(r.get("tipo", "")).strip() or "SIN TIPO"
        tipos[tipo] = tipos.get(tipo, 0) + 1

    semanas = {}
    for r in filas:
        fecha = pd.Timestamp(r["fecha"])
        lunes = fecha - pd.Timedelta(days=fecha.weekday())
        viernes = lunes + pd.Timedelta(days=4)
        etiqueta = f"{lunes:%d-%m} al {viernes:%d-%m-%Y}"
        semanas.setdefault(etiqueta, []).append(r)

    resumen_semanal = []
    for semana, actividades in semanas.items():
        resumen_semanal.append({
            "semana": semana,
            "total": len(actividades),
            "profesionales": len({
                str(r.get("profesional", "")).strip()
                for r in actividades
                if str(r.get("profesional", "")).strip()
                and normalizar(r.get("profesional", "")) != "n/a"
            }),
        })


    return {
        "actividades": filas,
        "total": len(filas),
        "profesionales": len(profesionales),
        "finalizados": finalizados,
        "pendientes": len(filas) - finalizados,
        "tipos": tipos,
        "semanas": resumen_semanal,
    }


# ============================================================
# RUTAS WEB Y API
# ============================================================
@app.get("/")
def load():
    return render_template("informe.html")



@app.get("/api/load")
def cargar_datos():
    """Carga registros del proceso seleccionado para tablas o vistas previas."""
    try:
        proceso = request.args.get("proceso", "infraestructura")
        inicio = request.args.get("inicio", "").strip()
        fin = request.args.get("fin", "").strip()

        print(
            f"[API LOAD] proceso={proceso}, "
            f"inicio={inicio}, fin={fin}"
        )

        config = obtener_configuracion(proceso)

        registros = leer_bitacora(
            "Todas",
            fecha_desde=inicio or None,
            fecha_hasta=fin or None
        )

        print(f"[API LOAD] Registros leídos: {len(registros)}")

        registros = enriquecer_registros(registros, config)

        print(f"[API LOAD] Registros filtrados: {len(registros)}")

        return jsonify({
            "ok": True,
            "proceso": config["PROCESO"],
            "registros": registros,
            "total": len(registros),
        })

    except ValueError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:
        app.logger.exception("Error al cargar la bitácora")
        return jsonify({"ok": False, "error": str(exc)}), 500



@app.route("/api/generar-informe", methods=["POST"])
def generar_informe_api():
    datos = request.get_json(silent=True) or {}

    try:
        config = obtener_configuracion(datos.get("proceso", "infraestructura"))
        inicio = str(datos.get("inicio", "")).strip()
        fin = str(datos.get("fin", "")).strip()
        hito = str(datos.get("hito", "")).strip()
        emision = str(datos.get("emision", "")).strip()


        registros = leer_bitacora(
            "Todas",
            fecha_desde=inicio,
            fecha_hasta=fin
        )

      


        registros_listos = enriquecer_registros(registros, config)

        for registro in registros_listos:
            registro["profesional"] = registro.get("profesional_nombre", "")
            registro["equipo"] = registro.get("equipo", "")
            print (registro["profesional"] )
            print (registro["equipo"] )

        informe = calcular_informe(registros_listos, inicio, fin)


        # Metadatos del informe para que el front pueda mostrarlos si corresponde.
        informe["proceso"] = config["PROCESO"]
        informe["nombre_proceso"] = (
            "Equipo NOC" if config["PROCESO"] == "monitoreo" else "Infraestructura"
        )
        informe["hito"] = hito
        informe["emision"] = emision
        informe["equipos_orden"] = config["EQUIPOS_ORDEN"]
        informe["equipos"] = config["EQUIPOS"]

        print(f"5 - Informe generado: {informe['total']} actividades, {informe['profesionales']} profesionales, {informe['finalizados']} finalizados.")

        return jsonify(informe)

    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except FileNotFoundError as exc:
        app.logger.exception("No se encontró la bitácora")
        return jsonify({"error": str(exc)}), 500
    except Exception:
        app.logger.exception("Error generando informe")
        return jsonify({"error": "No se pudo generar el informe. Revise el registro de la aplicación."}), 500


# Se conserva este endpoint por compatibilidad con versiones anteriores del front.
# El nuevo selector del front no necesita mostrar las hojas de Excel.
@app.get("/api/hojas")
def listar_hojas():
    try:
        hojas = obtener_hojas_excel()
        return jsonify({"ok": True, "hojas": ["Todas"] + hojas})
    except Exception as exc:
        app.logger.exception("Error al obtener las hojas del Excel")
        return jsonify({"ok": False, "error": str(exc)}), 500


if __name__ == "__main__":
    # debug=True solo para desarrollo local.
    app.run(debug=True)

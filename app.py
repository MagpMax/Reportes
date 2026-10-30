from logging import config
from pathlib import Path
from datetime import date, datetime
import json
import unicodedata

import pandas as pd
from flask import Flask, render_template, request, jsonify

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
BITACORA = BASE_DIR / "Bitacora.xlsx"
if not BITACORA.exists() and (BASE_DIR / "Bitacora(2).xlsx").exists():
    BITACORA = BASE_DIR / "Bitacora(2).xlsx"
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
    """Normaliza textos eliminando tildes, espacios y signos de puntuación."""
    texto = str(texto or "").strip().lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(
        caracter for caracter in texto
        if unicodedata.category(caracter) != "Mn"
    )
    return "".join(caracter for caracter in texto if caracter.isalnum())

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

        profesional = str(registro.get("profesional") or "").strip()
        clave_profesional = normalizar(profesional)


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
    """
    Lee las hojas de profesionales. Omite hojas de resumen y filas sin actividad.
    Los encabezados reales están en la primera fila de cada hoja de profesional.
    """
    if not BITACORA.exists():
        raise FileNotFoundError(f"No existe el archivo: {BITACORA}")

    hojas_disponibles = obtener_hojas_excel()
    if hoja == "Todas":
        hojas_a_leer = [
            nombre for nombre in hojas_disponibles
            if normalizar(nombre) != "calculo informes"
        ]
    elif hoja in hojas_disponibles:
        hojas_a_leer = [hoja]
    else:
        raise ValueError(f"La hoja '{hoja}' no existe en el archivo Excel.")

    alias = {
        "fecha": ["fecha solicitud", "fecha", "dia"],
        "periodo": ["periodo"],
        "profesional": ["profesional", "recurso", "responsable"],
        "tipo": ["inc req warning", "tipo", "tipo registro"],
        "actividad": ["actividad", "tarea", "resumen"],
        "detalle": ["detalle observacion", "descripcion", "resultado"],
        "estado": [
            "estado en proceso pendiente finalizado",
            "estado", "status"
        ],
        "tiempo": ["tiempo estimado actividad", "tiempo", "horas", "duracion"],
        "solicitante": [
            "area personal solicitante", "solicitante", "personal solicitante"
        ],
    }

    registros = []

    for nombre_hoja in hojas_a_leer:
        df = pd.read_excel(
            BITACORA,
            sheet_name=nombre_hoja,
            engine="openpyxl",
            header=0
        )

        if df.empty:
            continue

        columnas_originales = list(df.columns)
        df.columns = [normalizar(c) for c in df.columns]

        columnas = {}
        for destino, alternativas in alias.items():
            alternativas_normalizadas = {normalizar(a) for a in alternativas}
            for columna in df.columns:
                if columna in alternativas_normalizadas:
                    columnas[destino] = columna
                    break

        # Sin fecha o actividad no es una hoja de bitácora utilizable.
        if "fecha" not in columnas or "actividad" not in columnas:
            app.logger.warning(
                "Se omite hoja '%s': no se identificaron fecha y actividad. Columnas: %s",
                nombre_hoja, columnas_originales
            )
            continue

        resultado = pd.DataFrame(index=df.index)
        for destino in alias:
            if destino in columnas:
                resultado[destino] = df[columnas[destino]]
            else:
                resultado[destino] = ""

        resultado["fecha"] = pd.to_datetime(
            resultado["fecha"], errors="coerce"
        )
        resultado = resultado.dropna(subset=["fecha"]).copy()

        if fecha_desde:
            desde = pd.to_datetime(fecha_desde).normalize()
            resultado = resultado[
                resultado["fecha"].dt.normalize() >= desde
            ].copy()

        if fecha_hasta:
            hasta = pd.to_datetime(fecha_hasta).normalize()
            resultado = resultado[
                resultado["fecha"].dt.normalize() <= hasta
            ].copy()

        resultado["fecha"] = resultado["fecha"].dt.strftime("%Y-%m-%d")

        for columna in alias:
            if columna == "fecha":
                continue
            resultado[columna] = (
                resultado[columna]
                .fillna("")
                .astype(str)
                .replace({"nan": "", "NaT": ""})
                .str.replace("\u00a0", " ", regex=False)
                .str.strip()
            )

        # El nombre de la hoja es un respaldo cuando PROFESIONAL está vacío.
        resultado["profesional"] = resultado.apply(
            lambda fila: fila["profesional"] or nombre_hoja,
            axis=1
        )

        resultado["hoja_origen"] = nombre_hoja

        # Excluir filas que solo contienen fecha/horario y no una tarea.
        resultado["actividad"] = resultado["actividad"].fillna("").astype(str).str.strip()
        resultado = resultado[resultado["actividad"] != ""].copy()

        registros.extend(resultado.to_dict(orient="records"))

    app.logger.info("Registros leídos después de filtros: %s", len(registros))
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
            r["fecha"], errors="coerce", format="%Y-%m-%d"
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

    print(f"x - calcular_informe - Resumen semanal: {len(resumen_semanal)} semanas.")
    print(f"x - calcular_informe - Tipos de actividades: {tipos}")
    print(f"x - calcular_informe - Finalizados: {finalizados}, Pendientes: {len(filas) - finalizados}")
    print(f"x - calcular_informe - Total actividades: {len(filas)}, Total profesionales: {len(profesionales)}")
    print(f"x - calcular_informe - Primeras 5 actividades: {filas[:5]}")
    print(f"x - calcular_informe - Primeros 5 profesionales: {list(profesionales)[:5]}")
    print(f"x - calcular_informe - Primeros 5 tipos: {list(tipos.items())[:5]}")
    print(f"x - calcular_informe - Primeras 5 semanas: {resumen_semanal[:5]}")
    print(f"x - calcular_informe - Primeros 5 registros: {filas[:5]}")
    print(f"x - calcular_informe - Primeros 5 tipos: {list(tipos.items())[:5]}")


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


        print(f"1 - Generando informe para proceso: {config['PROCESO']}, desde {inicio} hasta {fin}, hito: {hito}, emisión: {emision}")


        registros = leer_bitacora(
            "Todas",
            fecha_desde=inicio,
            fecha_hasta=fin
        )

        for registro in registros:
            registro["profesional"] = str(registro.get("profesional") or "").strip()
            registro["equipo"] = str(registro.get("equipo") or "").strip()
            print (registro["profesional"] )
            print (registro["equipo"] )



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



# ============================================================
# CONSOLIDADO MENSUAL PERSISTENTE (JSON, SIN BASE DE DATOS)
# Período: día 10 de un mes al día 9 del mes siguiente.
# ============================================================

METRICAS_DIR = BASE_DIR / "metricas"

GRUPOS_MENSUALES = {
    "Administración Middleware": "Middleware",
    "Administración Infraestructura Cloud": "Administración Nube",
    "Administración de Servidores, Dominios, Sistemas Operativos":
        "Administración SO, Servidores y Dominios",
    "Monitoreo y Respaldos": "Monitoreo y Respaldo",
}

ORDEN_GRUPOS_MENSUALES = [
    "Middleware",
    "Administración Nube",
    "Administración SO, Servidores y Dominios",
    "Monitoreo y Respaldo",
    "Sin grupo",
]


def obtener_periodo_mensual(anio, mes):
    """Devuelve el período desde el día 10 hasta el día 9 del mes siguiente."""
    inicio = date(anio, mes, 10)
    if mes == 12:
        fin = date(anio + 1, 1, 9)
    else:
        fin = date(anio, mes + 1, 9)
    return inicio, fin


def obtener_mapa_grupos_mensuales():
    """Relaciona identificadores y nombres de profesionales con su grupo."""
    mapa = {}
    for configuracion in CONFIGURACIONES.values():
        for equipo, profesionales in configuracion["EQUIPOS"].items():
            grupo = GRUPOS_MENSUALES.get(equipo, equipo)
            for profesional in profesionales:
                mapa[normalizar(profesional)] = grupo
    return mapa


def calcular_metricas_mensuales(anio, mes):
    inicio, fin = obtener_periodo_mensual(anio, mes)
    registros = leer_bitacora(
        "Todas",
        fecha_desde=inicio.isoformat(),
        fecha_hasta=fin.isoformat()
    )

    mapa_grupos = obtener_mapa_grupos_mensuales()
    grupos = {}

    for registro in registros:
        actividad = str(registro.get("actividad") or "").strip()
        if not actividad:
            continue

        profesional = str(
            registro.get("profesional") or registro.get("hoja_origen") or ""
        ).strip()

        clave_profesional = normalizar(profesional)
        # La hoja es la referencia principal: cada pestaña corresponde
        # al profesional titular de esa bitácora.
        grupo = (
            mapa_grupos.get(normalizar(registro.get("hoja_origen", "")))
            or mapa_grupos.get(clave_profesional)
            or "Sin grupo"
        )

        tipo = normalizar(registro.get("tipo", ""))

        if tipo.startswith("inc"):
            categoria = "inc"
        elif "warning" in tipo:
            categoria = "warning"
        elif tipo.startswith("req"):
            categoria = "req"
        elif tipo.startswith("tarea") or tipo.startswith("task"):
            categoria = "tareas"
        else:
            continue

        if grupo not in grupos:
            grupos[grupo] = {
                "grupo": grupo,
                "inc": 0,
                "warning": 0,
                "req": 0,
                "tareas": 0,
                "total": 0
            }

        grupos[grupo][categoria] += 1
        grupos[grupo]["total"] += 1

    grupos_ordenados = [
        grupos[nombre]
        for nombre in ORDEN_GRUPOS_MENSUALES
        if nombre in grupos
    ]
    # Incluir cualquier grupo adicional al final, sin perder sus registros.
    grupos_ordenados.extend(
        grupo for nombre, grupo in grupos.items()
        if nombre not in ORDEN_GRUPOS_MENSUALES
    )

    totales = {
        "inc": sum(g["inc"] for g in grupos_ordenados),
        "warning": sum(g["warning"] for g in grupos_ordenados),
        "req": sum(g["req"] for g in grupos_ordenados),
        "tareas": sum(g["tareas"] for g in grupos_ordenados),
        "general": sum(g["total"] for g in grupos_ordenados)
    }

    return {
        "anio": anio,
        "mes": mes,
        "periodo_inicio": inicio.isoformat(),
        "periodo_fin": fin.isoformat(),
        "fecha_generacion": datetime.now().isoformat(timespec="seconds"),
        "grupos": grupos_ordenados,
        "totales": totales
    }


def ruta_archivo_metricas(anio, mes):
    return METRICAS_DIR / str(anio) / f"{anio}-{mes:02d}.json"


def guardar_metricas_mensuales(datos):
    archivo = ruta_archivo_metricas(datos["anio"], datos["mes"])
    archivo.parent.mkdir(parents=True, exist_ok=True)

    # Escritura temporal para evitar dejar un JSON incompleto si se interrumpe.
    temporal = archivo.with_suffix(".json.tmp")
    with open(temporal, "w", encoding="utf-8") as salida:
        json.dump(datos, salida, ensure_ascii=False, indent=4)
    temporal.replace(archivo)
    return archivo


def leer_metricas_guardadas(anio, mes):
    archivo = ruta_archivo_metricas(anio, mes)
    if not archivo.exists():
        return None
    with open(archivo, "r", encoding="utf-8") as entrada:
        return json.load(entrada)


@app.route("/api/metricas-mensuales/generar", methods=["POST"])
def generar_metricas_mensuales_api():
    try:
        datos = request.get_json(silent=True) or {}
        anio = int(datos.get("anio", 2026))
        mes = int(datos.get("mes", 1))

        if anio < 2025 or anio > 2027 or mes < 1 or mes > 12:
            return jsonify({
                "ok": False,
                "error": "Indique un año entre 2025 y 2027 y un mes entre 1 y 12."
            }), 400

        metricas = calcular_metricas_mensuales(anio, mes)
        archivo = guardar_metricas_mensuales(metricas)

        return jsonify({
            "ok": True,
            "mensaje": "Métricas generadas y guardadas correctamente.",
            "archivo": str(archivo.relative_to(BASE_DIR)),
            "metricas": metricas
        })

    except Exception as exc:
        app.logger.exception("Error generando métricas mensuales")
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/metricas-mensuales/<int:anio>/<int:mes>", methods=["GET"])
def consultar_metricas_mensuales_api(anio, mes):
    if anio < 2025 or anio > 2027 or mes < 1 or mes > 12:
        return jsonify({
            "ok": False,
            "error": "Indique un año entre 2025 y 2027 y un mes entre 1 y 12."
        }), 400

    datos = leer_metricas_guardadas(anio, mes)
    if datos is None:
        return jsonify({
            "ok": False,
            "error": "No existen métricas guardadas para este período."
        }), 404

    return jsonify({"ok": True, "metricas": datos})


@app.route("/api/metricas-mensuales", methods=["GET"])
def listar_metricas_mensuales_api():
    disponibles = []
    if METRICAS_DIR.exists():
        for archivo in sorted(METRICAS_DIR.glob("*/*.json")):
            disponibles.append({
                "anio": int(archivo.parent.name),
                "mes": int(archivo.stem.split("-")[-1]),
                "archivo": str(archivo.relative_to(BASE_DIR))
            })
    return jsonify({"ok": True, "periodos": disponibles})


@app.get("/estadisticas")
def estadisticas():
    return render_template("estadisticas.html")


if __name__ == "__main__":
    # debug=True solo para desarrollo local.
    app.run(debug=True)




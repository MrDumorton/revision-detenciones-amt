import io
import re
import zipfile
import unicodedata
from dataclasses import dataclass
from datetime import datetime, date, time, timedelta
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
)

# ============================================================
# CONFIGURACION GENERAL
# ============================================================

st.set_page_config(
    page_title="Revisión de Detenciones AMT vs Collahuasi",
    page_icon="🛠️",
    layout="wide",
)

MESES = {
    "ene": 1, "jan": 1,
    "feb": 2,
    "mar": 3,
    "abr": 4, "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "ago": 8, "aug": 8,
    "sep": 9, "sept": 9,
    "oct": 10,
    "nov": 11,
    "dic": 12, "dec": 12,
}

ERROR_DESCRIPCION = {
    "REGISTRO_COLLAHUASI_SIN_RESPALDO_AMT": "Registro en Collahuasi sin respaldo en DailyDowntimeLog/AMT.",
    "TRAMO_INICIA_ANTES_AMT": "El tramo de Collahuasi inicia antes del evento AMT.",
    "TRAMO_TERMINA_DESPUES_AMT": "El tramo de Collahuasi termina después del evento AMT.",
    "DURACION_COLLAHUASI_EXCEDE_AMT": "La suma de tramos Collahuasi excede la duración AMT.",
    "INICIO_EVENTO_NO_COINCIDE": "El primer tramo Collahuasi no coincide con el inicio del evento AMT.",
    "TERMINO_EVENTO_NO_COINCIDE": "El último tramo Collahuasi no coincide con el término del evento AMT.",
    "DURACION_EVENTO_NO_COINCIDE": "La duración total Collahuasi no coincide con la duración AMT.",
    "GAP_ENTRE_CORTES": "Hay espacio sin cubrir entre cortes Collahuasi dentro del mismo evento AMT.",
    "SOLAPAMIENTO_ENTRE_CORTES": "Hay cortes Collahuasi superpuestos dentro del mismo evento AMT.",
    "IN_PROGRESS_CON_TERMINO_COLLAHUASI": "AMT muestra la detención con término referencial del reporte y Detenciones Collahuasi tiene un término distinto; se deja como observación de revisión, no como error.",
    "IN_PROGRESS_AMBOS_REFERENCIAL_0800": "AMT y Detenciones Collahuasi mantienen el mismo término referencial del reporte; se deja solo como observación.",
    "IN_PROGRESS_SIN_TERMINO_COLLAHUASI": "DailyDowntimeLog muestra el evento como In Progress y Collahuasi tampoco presenta un término real distinto al término referencial del reporte.",
    "TERMINO_AMT_REFERENCIAL_RANGO": "El término AMT coincide con el término del rango descargado del reporte; no se considera error por diferencia de término ni de duración.",
    "TERMINO_COLLAHUASI_REFERENCIAL_0800": "El término Collahuasi corresponde a la fecha más reciente del archivo con hora 08:00; se considera término referencial del registro y no se evalúa como diferencia de término ni de duración.",
    "IN_PROGRESS_SIN_REGISTRO_COLLAHUASI": "DailyDowntimeLog muestra el evento como In Progress, pero no se encontraron tramos asociados en Detenciones Collahuasi.",
    "INICIO_AMT_REFERENCIAL_RANGO": "El inicio AMT coincide con el inicio del rango del reporte, por lo que se considera inicio referencial y no se evalúa como diferencia de inicio.",
}

# ============================================================
# FUNCIONES DE FECHA / TIEMPO
# ============================================================

def limpiar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def normalizar_equipo(valor) -> str:
    txt = limpiar_texto(valor).upper().replace(" ", "")
    txt = txt.replace("TN-", "TN")
    return txt




def normalizar_nombre_columna(nombre) -> str:
    """Normaliza nombres de columnas para tolerar saltos de línea, espacios y acentos."""
    txt = "" if nombre is None else str(nombre)
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = re.sub(r"[^a-zA-Z0-9]+", "", txt).lower()
    return txt


def buscar_columna(columnas, *opciones) -> Optional[str]:
    """Devuelve el nombre real de una columna buscando por variantes normalizadas."""
    mapa = {normalizar_nombre_columna(c): c for c in columnas}
    for opcion in opciones:
        clave = normalizar_nombre_columna(opcion)
        if clave in mapa:
            return mapa[clave]
    return None


def obtener_hoja_excel(archivo, preferidas: List[str]) -> str:
    """Elige una hoja disponible. Primero busca las preferidas; si no, usa la primera."""
    xls = pd.ExcelFile(archivo, engine="openpyxl")
    hojas = xls.sheet_names
    hojas_norm = {normalizar_nombre_columna(h): h for h in hojas}

    for hoja in preferidas:
        clave = normalizar_nombre_columna(hoja)
        if clave in hojas_norm:
            return hojas_norm[clave]

    if not hojas:
        raise ValueError("El archivo Excel no contiene hojas.")

    return hojas[0]


def detectar_fila_encabezado_daily(archivo, hoja: str, max_filas: int = 60) -> int:
    """Detecta la fila donde están las columnas Equip Plan, DownDate y Up Date."""
    muestra = pd.read_excel(archivo, sheet_name=hoja, header=None, nrows=max_filas, engine="openpyxl")

    for idx, row in muestra.iterrows():
        valores = [normalizar_nombre_columna(v) for v in row.tolist()]
        texto_fila = "|".join(valores)

        tiene_equipo = "equipplan" in valores or "equipplan" in texto_fila
        tiene_down = "downdate" in valores or "downdate" in texto_fila
        tiene_up = "update" in valores or "update" in texto_fila

        if tiene_equipo and tiene_down and tiene_up:
            return int(idx)

    raise ValueError(
        "No se encontró la fila de encabezados del DailyDowntimeLog. "
        "Debe existir una fila con columnas como Equip Plan, DownDate y Up Date."
    )

def parse_fecha_daily(valor) -> Optional[datetime]:
    """
    Convierte fechas del DailyDowntimeLog como:
    '13-may-26\n08:00' o '13-may-2026 08:00'
    a datetime.
    """
    if valor is None or pd.isna(valor):
        return None

    if isinstance(valor, datetime):
        return valor

    texto = str(valor).strip().lower()
    texto = texto.replace("\n", " ").replace("  ", " ")

    patron = r"(\d{1,2})[-/](\w+)[-/](\d{2,4})\s+(\d{1,2}):(\d{2})"
    m = re.search(patron, texto)
    if not m:
        return None

    dia = int(m.group(1))
    mes_txt = m.group(2)[:4].replace(".", "")
    anio = int(m.group(3))
    hora = int(m.group(4))
    minuto = int(m.group(5))

    if anio < 100:
        anio += 2000

    mes = MESES.get(mes_txt[:3])
    if mes is None:
        return None

    return datetime(anio, mes, dia, hora, minuto)


def parse_rango_daily(texto) -> Tuple[Optional[datetime], Optional[datetime]]:
    """
    Extrae el rango del reporte desde la cabecera del DailyDowntimeLog.
    Ejemplo:
    '[13-may-2026 08:00 - 19-may-2026 00:00]'
    """
    if texto is None:
        return None, None

    texto = str(texto).replace("\n", " ")
    patron = r"(\d{1,2}-\w+-\d{2,4}\s+\d{1,2}:\d{2})\s*-\s*(\d{1,2}-\w+-\d{2,4}\s+\d{1,2}:\d{2})"
    m = re.search(patron, texto, flags=re.IGNORECASE)
    if not m:
        return None, None

    return parse_fecha_daily(m.group(1)), parse_fecha_daily(m.group(2))


def combinar_fecha_hora(fecha_valor, hora_valor) -> Optional[datetime]:
    """
    Une Fecha + Hora de la hoja Detenciones Collahuasi.
    """
    if pd.isna(fecha_valor) or pd.isna(hora_valor):
        return None

    if isinstance(fecha_valor, pd.Timestamp):
        f = fecha_valor.to_pydatetime().date()
    elif isinstance(fecha_valor, datetime):
        f = fecha_valor.date()
    elif isinstance(fecha_valor, date):
        f = fecha_valor
    else:
        try:
            f = pd.to_datetime(fecha_valor).date()
        except Exception:
            return None

    if isinstance(hora_valor, pd.Timestamp):
        h = hora_valor.to_pydatetime().time()
    elif isinstance(hora_valor, datetime):
        h = hora_valor.time()
    elif isinstance(hora_valor, time):
        h = hora_valor
    elif isinstance(hora_valor, timedelta):
        total_seconds = int(hora_valor.total_seconds())
        h = (datetime(1900, 1, 1) + timedelta(seconds=total_seconds)).time()
    else:
        texto = str(hora_valor).strip()
        try:
            h = pd.to_datetime(texto).time()
        except Exception:
            return None

    return datetime.combine(f, h)


def duracion_a_horas(valor, fallback=None, numeric_as_excel_time: bool = False) -> Optional[float]:
    """
    Convierte una duración a horas.

    Importante:
    - En DailyDowntimeLog, los números como 1.83 o 20.33 ya vienen en horas.
    - En algunas hojas Excel, una duración puede venir como fracción de día.
      Para ese caso se usa numeric_as_excel_time=True.
    - Si existe fallback, se prioriza porque normalmente corresponde a Tiempo Horas.
    """
    if fallback is not None and not pd.isna(fallback):
        try:
            return float(str(fallback).replace(",", "."))
        except Exception:
            pass

    if valor is None or pd.isna(valor):
        return None

    if isinstance(valor, pd.Timedelta):
        return valor.total_seconds() / 3600

    if isinstance(valor, timedelta):
        return valor.total_seconds() / 3600

    if isinstance(valor, time):
        return valor.hour + valor.minute / 60 + valor.second / 3600

    if isinstance(valor, datetime):
        return valor.hour + valor.minute / 60 + valor.second / 3600

    if isinstance(valor, (int, float)):
        valor = float(valor)
        if numeric_as_excel_time:
            return valor * 24
        return valor

    texto = str(valor).strip().replace(",", ".")
    if texto == "":
        return None

    if ":" in texto:
        partes = texto.split(":")
        try:
            horas = int(partes[0])
            minutos = int(partes[1]) if len(partes) > 1 else 0
            segundos = int(partes[2]) if len(partes) > 2 else 0
            return horas + minutos / 60 + segundos / 3600
        except Exception:
            return None

    try:
        return float(texto)
    except Exception:
        return None


def horas_entre(inicio: datetime, termino: datetime) -> float:
    return (termino - inicio).total_seconds() / 3600


def fmt_dt(valor) -> str:
    if pd.isna(valor) or valor is None:
        return ""
    if isinstance(valor, pd.Timestamp):
        valor = valor.to_pydatetime()
    if isinstance(valor, datetime):
        return valor.strftime("%d-%m-%Y %H:%M")
    return str(valor)


def fmt_horas(valor) -> str:
    try:
        return f"{float(valor):.2f}"
    except Exception:
        return ""


def es_hora_0800(valor: datetime) -> bool:
    """Retorna True cuando un datetime tiene hora 08:00."""
    if valor is None or pd.isna(valor):
        return False
    if isinstance(valor, pd.Timestamp):
        valor = valor.to_pydatetime()
    if not isinstance(valor, datetime):
        return False
    return valor.hour == 8 and valor.minute == 0


def es_inicio_referencial_rango(
    inicio_amt: Optional[datetime],
    rango_inicio: Optional[datetime],
    tolerancia_horas: float,
) -> bool:
    """
    Retorna True cuando el inicio del evento AMT coincide con el inicio del
    rango descargado del reporte. En ese caso el inicio puede ser referencial
    y no necesariamente el inicio real de la detención.
    """
    if inicio_amt is None or rango_inicio is None:
        return False
    if pd.isna(inicio_amt) or pd.isna(rango_inicio):
        return False
    if isinstance(inicio_amt, pd.Timestamp):
        inicio_amt = inicio_amt.to_pydatetime()
    if isinstance(rango_inicio, pd.Timestamp):
        rango_inicio = rango_inicio.to_pydatetime()
    return abs((inicio_amt - rango_inicio).total_seconds()) <= tolerancia_horas * 3600




def es_termino_referencial_rango(
    termino_amt: Optional[datetime],
    rango_termino: Optional[datetime],
    tolerancia_horas: float,
) -> bool:
    """
    Retorna True cuando el término del evento AMT coincide con el término del
    rango descargado del reporte. En ese caso el término puede ser referencial
    y no necesariamente el término real de la detención.
    """
    if termino_amt is None or rango_termino is None:
        return False
    if pd.isna(termino_amt) or pd.isna(rango_termino):
        return False
    if isinstance(termino_amt, pd.Timestamp):
        termino_amt = termino_amt.to_pydatetime()
    if isinstance(rango_termino, pd.Timestamp):
        rango_termino = rango_termino.to_pydatetime()
    return abs((termino_amt - rango_termino).total_seconds()) <= tolerancia_horas * 3600




def es_termino_collahuasi_referencial_0800(
    termino_collahuasi: Optional[datetime],
    ultimo_termino_collahuasi: Optional[datetime],
    tolerancia_horas: float,
) -> bool:
    """
    Retorna True cuando el término de Collahuasi corresponde a la fecha más
    reciente registrada en la planilla y esa hora es 08:00.

    En este caso se interpreta como un posible corte del archivo Collahuasi,
    no necesariamente como término real de la detención. Se compara por fecha
    para cubrir casos donde la última fecha del archivo está asociada al cierre
    del turno a las 08:00.
    """
    if termino_collahuasi is None or ultimo_termino_collahuasi is None:
        return False
    if pd.isna(termino_collahuasi) or pd.isna(ultimo_termino_collahuasi):
        return False
    if isinstance(termino_collahuasi, pd.Timestamp):
        termino_collahuasi = termino_collahuasi.to_pydatetime()
    if isinstance(ultimo_termino_collahuasi, pd.Timestamp):
        ultimo_termino_collahuasi = ultimo_termino_collahuasi.to_pydatetime()

    return (
        termino_collahuasi.date() == ultimo_termino_collahuasi.date()
        and es_hora_0800(termino_collahuasi)
    )


def es_evento_in_progress(
    inicio: Optional[datetime],
    termino: Optional[datetime],
    rango_termino: Optional[datetime],
    detectar_cualquier_0800: bool = False,
) -> bool:
    """
    Identifica eventos AMT que probablemente están In Progress.

    Criterio principal:
    - Si el término del evento AMT coincide con el término del rango descargado
      del DailyDowntimeLog, se interpreta como posible término referencial del
      reporte. Esto aplica aunque el cierre del rango sea 00:00, 08:00 u otra hora.
      Ejemplo: reporte 01-06-2026 08:00 a 19-06-2026 00:00; si varios equipos
      terminan exactamente el 19-06-2026 00:00, probablemente siguen In Progress.

    Criterio opcional:
    - Considerar cualquier término 08:00 como posible In Progress.
      Esta opción se deja configurable porque una detención también puede terminar
      realmente a las 08:00.
    """
    if inicio is None or termino is None or pd.isna(inicio) or pd.isna(termino):
        return False

    if isinstance(inicio, pd.Timestamp):
        inicio = inicio.to_pydatetime()
    if isinstance(termino, pd.Timestamp):
        termino = termino.to_pydatetime()

    if rango_termino is not None and not pd.isna(rango_termino):
        if isinstance(rango_termino, pd.Timestamp):
            rango_termino = rango_termino.to_pydatetime()
        diferencia_seg = abs((termino - rango_termino).total_seconds())
        if diferencia_seg <= 60:
            return True

    return bool(detectar_cualquier_0800 and es_hora_0800(termino))

# ============================================================
# LECTURA DAILY DOWNTIME LOG
# ============================================================

def leer_daily_downtime_log(
    archivo,
    detectar_in_progress_por_0800: bool = False,
) -> Tuple[pd.DataFrame, pd.DataFrame, Optional[datetime], Optional[datetime]]:
    """
    Lee DailyDowntimeLog y devuelve:
    - df_eventos: eventos únicos AMT por equipo + inicio + término.
    - df_asignaciones: detalle de asignaciones HD Actual por evento.
    - rango_inicio, rango_termino: rango del reporte extraído desde cabecera.

    Esta versión detecta automáticamente la hoja y la fila de encabezados porque
    algunos archivos vienen con hoja "Table 1" y otros con hoja "DailyDowntimeLog".
    """
    hoja = obtener_hoja_excel(archivo, ["Table 1", "DailyDowntimeLog"] )

    # Leer rango desde filas superiores. En algunos reportes viene en la fila 1
    # y en otros cerca de la fila 8, por eso se leen más filas.
    cabecera = pd.read_excel(archivo, sheet_name=hoja, header=None, nrows=25, engine="openpyxl")
    texto_rango = " ".join(cabecera.astype(str).fillna("").values.flatten().tolist())
    rango_inicio, rango_termino = parse_rango_daily(texto_rango)

    # Detectar automáticamente la fila donde están los encabezados.
    fila_header = detectar_fila_encabezado_daily(archivo, hoja)
    df = pd.read_excel(archivo, sheet_name=hoja, header=fila_header, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    col_equipo = buscar_columna(df.columns, "Equip Plan")
    col_modelo = buscar_columna(df.columns, "Modelo")
    col_desc = buscar_columna(df.columns, "Descripción", "Descripcion")
    col_down = buscar_columna(df.columns, "DownDate")
    col_up = buscar_columna(df.columns, "Up Date", "UpDate")
    col_horas = buscar_columna(df.columns, "Horas detención", "Horas detenc ión", "Horas detencion")
    col_razon = buscar_columna(df.columns, "Razón", "Razon")
    col_actividad = buscar_columna(df.columns, "Actividad")
    col_varianza = buscar_columna(df.columns, "Varianza de la causa")
    col_resp = buscar_columna(df.columns, "Resp")
    col_hd = buscar_columna(df.columns, "HD Actual", "HD\nActual")

    columnas_obligatorias = {
        "Equip Plan": col_equipo,
        "DownDate": col_down,
        "Up Date": col_up,
        "Horas detención": col_horas,
        "HD Actual": col_hd,
    }
    faltantes = [nombre for nombre, columna in columnas_obligatorias.items() if columna is None]
    if faltantes:
        raise ValueError(
            "No se encontraron columnas obligatorias en DailyDowntimeLog: "
            + ", ".join(faltantes)
            + f". Hoja usada: {hoja}. Columnas detectadas: {list(df.columns)}"
        )

    eventos: Dict[str, Dict] = {}
    asignaciones: List[Dict] = []
    contexto = None

    for idx, row in df.iterrows():
        fila_excel = idx + fila_header + 2
        equipo_raw = row.get(col_equipo)
        equipo = normalizar_equipo(equipo_raw)

        if equipo == "TOTAL":
            break

        inicio = parse_fecha_daily(row.get(col_down))
        termino = parse_fecha_daily(row.get(col_up))

        # Si la fila trae equipo e inicio/término, actualiza contexto del evento.
        if equipo and inicio and termino:
            duracion_horas = duracion_a_horas(row.get(col_horas))
            if duracion_horas is None:
                duracion_horas = horas_entre(inicio, termino)

            evento_id = f"{equipo}|{inicio.isoformat()}|{termino.isoformat()}"
            contexto = {
                "evento_id": evento_id,
                "equipo": equipo,
                "modelo": limpiar_texto(row.get(col_modelo)) if col_modelo else "",
                "descripcion": limpiar_texto(row.get(col_desc)) if col_desc else "",
                "inicio_amt": inicio,
                "termino_amt": termino,
                "duracion_amt_h": float(duracion_horas),
                "in_progress_amt": es_evento_in_progress(
                    inicio=inicio,
                    termino=termino,
                    rango_termino=rango_termino,
                    detectar_cualquier_0800=detectar_in_progress_por_0800,
                ),
                "termino_amt_referencial": es_evento_in_progress(
                    inicio=inicio,
                    termino=termino,
                    rango_termino=rango_termino,
                    detectar_cualquier_0800=detectar_in_progress_por_0800,
                ),
                "fila_daily": fila_excel,
            }

            if evento_id not in eventos:
                eventos[evento_id] = contexto.copy()

        # Fila sin equipo puede ser continuación de asignación del evento anterior.
        if contexto is None:
            continue

        hd_actual = duracion_a_horas(row.get(col_hd))
        if hd_actual is None:
            continue

        if abs(hd_actual) < 0.0001:
            continue

        asignaciones.append({
            "evento_id": contexto["evento_id"],
            "equipo": contexto["equipo"],
            "inicio_amt": contexto["inicio_amt"],
            "termino_amt": contexto["termino_amt"],
            "in_progress_amt": contexto.get("in_progress_amt", False),
            "termino_amt_referencial": contexto.get("termino_amt_referencial", False),
            "descripcion_evento": contexto["descripcion"],
            "razon_amt": limpiar_texto(row.get(col_razon)) if col_razon else "",
            "actividad_amt": limpiar_texto(row.get(col_actividad)) if col_actividad else "",
            "varianza_causa_amt": limpiar_texto(row.get(col_varianza)) if col_varianza else "",
            "responsable_amt": limpiar_texto(row.get(col_resp)) if col_resp else "",
            "hd_actual_h": float(hd_actual),
            "fila_daily": fila_excel,
        })

    df_eventos = pd.DataFrame(eventos.values())
    df_asignaciones = pd.DataFrame(asignaciones)

    if not df_eventos.empty:
        df_eventos = df_eventos.sort_values(["equipo", "inicio_amt", "termino_amt"]).reset_index(drop=True)

    return df_eventos, df_asignaciones, rango_inicio, rango_termino

# ============================================================
# LECTURA DETENCIONES COLLAHUASI
# ============================================================

def leer_detenciones_collahuasi(archivo) -> pd.DataFrame:
    """
    Lee la hoja DETENCIONES 2026 y crea inicio/termino por cada registro.
    """
    df = pd.read_excel(archivo, sheet_name="DETENCIONES 2026", header=4, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    # Nombre real puede incluir salto de línea.
    col_tiempo_horas = None
    for c in df.columns:
        if "Tiempo" in c and "Horas" in c:
            col_tiempo_horas = c
            break

    registros = []
    for idx, row in df.iterrows():
        fila_excel = idx + 6
        equipo = normalizar_equipo(row.get("Equipo"))
        if not equipo:
            continue

        inicio = combinar_fecha_hora(row.get("Fecha"), row.get("Hora"))
        duracion_h = duracion_a_horas(row.get("Duracion"), fallback=row.get(col_tiempo_horas) if col_tiempo_horas else None, numeric_as_excel_time=True)

        if inicio is None or duracion_h is None:
            continue

        if duracion_h <= 0:
            continue

        termino = inicio + timedelta(hours=float(duracion_h))

        registros.append({
            "fila_collahuasi": fila_excel,
            "equipo": equipo,
            "inicio_collahuasi": inicio,
            "termino_collahuasi": termino,
            "duracion_collahuasi_h": float(duracion_h),
            "estatus": limpiar_texto(row.get("Estatus")),
            "codigo": limpiar_texto(row.get("Codigo")),
            "cat": limpiar_texto(row.get("Cat")),
            "categoria": limpiar_texto(row.get("Categoria")),
            "razon_collahuasi": limpiar_texto(row.get("Razon")),
            "comentario_collahuasi": limpiar_texto(row.get("Comentario")),
            "flota": limpiar_texto(row.get("FLOTA")),
        })

    df_registros = pd.DataFrame(registros)
    if not df_registros.empty:
        df_registros = df_registros.sort_values(["equipo", "inicio_collahuasi", "termino_collahuasi"]).reset_index(drop=True)

    return df_registros

# ============================================================
# COMPARACION
# ============================================================

def traslape_horas(a_ini: datetime, a_fin: datetime, b_ini: datetime, b_fin: datetime) -> float:
    ini = max(a_ini, b_ini)
    fin = min(a_fin, b_fin)
    if fin <= ini:
        return 0.0
    return horas_entre(ini, fin)


def distancia_minima_horas(a_ini: datetime, a_fin: datetime, b_ini: datetime, b_fin: datetime) -> float:
    """Distancia temporal entre dos intervalos si no se traslapan."""
    if a_fin < b_ini:
        return horas_entre(a_fin, b_ini)
    if b_fin < a_ini:
        return horas_entre(b_fin, a_ini)
    return 0.0


def buscar_mejor_evento_amt(registro: pd.Series, df_eventos: pd.DataFrame, tolerancia_horas: float) -> Tuple[Optional[Dict], float]:
    """
    Busca el evento AMT que mejor respalda un registro de Collahuasi.
    Criterio: mismo equipo y mayor traslape de tiempo.
    """
    if df_eventos.empty:
        return None, 0.0

    equipo = registro["equipo"]
    ini_c = registro["inicio_collahuasi"]
    fin_c = registro["termino_collahuasi"]

    candidatos = df_eventos[df_eventos["equipo"] == equipo].copy()
    if candidatos.empty:
        return None, 0.0

    mejor = None
    mejor_traslape = 0.0
    mejor_distancia = 999999.0

    for _, ev in candidatos.iterrows():
        ini_a = ev["inicio_amt"]
        fin_a = ev["termino_amt"]
        overlap = traslape_horas(ini_c, fin_c, ini_a, fin_a)
        dist = distancia_minima_horas(ini_c, fin_c, ini_a, fin_a)

        if overlap > mejor_traslape:
            mejor = ev.to_dict()
            mejor_traslape = overlap
            mejor_distancia = dist
        elif overlap == mejor_traslape and dist < mejor_distancia:
            mejor = ev.to_dict()
            mejor_distancia = dist

    # Si no existe traslape, solo acepta coincidencia cercana por tolerancia.
    if mejor_traslape <= 0 and mejor_distancia > tolerancia_horas:
        return None, 0.0

    return mejor, mejor_traslape


def detectar_continuidad(tramos: pd.DataFrame, tolerancia_horas: float) -> List[Dict]:
    """
    Detecta gaps y solapamientos entre tramos Collahuasi asociados al mismo evento AMT.
    """
    hallazgos = []
    if len(tramos) <= 1:
        return hallazgos

    orden = tramos.sort_values("inicio_collahuasi").reset_index(drop=True)

    for i in range(1, len(orden)):
        anterior = orden.loc[i - 1]
        actual = orden.loc[i]
        fin_ant = anterior["termino_collahuasi"]
        ini_act = actual["inicio_collahuasi"]
        dif_h = horas_entre(fin_ant, ini_act)

        if dif_h > tolerancia_horas:
            hallazgos.append({
                "tipo_error": "GAP_ENTRE_CORTES",
                "desde": fin_ant,
                "hasta": ini_act,
                "diferencia_h": dif_h,
                "fila_anterior": anterior["fila_collahuasi"],
                "fila_actual": actual["fila_collahuasi"],
            })
        elif dif_h < -tolerancia_horas:
            hallazgos.append({
                "tipo_error": "SOLAPAMIENTO_ENTRE_CORTES",
                "desde": ini_act,
                "hasta": fin_ant,
                "diferencia_h": abs(dif_h),
                "fila_anterior": anterior["fila_collahuasi"],
                "fila_actual": actual["fila_collahuasi"],
            })

    return hallazgos


def comparar_detenciones(
    df_collahuasi: pd.DataFrame,
    df_eventos_amt: pd.DataFrame,
    rango_inicio: Optional[datetime],
    rango_termino: Optional[datetime],
    tolerancia_minutos: int = 3,
    filtrar_por_rango_daily: bool = True,
    validar_cobertura_total: bool = False,
    validar_continuidad: bool = True,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Comparación principal:
    - Parte desde Detenciones Collahuasi.
    - Busca respaldo en DailyDowntimeLog/AMT.
    - No marca como error que un evento AMT no esté en Collahuasi; eso queda solo informativo.
    """
    tolerancia_horas = tolerancia_minutos / 60
    registros = df_collahuasi.copy()
    ultimo_termino_collahuasi = df_collahuasi["termino_collahuasi"].dropna().max() if not df_collahuasi.empty else None

    if filtrar_por_rango_daily and rango_inicio and rango_termino:
        registros = registros[
            (registros["termino_collahuasi"] > rango_inicio) &
            (registros["inicio_collahuasi"] < rango_termino)
        ].copy()

    resultados = []

    for _, reg in registros.iterrows():
        evento, overlap_h = buscar_mejor_evento_amt(reg, df_eventos_amt, tolerancia_horas)
        errores = []

        if evento is None:
            errores.append("REGISTRO_COLLAHUASI_SIN_RESPALDO_AMT")
            resultados.append({
                **reg.to_dict(),
                "evento_id": "",
                "inicio_amt": pd.NaT,
                "termino_amt": pd.NaT,
                "duracion_amt_h": None,
                "in_progress_amt": False,
                "termino_amt_referencial": False,
                "descripcion_amt": "",
                "traslape_h": 0,
                "resultado": "ERROR",
                "errores": "; ".join(errores),
                "observaciones": "",
                "detalle": ERROR_DESCRIPCION[errores[0]],
            })
            continue

        ini_c = reg["inicio_collahuasi"]
        fin_c = reg["termino_collahuasi"]
        ini_a = evento["inicio_amt"]
        fin_a = evento["termino_amt"]

        in_progress_amt = bool(evento.get("in_progress_amt", False))
        inicio_amt_referencial = es_inicio_referencial_rango(ini_a, rango_inicio, tolerancia_horas)
        termino_amt_referencial = es_termino_referencial_rango(fin_a, rango_termino, tolerancia_horas)
        termino_collahuasi_referencial = es_termino_collahuasi_referencial_0800(fin_c, ultimo_termino_collahuasi, tolerancia_horas)
        observaciones = []

        if termino_collahuasi_referencial:
            observaciones.append("TERMINO_COLLAHUASI_REFERENCIAL_0800")

        if inicio_amt_referencial and abs(horas_entre(ini_c, ini_a)) > tolerancia_horas:
            observaciones.append("INICIO_AMT_REFERENCIAL_RANGO")
        elif ini_c < ini_a - timedelta(hours=tolerancia_horas):
            errores.append("TRAMO_INICIA_ANTES_AMT")

        if fin_c > fin_a + timedelta(hours=tolerancia_horas):
            if termino_amt_referencial:
                observaciones.append("TERMINO_AMT_REFERENCIAL_RANGO")
            elif termino_collahuasi_referencial:
                pass
            elif in_progress_amt:
                errores.append("IN_PROGRESS_CON_TERMINO_COLLAHUASI")
            else:
                errores.append("TRAMO_TERMINA_DESPUES_AMT")

        # Si un tramo Collahuasi por sí solo dura más que el evento AMT, es error directo.
        # En eventos In Progress, si Collahuasi supera el término referencial AMT, también debe ser error,
        # porque significa que AMT no está cerrado con el término real que sí aparece en Collahuasi.
        if (
            not inicio_amt_referencial
            and not termino_amt_referencial
            and not termino_collahuasi_referencial
            and reg["duracion_collahuasi_h"] > float(evento["duracion_amt_h"]) + tolerancia_horas
        ):
            if in_progress_amt:
                if "IN_PROGRESS_CON_TERMINO_COLLAHUASI" not in errores:
                    errores.append("IN_PROGRESS_CON_TERMINO_COLLAHUASI")
            else:
                errores.append("DURACION_COLLAHUASI_EXCEDE_AMT")

        resultados.append({
            **reg.to_dict(),
            "evento_id": evento["evento_id"],
            "inicio_amt": ini_a,
            "termino_amt": fin_a,
            "duracion_amt_h": evento["duracion_amt_h"],
            "in_progress_amt": in_progress_amt,
            "termino_amt_referencial": bool(evento.get("termino_amt_referencial", False)),
            "descripcion_amt": evento.get("descripcion", ""),
            "traslape_h": overlap_h,
            "resultado": "OK" if not errores else "ERROR",
            "errores": "; ".join(errores),
            "observaciones": "; ".join(observaciones),
            "detalle": "; ".join(ERROR_DESCRIPCION[e] for e in errores + observaciones),
        })

    df_resultado = pd.DataFrame(resultados)

    # Validaciones agrupadas por evento AMT.
    errores_evento = []
    if not df_resultado.empty:
        mapeados = df_resultado[df_resultado["evento_id"].astype(str) != ""].copy()

        for evento_id, grupo in mapeados.groupby("evento_id"):
            evento = df_eventos_amt[df_eventos_amt["evento_id"] == evento_id].iloc[0]
            grupo_ordenado = grupo.sort_values("inicio_collahuasi")

            inicio_c = grupo_ordenado["inicio_collahuasi"].min()
            termino_c = grupo_ordenado["termino_collahuasi"].max()
            duracion_c = grupo_ordenado["duracion_collahuasi_h"].sum()
            inicio_a = evento["inicio_amt"]
            termino_a = evento["termino_amt"]
            duracion_a = float(evento["duracion_amt_h"])

            errores = []
            detalles_extra = []

            in_progress_amt = bool(evento.get("in_progress_amt", False))
            inicio_amt_referencial = es_inicio_referencial_rango(inicio_a, rango_inicio, tolerancia_horas)
            termino_amt_referencial = es_termino_referencial_rango(termino_a, rango_termino, tolerancia_horas)
            termino_collahuasi_referencial = es_termino_collahuasi_referencial_0800(termino_c, ultimo_termino_collahuasi, tolerancia_horas)

            if validar_cobertura_total:
                # Si el inicio AMT coincide con el inicio del rango del reporte, ese inicio puede
                # ser referencial. En ese caso no se marca diferencia de inicio.
                if not inicio_amt_referencial and inicio_c > inicio_a + timedelta(hours=tolerancia_horas):
                    errores.append("INICIO_EVENTO_NO_COINCIDE")

                # Si el término AMT coincide con el término del rango del reporte, ese término puede
                # ser referencial. En ese caso no se marca diferencia de término.
                if not termino_amt_referencial and not termino_collahuasi_referencial and termino_c < termino_a - timedelta(hours=tolerancia_horas):
                    errores.append("TERMINO_EVENTO_NO_COINCIDE")

                # La duración total solo se valida cuando ni el inicio ni el término vienen
                # determinados por el rango del reporte.
                if (
                    not inicio_amt_referencial
                    and not termino_amt_referencial
                    and not termino_collahuasi_referencial
                    and abs(duracion_c - duracion_a) > tolerancia_horas
                ):
                    errores.append("DURACION_EVENTO_NO_COINCIDE")

            if validar_continuidad:
                hallazgos = detectar_continuidad(grupo_ordenado, tolerancia_horas)
                for h in hallazgos:
                    errores.append(h["tipo_error"])
                    detalles_extra.append(
                        f"{h['tipo_error']} entre filas {h['fila_anterior']} y {h['fila_actual']} "
                        f"({fmt_dt(h['desde'])} a {fmt_dt(h['hasta'])}, {h['diferencia_h']:.2f} h)"
                    )

            if errores:
                errores_evento.append({
                    "evento_id": evento_id,
                    "equipo": evento["equipo"],
                    "inicio_amt": inicio_a,
                    "termino_amt": termino_a,
                    "duracion_amt_h": duracion_a,
                    "inicio_collahuasi": inicio_c,
                    "termino_collahuasi": termino_c,
                    "duracion_collahuasi_h": duracion_c,
                    "diferencia_h": duracion_c - duracion_a,
                    "errores": "; ".join(sorted(set(errores))),
                    "detalle": "; ".join(detalles_extra) if detalles_extra else "; ".join(ERROR_DESCRIPCION[e] for e in sorted(set(errores))),
                })

    df_eventos_error = pd.DataFrame(errores_evento)

    # Informativo: eventos AMT sin registros Collahuasi mapeados.
    if not df_resultado.empty:
        eventos_mapeados = set(df_resultado["evento_id"].dropna().astype(str))
    else:
        eventos_mapeados = set()

    df_amt_sin_coll = df_eventos_amt[~df_eventos_amt["evento_id"].astype(str).isin(eventos_mapeados)].copy()
    if not df_amt_sin_coll.empty:
        df_amt_sin_coll["observacion"] = "Informativo: evento AMT no encontrado en Collahuasi. No se considera error principal."

    return df_resultado, df_eventos_error, df_amt_sin_coll

# ============================================================
# REVISION ESPECIAL IN PROGRESS
# ============================================================

def revisar_eventos_in_progress(
    df_eventos_amt: pd.DataFrame,
    df_collahuasi: pd.DataFrame,
    rango_termino: Optional[datetime],
    tolerancia_minutos: int = 3,
    ventana_busqueda_horas: int = 168,
    max_gap_horas: float = 2.0,
) -> pd.DataFrame:
    """
    Revisa eventos AMT marcados como In Progress.

    Objetivo:
    - DailyDowntimeLog puede mostrar el término igual al fin del rango descargado
      del reporte, por ejemplo 19-06-2026 00:00. Eso no necesariamente es el término
      real de la detención: puede significar que la tarea sigue In Progress en AMT.
    - Para esos eventos se busca en Detenciones Collahuasi si existe un término real
      distinto al término referencial AMT.

    Resultado:
    - OK / observación: si el término AMT coincide con el término del rango del reporte,
      no se considera error que Collahuasi tenga un término distinto, porque ese horario
      puede ser solo el cierre del reporte.
    - OK / observación: AMT y Collahuasi mantienen el mismo término referencial.
    - ERROR: no se encontraron tramos asociados en Collahuasi.
    """
    if df_eventos_amt.empty:
        return pd.DataFrame()

    if "in_progress_amt" not in df_eventos_amt.columns:
        return pd.DataFrame()

    eventos_ip = df_eventos_amt[df_eventos_amt["in_progress_amt"].fillna(False).astype(bool)].copy()
    if eventos_ip.empty:
        return pd.DataFrame()

    tolerancia_horas = tolerancia_minutos / 60
    tolerancia = timedelta(hours=tolerancia_horas)
    max_gap = timedelta(hours=max_gap_horas)
    ultimo_termino_collahuasi = df_collahuasi["termino_collahuasi"].dropna().max() if not df_collahuasi.empty else None
    filas = []

    for _, ev in eventos_ip.iterrows():
        equipo = ev["equipo"]
        inicio_amt = ev["inicio_amt"]
        termino_ref = ev["termino_amt"]

        limite_busqueda = termino_ref + timedelta(hours=ventana_busqueda_horas)
        if rango_termino is not None and not pd.isna(rango_termino):
            limite_busqueda = max(limite_busqueda, rango_termino + timedelta(hours=ventana_busqueda_horas))

        candidatos = df_collahuasi[
            (df_collahuasi["equipo"] == equipo)
            & (df_collahuasi["termino_collahuasi"] > inicio_amt - tolerancia)
            & (df_collahuasi["inicio_collahuasi"] <= limite_busqueda)
        ].copy()

        candidatos = candidatos.sort_values("inicio_collahuasi").reset_index(drop=True)

        cadena = []
        fin_actual = termino_ref
        cadena_iniciada = False

        for _, tramo in candidatos.iterrows():
            ini_c = tramo["inicio_collahuasi"]
            fin_c = tramo["termino_collahuasi"]

            # Primer tramo: debe traslapar el evento AMT o comenzar cerca del término referencial.
            if not cadena_iniciada:
                traslapa_evento = fin_c > inicio_amt - tolerancia and ini_c <= termino_ref + max_gap
                if not traslapa_evento:
                    continue
                cadena.append(tramo)
                fin_actual = max(fin_actual, fin_c)
                cadena_iniciada = True
                continue

            # Siguientes tramos: se consideran continuación si no existe un gap relevante.
            if ini_c <= fin_actual + max_gap:
                cadena.append(tramo)
                fin_actual = max(fin_actual, fin_c)
            else:
                break

        if not cadena:
            filas.append({
                "evento_id": ev["evento_id"],
                "equipo": equipo,
                "inicio_amt": inicio_amt,
                "termino_amt_referencial": termino_ref,
                "termino_collahuasi_detectado": pd.NaT,
                "duracion_amt_referencial_h": ev["duracion_amt_h"],
                "duracion_collahuasi_detectada_h": 0,
                "diferencia_vs_referencial_h": 0,
                "filas_collahuasi": "",
                "resultado": "ERROR",
                "errores": "IN_PROGRESS_SIN_REGISTRO_COLLAHUASI",
                "detalle": ERROR_DESCRIPCION["IN_PROGRESS_SIN_REGISTRO_COLLAHUASI"],
                "descripcion_amt": ev.get("descripcion", ""),
            })
            continue

        df_cadena = pd.DataFrame([t.to_dict() for t in cadena])
        termino_coll = df_cadena["termino_collahuasi"].max()
        duracion_coll = df_cadena["duracion_collahuasi_h"].sum()
        diferencia = horas_entre(termino_ref, termino_coll)
        filas_str = ", ".join(str(int(x)) for x in df_cadena["fila_collahuasi"].dropna().tolist())

        termino_collahuasi_referencial = es_termino_collahuasi_referencial_0800(
            termino_coll, ultimo_termino_collahuasi, tolerancia_horas
        )

        if termino_collahuasi_referencial:
            resultado = "OK"
            errores = ""
            detalle = ERROR_DESCRIPCION["TERMINO_COLLAHUASI_REFERENCIAL_0800"]
        elif abs(horas_entre(termino_ref, termino_coll)) > tolerancia_horas:
            resultado = "OK"
            errores = ""
            detalle = ERROR_DESCRIPCION["TERMINO_AMT_REFERENCIAL_RANGO"]
        else:
            resultado = "OK"
            errores = ""
            detalle = ERROR_DESCRIPCION["IN_PROGRESS_AMBOS_REFERENCIAL_0800"]

        filas.append({
            "evento_id": ev["evento_id"],
            "equipo": equipo,
            "inicio_amt": inicio_amt,
            "termino_amt_referencial": termino_ref,
            "termino_collahuasi_detectado": termino_coll,
            "duracion_amt_referencial_h": ev["duracion_amt_h"],
            "duracion_collahuasi_detectada_h": duracion_coll,
            "diferencia_vs_referencial_h": diferencia,
            "filas_collahuasi": filas_str,
            "resultado": resultado,
            "errores": errores,
            "detalle": detalle,
            "descripcion_amt": ev.get("descripcion", ""),
        })

    return pd.DataFrame(filas)

# ============================================================
# PDF
# ============================================================

def agregar_pie_pagina(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)
    canvas.drawString(1.5 * cm, 1 * cm, f"Página {doc.page}")
    canvas.drawRightString(28 * cm, 1 * cm, "Revisión Detenciones Collahuasi vs DailyDowntimeLog/AMT")
    canvas.restoreState()


def tabla_pdf(data, col_widths=None, font_size=7):
    t = Table(data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F6FA")]),
    ]))
    return t


def generar_pdf(
    df_resultado: pd.DataFrame,
    df_eventos_error: pd.DataFrame,
    df_amt_sin_coll: pd.DataFrame,
    df_in_progress: pd.DataFrame,
    rango_inicio: Optional[datetime],
    rango_termino: Optional[datetime],
    tolerancia_minutos: int,
) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=1.2 * cm,
        rightMargin=1.2 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.5 * cm,
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="TituloCentro",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=16,
        leading=20,
    ))
    styles.add(ParagraphStyle(
        name="Subtitulo",
        parent=styles["Heading2"],
        fontSize=11,
        leading=14,
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        name="Texto",
        parent=styles["Normal"],
        fontSize=9,
        leading=12,
    ))

    elementos = []
    elementos.append(Paragraph("Informe de Revisión de Detenciones", styles["TituloCentro"]))
    elementos.append(Paragraph("DailyDowntimeLog / AMT vs Detenciones Collahuasi", styles["TituloCentro"]))
    elementos.append(Spacer(1, 0.4 * cm))

    rango_txt = f"{fmt_dt(rango_inicio)} a {fmt_dt(rango_termino)}" if rango_inicio and rango_termino else "No detectado"
    total_coll = len(df_resultado)
    errores_reg = len(df_resultado[df_resultado["resultado"] == "ERROR"]) if not df_resultado.empty else 0
    ok_reg = len(df_resultado[df_resultado["resultado"] == "OK"]) if not df_resultado.empty else 0
    errores_evento = len(df_eventos_error)
    informativos = len(df_amt_sin_coll)
    total_in_progress = len(df_in_progress) if df_in_progress is not None and not df_in_progress.empty else 0
    errores_in_progress = len(df_in_progress[df_in_progress["resultado"] == "ERROR"]) if total_in_progress else 0

    resumen = [
        ["Rango DailyDowntimeLog", rango_txt],
        ["Tolerancia aplicada", f"{tolerancia_minutos} minutos"],
        ["Registros Collahuasi revisados", str(total_coll)],
        ["Registros correctos", str(ok_reg)],
        ["Registros con error", str(errores_reg)],
        ["Eventos con diferencias de cortes/continuidad", str(errores_evento)],
        ["Eventos AMT sin Collahuasi", f"{informativos} (informativo)"],
        ["Eventos AMT In Progress revisados", str(total_in_progress)],
        ["Eventos In Progress con alerta", str(errores_in_progress)],
    ]
    elementos.append(Paragraph("Resumen ejecutivo", styles["Subtitulo"]))
    elementos.append(tabla_pdf([["Indicador", "Valor"]] + resumen, col_widths=[7 * cm, 14 * cm], font_size=8))
    elementos.append(Spacer(1, 0.4 * cm))

    # Errores a nivel registro.
    elementos.append(Paragraph("Errores a nivel registro Collahuasi", styles["Subtitulo"]))
    if df_resultado.empty or errores_reg == 0:
        elementos.append(Paragraph("No se detectaron errores a nivel registro.", styles["Texto"]))
    else:
        cols = [
            "fila_collahuasi", "equipo", "inicio_collahuasi", "termino_collahuasi",
            "inicio_amt", "termino_amt", "duracion_collahuasi_h", "duracion_amt_h", "errores"
        ]
        vista = df_resultado[df_resultado["resultado"] == "ERROR"].copy()
        vista = vista[cols].head(80)
        data = [[
            "Fila", "Equipo", "Inicio Coll.", "Término Coll.",
            "Inicio AMT", "Término AMT", "H Coll.", "H AMT", "Error"
        ]]
        for _, r in vista.iterrows():
            data.append([
                str(r["fila_collahuasi"]),
                r["equipo"],
                fmt_dt(r["inicio_collahuasi"]),
                fmt_dt(r["termino_collahuasi"]),
                fmt_dt(r["inicio_amt"]),
                fmt_dt(r["termino_amt"]),
                fmt_horas(r["duracion_collahuasi_h"]),
                fmt_horas(r["duracion_amt_h"]),
                Paragraph(str(r["errores"]), styles["Texto"]),
            ])
        elementos.append(tabla_pdf(data, col_widths=[1.3*cm, 1.5*cm, 2.6*cm, 2.6*cm, 2.6*cm, 2.6*cm, 1.4*cm, 1.4*cm, 7.0*cm], font_size=6))

    # Errores agrupados por evento.
    elementos.append(PageBreak())
    elementos.append(Paragraph("Errores de cortes / continuidad por evento AMT", styles["Subtitulo"]))
    if df_eventos_error.empty:
        elementos.append(Paragraph("No se detectaron errores agrupados por evento.", styles["Texto"]))
    else:
        data = [[
            "Equipo", "Inicio AMT", "Término AMT", "H AMT",
            "Inicio Coll.", "Término Coll.", "H Coll.", "Dif. H", "Error"
        ]]
        for _, r in df_eventos_error.head(80).iterrows():
            data.append([
                r["equipo"],
                fmt_dt(r["inicio_amt"]),
                fmt_dt(r["termino_amt"]),
                fmt_horas(r["duracion_amt_h"]),
                fmt_dt(r["inicio_collahuasi"]),
                fmt_dt(r["termino_collahuasi"]),
                fmt_horas(r["duracion_collahuasi_h"]),
                fmt_horas(r["diferencia_h"]),
                Paragraph(str(r["errores"]), styles["Texto"]),
            ])
        elementos.append(tabla_pdf(data, col_widths=[1.5*cm, 2.7*cm, 2.7*cm, 1.3*cm, 2.7*cm, 2.7*cm, 1.3*cm, 1.3*cm, 8.0*cm], font_size=6))

    # Revision especial In Progress.
    elementos.append(PageBreak())
    elementos.append(Paragraph("Revisión especial eventos AMT In Progress", styles["Subtitulo"]))
    if df_in_progress is None or df_in_progress.empty:
        elementos.append(Paragraph("No se detectaron eventos AMT In Progress con los criterios configurados.", styles["Texto"]))
    else:
        data = [[
            "Resultado", "Equipo", "Inicio AMT", "Término AMT ref.",
            "Término Coll. detectado", "H Coll.", "Dif. H", "Filas Coll.", "Detalle"
        ]]
        for _, r in df_in_progress.head(80).iterrows():
            data.append([
                r["resultado"],
                r["equipo"],
                fmt_dt(r["inicio_amt"]),
                fmt_dt(r["termino_amt_referencial"]),
                fmt_dt(r["termino_collahuasi_detectado"]),
                fmt_horas(r["duracion_collahuasi_detectada_h"]),
                fmt_horas(r["diferencia_vs_referencial_h"]),
                Paragraph(str(r["filas_collahuasi"]), styles["Texto"]),
                Paragraph(str(r["detalle"]), styles["Texto"]),
            ])
        elementos.append(tabla_pdf(data, col_widths=[1.7*cm, 1.5*cm, 2.7*cm, 2.7*cm, 2.9*cm, 1.3*cm, 1.3*cm, 2.3*cm, 8.0*cm], font_size=6))

    # Informativos AMT sin Collahuasi.
    elementos.append(PageBreak())
    elementos.append(Paragraph("Eventos AMT no encontrados en Collahuasi - Informativo", styles["Subtitulo"]))
    if df_amt_sin_coll.empty:
        elementos.append(Paragraph("No hay eventos AMT sin registros Collahuasi asociados.", styles["Texto"]))
    else:
        elementos.append(Paragraph(
            "Estos registros no se consideran error principal, porque la lógica de control parte desde Detenciones Collahuasi hacia DailyDowntimeLog/AMT.",
            styles["Texto"]
        ))
        data = [["Equipo", "Inicio AMT", "Término AMT", "H AMT", "Descripción"]]
        for _, r in df_amt_sin_coll.head(80).iterrows():
            data.append([
                r["equipo"],
                fmt_dt(r["inicio_amt"]),
                fmt_dt(r["termino_amt"]),
                fmt_horas(r["duracion_amt_h"]),
                Paragraph(str(r.get("descripcion", "")), styles["Texto"]),
            ])
        elementos.append(tabla_pdf(data, col_widths=[1.6*cm, 3*cm, 3*cm, 1.5*cm, 15*cm], font_size=6))

    doc.build(elementos, onFirstPage=agregar_pie_pagina, onLaterPages=agregar_pie_pagina)
    buffer.seek(0)
    return buffer.getvalue()

# ============================================================
# EXPORT EXCEL
# ============================================================

def generar_excel_resultados(
    df_resultado: pd.DataFrame,
    df_eventos_error: pd.DataFrame,
    df_amt_sin_coll: pd.DataFrame,
    df_in_progress: pd.DataFrame,
    df_eventos_amt: pd.DataFrame,
    df_asignaciones_amt: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resultado.to_excel(writer, index=False, sheet_name="Revision registros")
        df_eventos_error.to_excel(writer, index=False, sheet_name="Errores eventos")
        df_amt_sin_coll.to_excel(writer, index=False, sheet_name="AMT informativo")
        df_in_progress.to_excel(writer, index=False, sheet_name="In Progress")
        df_eventos_amt.to_excel(writer, index=False, sheet_name="Eventos AMT")
        df_asignaciones_amt.to_excel(writer, index=False, sheet_name="Asignaciones AMT")
    output.seek(0)
    return output.getvalue()

# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

def main():
    st.title("🛠️ Revisión de Detenciones Collahuasi vs DailyDowntimeLog / AMT")

    st.markdown("Carga ambos archivos Excel y presiona **Comparar detenciones** para generar el informe.")

    col1, col2 = st.columns(2)
    with col1:
        archivo_daily = st.file_uploader("1. Cargar DailyDowntimeLog.xlsx", type=["xlsx"])
    with col2:
        archivo_collahuasi = st.file_uploader("2. Cargar DETENCIONES COLLAHUASI 2026.xlsx", type=["xlsx"])

    st.sidebar.header("Configuración")
    tolerancia_minutos = st.sidebar.number_input(
        "Tolerancia en minutos",
        min_value=0,
        max_value=60,
        value=3,
        step=1,
        help="Tolerancia permitida para diferencias pequeñas por segundos/redondeos."
    )
    filtrar_por_rango_daily = st.sidebar.checkbox(
        "Validar solo registros Collahuasi dentro del rango DailyDowntimeLog",
        value=True,
    )
    validar_continuidad = st.sidebar.checkbox(
        "Validar gaps/solapamientos entre cortes Collahuasi",
        value=True,
    )
    validar_cobertura_total = st.sidebar.checkbox(
        "Validar cobertura total del evento AMT",
        value=False,
        help="Si se activa, marca error cuando Collahuasi no cubre todo el inicio/término/duración del evento AMT. Por defecto queda desactivado para respetar la lógica inversa solicitada."
    )

    st.sidebar.subheader("In Progress")
    detectar_in_progress_por_0800 = st.sidebar.checkbox(
        "Tratar cualquier término 08:00 como posible In Progress",
        value=False,
        help="Por defecto se considera In Progress cuando el término AMT coincide con el fin del rango descargado del DailyDowntimeLog, aunque sea 00:00, 08:00 u otra hora. Activa esta opción si tus reportes también usan 08:00 como término referencial en cualquier día."
    )
    ventana_in_progress_horas = st.sidebar.number_input(
        "Ventana búsqueda término Collahuasi (horas)",
        min_value=8,
        max_value=720,
        value=168,
        step=8,
        help="Cantidad de horas posteriores al término referencial AMT para buscar continuidad en Collahuasi."
    )
    max_gap_in_progress_horas = st.sidebar.number_input(
        "Gap máximo entre cortes In Progress (horas)",
        min_value=0.0,
        max_value=24.0,
        value=2.0,
        step=0.5,
        help="Permite unir cortes Collahuasi posteriores al término 08:00 cuando pertenecen a la misma detención."
    )

    if archivo_daily is None or archivo_collahuasi is None:
        st.warning("Carga ambos archivos para iniciar la revisión.")
        return

    if st.button("Comparar detenciones", type="primary"):
        with st.spinner("Leyendo archivos y comparando registros..."):
            try:
                df_eventos_amt, df_asignaciones_amt, rango_inicio, rango_termino = leer_daily_downtime_log(
                    archivo_daily,
                    detectar_in_progress_por_0800=detectar_in_progress_por_0800,
                )
                df_collahuasi = leer_detenciones_collahuasi(archivo_collahuasi)

                df_resultado, df_eventos_error, df_amt_sin_coll = comparar_detenciones(
                    df_collahuasi=df_collahuasi,
                    df_eventos_amt=df_eventos_amt,
                    rango_inicio=rango_inicio,
                    rango_termino=rango_termino,
                    tolerancia_minutos=int(tolerancia_minutos),
                    filtrar_por_rango_daily=filtrar_por_rango_daily,
                    validar_cobertura_total=validar_cobertura_total,
                    validar_continuidad=validar_continuidad,
                )

                df_in_progress = revisar_eventos_in_progress(
                    df_eventos_amt=df_eventos_amt,
                    df_collahuasi=df_collahuasi,
                    rango_termino=rango_termino,
                    tolerancia_minutos=int(tolerancia_minutos),
                    ventana_busqueda_horas=int(ventana_in_progress_horas),
                    max_gap_horas=float(max_gap_in_progress_horas),
                )

                pdf_bytes = generar_pdf(
                    df_resultado=df_resultado,
                    df_eventos_error=df_eventos_error,
                    df_amt_sin_coll=df_amt_sin_coll,
                    df_in_progress=df_in_progress,
                    rango_inicio=rango_inicio,
                    rango_termino=rango_termino,
                    tolerancia_minutos=int(tolerancia_minutos),
                )

                excel_bytes = generar_excel_resultados(
                    df_resultado=df_resultado,
                    df_eventos_error=df_eventos_error,
                    df_amt_sin_coll=df_amt_sin_coll,
                    df_in_progress=df_in_progress,
                    df_eventos_amt=df_eventos_amt,
                    df_asignaciones_amt=df_asignaciones_amt,
                )

            except Exception as e:
                st.error(f"No fue posible procesar los archivos: {e}")
                st.exception(e)
                return

        st.success("Revisión finalizada.")

        rango_txt = f"{fmt_dt(rango_inicio)} a {fmt_dt(rango_termino)}" if rango_inicio and rango_termino else "No detectado"
        total = len(df_resultado)
        errores = len(df_resultado[df_resultado["resultado"] == "ERROR"]) if not df_resultado.empty else 0
        correctos = len(df_resultado[df_resultado["resultado"] == "OK"]) if not df_resultado.empty else 0

        total_in_progress = len(df_in_progress) if not df_in_progress.empty else 0
        errores_in_progress = len(df_in_progress[df_in_progress["resultado"] == "ERROR"]) if total_in_progress else 0

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Rango Daily", rango_txt)
        c2.metric("Registros revisados", total)
        c3.metric("Correctos", correctos)
        c4.metric("Con error", errores)
        c5.metric("In Progress", f"{total_in_progress} / {errores_in_progress} alerta")

        st.subheader("Errores encontrados")
        if df_resultado.empty or errores == 0:
            st.success("No se detectaron errores a nivel registro Collahuasi.")
        else:
            columnas_vista = [
                "fila_collahuasi", "equipo", "inicio_collahuasi", "termino_collahuasi",
                "inicio_amt", "termino_amt", "duracion_collahuasi_h", "duracion_amt_h",
                "errores", "observaciones", "razon_collahuasi", "comentario_collahuasi", "descripcion_amt"
            ]
            st.dataframe(
                df_resultado[df_resultado["resultado"] == "ERROR"][columnas_vista],
                use_container_width=True,
                hide_index=True,
            )

        st.subheader("Errores de cortes / continuidad por evento")
        if df_eventos_error.empty:
            st.info("No se detectaron errores agrupados por evento.")
        else:
            st.dataframe(df_eventos_error, use_container_width=True, hide_index=True)

        st.subheader("Revisión especial In Progress")
        if df_in_progress.empty:
            st.info("No se detectaron eventos In Progress con la configuración actual.")
        else:
            st.dataframe(df_in_progress, use_container_width=True, hide_index=True)

        st.subheader("Eventos AMT no encontrados en Collahuasi - Informativo")
        st.caption("Estos registros no se consideran error principal según la lógica solicitada.")
        if df_amt_sin_coll.empty:
            st.info("No hay eventos AMT sin registros Collahuasi asociados.")
        else:
            st.dataframe(df_amt_sin_coll, use_container_width=True, hide_index=True)

        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "📄 Descargar informe PDF",
                data=pdf_bytes,
                file_name="informe_revision_detenciones.pdf",
                mime="application/pdf",
            )
        with d2:
            st.download_button(
                "📊 Descargar detalle Excel",
                data=excel_bytes,
                file_name="detalle_revision_detenciones.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with st.expander("Ver todos los registros revisados"):
            st.dataframe(df_resultado, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()

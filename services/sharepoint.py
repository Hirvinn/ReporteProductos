import io
from pathlib import Path

import pandas as pd
import requests


# ============================================================
# CONFIGURACIÓN
# ============================================================

NOMBRE_ARCHIVO = "ITEM MASTER ORIGINAL.xlsx"

COLUMNAS_REQUERIDAS = [
    "COD_ITEM",
    "DESCRIPCION",
    "MACROCATEGORIA",
    "CATEGORIA",
    "SUBCATEGORIA",
    "FECHA_CREACION"
]


# ============================================================
# DESCARGAR EXCEL DESDE URL
# ============================================================

def descargar_excel_desde_url(url: str) -> bytes:

    if not url:
        raise ValueError(
            "No se ha configurado la URL del archivo."
        )

    respuesta = requests.get(
        url,
        timeout=60
    )

    respuesta.raise_for_status()

    return respuesta.content


# ============================================================
# LEER EXCEL
# ============================================================

def leer_excel_desde_bytes(contenido: bytes) -> pd.DataFrame:

    archivo = io.BytesIO(contenido)

    df = pd.read_excel(
        archivo,
        engine="openpyxl"
    )

    return df


# ============================================================
# VALIDAR COLUMNAS
# ============================================================

def validar_columnas(df: pd.DataFrame):

    # Limpiar espacios de los nombres
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    faltantes = [
        columna
        for columna in COLUMNAS_REQUERIDAS
        if columna not in df.columns
    ]

    if faltantes:

        raise ValueError(
            "El Excel no contiene las siguientes columnas: "
            + ", ".join(faltantes)
        )

    return True


# ============================================================
# PREPARAR DATA
# ============================================================

def preparar_datos(df: pd.DataFrame) -> pd.DataFrame:

    validar_columnas(df)

    df = df[COLUMNAS_REQUERIDAS].copy()

    # --------------------------------------------------------
    # Campos de texto
    # --------------------------------------------------------

    campos_texto = [
        "COD_ITEM",
        "DESCRIPCION",
        "MACROCATEGORIA",
        "CATEGORIA",
        "SUBCATEGORIA"
    ]

    for columna in campos_texto:

        df[columna] = (
            df[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # --------------------------------------------------------
    # Fecha
    # --------------------------------------------------------

    df["FECHA_CREACION"] = pd.to_datetime(
        df["FECHA_CREACION"],
        errors="coerce"
    )

    return df


# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================

def cargar_excel_sharepoint(url: str) -> pd.DataFrame:

    contenido = descargar_excel_desde_url(url)

    df = leer_excel_desde_bytes(
        contenido
    )

    df = preparar_datos(
        df
    )

    return df
from pathlib import Path

import pandas as pd


# ============================================================
# CONFIGURACIÓN
# ============================================================

RUTA_EXCEL = Path(
    r"C:\Users\hmbelalcazars\OneDrive - CorporacionGPF - Femsa Salud"
    r"\Documentos\ITEM MASTER ORIGINAL - Copia.xlsx"
)


COLUMNAS_REQUERIDAS = [
    "COD_ITEM",
    "DESCRIPCION",
    "MACROCATEGORIA",
    "CATEGORIA",
    "SUBCATEGORIA",
    "FECHA_CREACION"
]


# ============================================================
# VERIFICAR ARCHIVO
# ============================================================

def verificar_archivo():

    if not RUTA_EXCEL.exists():

        raise FileNotFoundError(
            "No se encontró el archivo Excel:\n"
            f"{RUTA_EXCEL}"
        )

    return True


# ============================================================
# LEER EXCEL
# ============================================================

def cargar_excel():

    verificar_archivo()

    df = pd.read_excel(
        RUTA_EXCEL,
        engine="openpyxl"
    )

    return df


# ============================================================
# VALIDAR COLUMNAS
# ============================================================

def validar_columnas(df):

    # Limpiar nombres de columnas
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
            "El Excel no contiene las siguientes columnas:\n\n"
            + "\n".join(
                f"- {columna}"
                for columna in faltantes
            )
        )

    return True


# ============================================================
# PREPARAR DATA
# ============================================================

def preparar_datos(df):

    validar_columnas(df)

    # Seleccionar únicamente las columnas necesarias
    df = df[COLUMNAS_REQUERIDAS].copy()

    # --------------------------------------------------------
    # CAMPOS DE TEXTO
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
    # FECHA
    # --------------------------------------------------------

    df["FECHA_CREACION"] = pd.to_datetime(
        df["FECHA_CREACION"],
        errors="coerce"
    )

    # --------------------------------------------------------
    # CAMPOS AUXILIARES PARA FILTROS
    # --------------------------------------------------------

    df["ANIO"] = df["FECHA_CREACION"].dt.year

    df["MES"] = df["FECHA_CREACION"].dt.month

    df["DIA"] = df["FECHA_CREACION"].dt.day

    return df


# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================

def obtener_productos():

    df = cargar_excel()

    df = preparar_datos(df)

    return df
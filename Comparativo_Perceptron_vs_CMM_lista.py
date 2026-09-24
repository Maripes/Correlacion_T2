import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import io
from collections import OrderedDict
import re
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font
from scipy import stats
 
# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Convertir TXT Perceptron a Excel", layout="wide")
 
# --- ESTILO GLOBAL (FONDO OSCURO, TABLAS CLARAS) ---
st.markdown("""
    <style>
    body {
        background-color: #121212;
        color: #FFFFFF;
        font-family: 'Poppins', sans-serif;
    }
 
    .stApp {
        background-color: #121212;
    }
 
    /* Encabezados */
    h1, h2, h3, h4 {
        color: #ffc107;
    }
 
    /* Área de subida de archivos */
    div[data-testid="stFileUploader"] {
        border: 2px dashed #5a5a5a !important;
        background-color: rgba(50,50,50,0.7);
        border-radius: 15px;
        padding: 20px;
    }
 
    div[data-testid="stFileUploader"]:hover {
        border-color: #ffc107 !important;
        background-color: rgba(80,80,80,0.9);
    }
 
    /* Tabla de correlación */
    .dataframe {
        background: #2b2b2b !important;
        color: #ffffff !important;
        border-radius: 10px;
        font-size: 15px;
    }
 
    .dataframe td, .dataframe th {
        text-align: center !important;
        padding: 8px !important;
    }
 
    /* Botón de descarga */
    div.stDownloadButton > button {
        background-color: #ffc107;
        color: #000;
        font-weight: bold;
        border-radius: 10px;
        border: none;
        padding: 10px 25px;
    }
 
    div.stDownloadButton > button:hover {
        background-color: #ffde59;
        color: #000;
    }
    </style>
""", unsafe_allow_html=True)
 
# --- TÍTULO ---
st.title("📄 Comparativo PU T2")
 
# --- FUNCIONES PARA PROCESAR ARCHIVOS ---

def procesar_perceptron_txt(archivo):
    contenido = archivo.read().decode("latin-1").splitlines()
    encabezados, mediciones = [], []
    encabezado_encontrado = False

    for linea in contenido:
        partes = linea.strip().split("\t")

        if "JSN" in partes and "PSN" in partes:
            encabezados = partes
            encabezado_encontrado = True
            continue

        # Solo leer mediciones después del encabezado real.
        if not encabezado_encontrado:
            continue

        if (
            partes
            and len(partes) >= 2
            and partes[0].strip().upper()
            not in ["NOMINAL", "USL", "LSL", "UTL", "LTL", "URL", "LRL"]
        ):
            mediciones.append(partes)

    if not encabezados or not mediciones:
        return None, []

    filas_med = []
    for med in mediciones:
        fila = OrderedDict({
            "JSN": med[0],
            "PSN": med[1] if len(med) > 1 else "",
            "Fecha": med[2] if len(med) > 2 else "",
            "Hora": med[3] if len(med) > 3 else ""
        })
        for i, col in enumerate(encabezados[4:], start=4):
            fila[col] = med[i] if i < len(med) else ""
        filas_med.append(fila)

    eje_cols = encabezados[4:]
    return pd.DataFrame(filas_med), eje_cols


def procesar_cmm_txt(archivo):
    """
    Lee archivos CMM con estructura:

        DIM 3000L PQC-15
        AX MEAS NOMINAL +TOL -TOL DEV OUTTOL
        X ...
        Y ...
        Z ...

    Guarda el DEV de cada eje como:

        3000L PQC-15[X]
        3000L PQC-15[Y]
        3000L PQC-15[Z]
    """

    contenido = archivo.read().decode("latin-1").splitlines()

    # -----------------------------------------
    # OBTENER JSN
    # -----------------------------------------
    jsn = ""

    for linea in contenido:

        m = re.search(
            r"TRACEFIELD\s+JSN\s*=\s*(\S+)",
            linea,
            re.IGNORECASE
        )

        if m:
            jsn = m.group(1).strip()
            break

    # -----------------------------------------
    # MEDICIONES
    # -----------------------------------------
    mediciones = OrderedDict()

    dim_actual = None
    leyendo_axis = False

    for linea in contenido:

        linea_limpia = linea.strip()

        # -------------------------------------
        # DIM
        # -------------------------------------
        m_dim = re.match(
            r"^\*?DIM\s+(.+?)\s*\*?$",
            linea_limpia,
            re.IGNORECASE
        )

        if m_dim:

            dim_actual = m_dim.group(1).strip()

            # El CMM puede traer:
            # 3000L PQC-15&#x20;
            # quitamos basura HTML si aparece

            dim_actual = (
                dim_actual
                .replace("&#x20;", "")
                .replace("&amp;", "&")
                .strip()
            )

            leyendo_axis = False

            continue

        # -------------------------------------
        # ENCABEZADO AX
        # -------------------------------------
        if re.match(
            r"^\*?AX\s+MEAS\s+NOMINAL",
            linea_limpia,
            re.IGNORECASE
        ):

            if dim_actual is not None:
                leyendo_axis = True

            continue

        # -------------------------------------
        # SI TODAVÍA NO ESTAMOS EN AX
        # -------------------------------------
        if not leyendo_axis or dim_actual is None:
            continue

        # -------------------------------------
        # X / Y / Z / M
        # -------------------------------------
        m_axis = re.match(
            r"^([XYZM])\s+(.+)$",
            linea_limpia,
            re.IGNORECASE
        )

        if m_axis:

            eje = m_axis.group(1).upper()

            resto = m_axis.group(2).strip()

            valores = resto.split()

            # Necesitamos:
            #
            # NOMINAL
            # +TOL
            # -TOL
            # DEV
            # OUTTOL
            #
            # por eso mínimo 5 valores después del eje

            if len(valores) >= 5:

                try:

                    # DEV = posición 4
                    dev = float(valores[4])

                except ValueError:
                    continue

                nombre_cmm = f"{dim_actual}[{eje}]"

                mediciones[nombre_cmm] = dev

            continue

        # -------------------------------------
        # FIN DEL BLOQUE AX
        # -------------------------------------
        if (
            linea_limpia.startswith("*POINTDATA")
            or linea_limpia.startswith("POINTDATA")
            or linea_limpia.startswith("*DIM")
            or linea_limpia.startswith("DIM ")
        ):
            leyendo_axis = False

    # -----------------------------------------
    # VALIDACIÓN
    # -----------------------------------------
    if not jsn:
        return None, []

    if not mediciones:
        return None, []

    # -----------------------------------------
    # CREAR FILA
    # -----------------------------------------
    fila = OrderedDict({

        "JSN": jsn,

        "PSN": jsn,

        "Fecha": "",

        "Hora": ""

    })

    fila.update(mediciones)

    return (
        pd.DataFrame([fila]),
        list(mediciones.keys())
    )


def procesar_archivo(archivo, tipo):
    if tipo == "perceptron":
        return procesar_perceptron_txt(archivo)
    return procesar_cmm_txt(archivo)

# --- MAPEO DE EJES ---
FORCED_MAP = {
    "3000L PQC-15[X]": "3000L[X]",
    "3000L PQC-15[Y]": "3000L[Y]",
    "3000L PQC-15[Z]": "3000L[Z]",
    "3000R PQC-24[X]": "3000R[X]",
    "3000R PQC-24[Y]": "3000R[Y]",
    "3000R PQC-24[Z]": "3000R[Z]",
    "3003L PQC-15[X]": "3003L[X]",
    "3003L PQC-15[Y]": "3003L[Y]",
    "3003L PQC-15[Z]": "3003L[Z]",
    "3003R PQC-24[X]": "3003R[X]",
    "3003R PQC-24[Y]": "3003R[Y]",
    "3003R PQC-24[Z]": "3003R[Z]",
    "3004L PQC-18[X]": "3004L[X]",
    "3004L PQC-18[Y]": "3004L[Y]",
    "3004L PQC-18[Z]": "3004L[Z]",
    "3004R PQC-18[X]": "3004R[X]",
    "3004R PQC-18[Y]": "3004R[Y]",
    "3004R PQC-18[Z]": "3004R[Z]",
    "3005L PQC-18[X]": "3005L[X]",
    "3005L PQC-18[Y]": "3005L[Y]",
    "3005L PQC-18[Z]": "3005L[Z]",
    "3005R PQC-18[X]": "3005R[X]",
    "3005R PQC-18[Y]": "3005R[Y]",
    "3005R PQC-18[Z]": "3005R[Z]",
    "3006L PQC-18[X]": "3006L[X]",
    "3006L PQC-18[Y]": "3006L[Y]",
    "3006L PQC-18[Z]": "3006L[Z]",
    "3006R PQC-18[X]": "3006R[X]",
    "3006R PQC-18[Y]": "3006R[Y]",
    "3006R PQC-18[Z]": "3006R[Z]",
    "3007L PQC-18[X]": "3007L[X]",
    "3007L PQC-18[Y]": "3007L[Y]",
    "3007L PQC-18[Z]": "3007L[Z]",
    "3007R PQC-18[X]": "3007R[X]",
    "3007R PQC-18[Y]": "3007R[Y]",
    "3007R PQC-18[Z]": "3007R[Z]",
    "3084AL PQC 41[X]": "3084AL[X]",
    "3084AL PQC 41[Y]": "3084AL[Y]",
    "3084AL PQC 41[Z]": "3084AL[Z]",
    "3084BL PQC-41[X]": "3084BL[X]",
    "3084BL PQC-41[Y]": "3084BL[Y]",
    "3084BL PQC-41[Z]": "3084BL[Z]",
    "3124L PQC-9[X]": "3124L_P12[X]",
    "3124L PQC-9[Y]": "3124L_P12[Y]",
    "3124L PQC-9[Z]": "3124L_P12[Z]",
    "3124R PQC-9[X]": "3124R_P9[X]",
    "3124R PQC-9[Y]": "3124R_P9[Y]",
    "3124R PQC-9[Z]": "3124R_P9[Z]",
    "3125L PQC-9[X]": "3125L_P12[X]",
    "3125L PQC-9[Y]": "3125L_P12[Y]",
    "3125L PQC-9[Z]": "3125L_P12[Z]",
    "3125R PQC-9[X]": "3125R_P9[X]",
    "3125R PQC-9[Y]": "3125R_P9[Y]",
    "3125R PQC-9[Z]": "3125R_P9[Z]",
    "3126L PQC-9[X]": "3126L_P12[X]",
    "3126L PQC-9[Y]": "3126L_P12[Y]",
    "3126L PQC-9[Z]": "3126L_P12[Z]",
    "3126R PQC-9[X]": "3126R_P9[X]",
    "3126R PQC-9[Y]": "3126R_P9[Y]",
    "3126R PQC-9[Z]": "3126R_P9[Z]",
    "5001L[X]": "5001L[X]",
    "5001L[Y]": "5001L[Y]",
    "5001L[Z]": "5001L[Z]",
    "5001R[X]": "5001R[X]",
    "5001R[Y]": "5001R[Y]",
    "5001R[Z]": "5001R[Z]",
    "7002L PQC-3[X]": "7002L_P1_3[X]",
    "7002L PQC-3[Y]": "7002L_P1_3[Y]",
    "7002L PQC-3[Z]": "7002L_P1_3[Z]",
    "7002R PQC-3[X]": "7002R_P1_3[X]",
    "7002R PQC-3[Y]": "7002R_P1_3[Y]",
    "7002R PQC-3[Z]": "7002R_P1_3[Z]",
    "7005L PQC-3[X]": "7005L_P1_3[X]",
    "7005L PQC-3[Y]": "7005L_P1_3[Y]",
    "7005L PQC-3[Z]": "7005L_P1_3[Z]",
    "7005R PQC-3[X]": "7005R_P1_3[X]",
    "7005R PQC-3[Y]": "7005R_P1_3[Y]",
    "7005R PQC-3[Z]": "7005R_P1_3[Z]",
    "7006L PQC-3[X]": "7006L_P1_3[X]",
    "7006L PQC-3[Y]": "7006L_P1_3[Y]",
    "7006L PQC-3[Z]": "7006L_P1_3[Z]",
    "7006R PQC-3[X]": "7006R_P1_3[X]",
    "7006R PQC-3[Y]": "7006R_P1_3[Y]",
    "7006R PQC-3[Z]": "7006R_P1_3[Z]",
    "7007L PQC-3[X]": "7007L_743_543_P1_3[X]",
    "7007L PQC-3[Y]": "7007L_743_543_P1_3[Y]",
    "7007L PQC-3[Z]": "7007L_743_543_P1_3[Z]",
    "7007R PQC-3[X]": "7007R_743_543_P1_3[X]",
    "7007R PQC-3[Y]": "7007R_743_543_P1_3[Y]",
    "7007R PQC-3[Z]": "7007R_743_543_P1_3[Z]",
    "7012L PQC-3[X]": "7012L[X]",
    "7012L PQC-3[Y]": "7012L[Y]",
    "7012L PQC-3[Z]": "7012L[Z]",
    "7013L PQC-3[X]": "7013L[X]",
    "7013L PQC-3[Y]": "7013L[Y]",
    "7013L PQC-3[Z]": "7013L[Z]",
    "7022L[X]": "7022L_543[X]",
    "7022L[Y]": "7022L_543[Y]",
    "7022L[Z]": "7022L_543[Z]",
    "7022R[X]": "7022R_543[X]",
    "7022R[Y]": "7022R_543[Y]",
    "7022R[Z]": "7022R_543[Z]",
    "7026L[X]": "7026L[X]",
    "7026L[Y]": "7026L[Y]",
    "7026L[Z]": "7026L[Z]",
    "7026R[X]": "7026R[X]",
    "7026R[Y]": "7026R[Y]",
    "7026R[Z]": "7026R[Z]",
    "7027L[X]": "7027L[X]",
    "7027L[Y]": "7027L[Y]",
    "7027L[Z]": "7027L[Z]",
    "7027R[X]": "7027R[X]",
    "7027R[Y]": "7027R[Y]",
    "7027R[Z]": "7027R[Z]",
    "7028L[X]": "7028L[X]",
    "7028L[Y]": "7028L[Y]",
    "7028L[Z]": "7028L[Z]",
    "7028R[X]": "7028R[X]",
    "7028R[Y]": "7028R[Y]",
    "7028R[Z]": "7028R[Z]",
    "7029L[X]": "7029L[X]",
    "7029L[Y]": "7029L[Y]",
    "7029L[Z]": "7029L[Z]",
    "7029R[X]": "7029R[X]",
    "7029R[Y]": "7029R[Y]",
    "7029R[Z]": "7029R[Z]",
    "7035L[X]": "7035L[X]",
    "7035L[Y]": "7035L[Y]",
    "7035L[Z]": "7035L[Z]",
    "7200L PQC-3[X]": "7200L[X]",
    "7200L PQC-3[Y]": "7200L[Y]",
    "7200L PQC-3[Z]": "7200L[Z]",
    "7200R PQC-3[X]": "7200R[X]",
    "7200R PQC-3[Y]": "7200R[Y]",
    "7200R PQC-3[Z]": "7200R[Z]",
    "7201L PQC-3[X]": "7201L[X]",
    "7201L PQC-3[Y]": "7201L[Y]",
    "7201L PQC-3[Z]": "7201L[Z]",
    "7201R PQC-3[X]": "7201R[X]",
    "7201R PQC-3[Y]": "7201R[Y]",
    "7201R PQC-3[Z]": "7201R[Z]",
    "7202L PQC-3[X]": "7202L[X]",
    "7202L PQC-3[Y]": "7202L[Y]",
    "7202L PQC-3[Z]": "7202L[Z]",
    "7202R PQC-3[X]": "7202R[X]",
    "7202R PQC-3[Y]": "7202R[Y]",
    "7202R PQC-3[Z]": "7202R[Z]",
    "7203L PQC-3[X]": "7203L[X]",
    "7203L PQC-3[Y]": "7203L[Y]",
    "7203L PQC-3[Z]": "7203L[Z]",
    "7203R PQC-3[X]": "7203R[X]",
    "7203R PQC-3[Y]": "7203R[Y]",
    "7203R PQC-3[Z]": "7203R[Z]",
    "7204L PQC-3[X]": "7204L[X]",
    "7204L PQC-3[Y]": "7204L[Y]",
    "7204L PQC-3[Z]": "7204L[Z]",
    "7204R PQC-3[X]": "7204R[X]",
    "7204R PQC-3[Y]": "7204R[Y]",
    "7204R PQC-3[Z]": "7204R[Z]",
    "7205L PQC-3[X]": "7205L[X]",
    "7205L PQC-3[Y]": "7205L[Y]",
    "7205L PQC-3[Z]": "7205L[Z]",
    "7205R PQC-3[X]": "7205R[X]",
    "7205R PQC-3[Y]": "7205R[Y]",
    "7205R PQC-3[Z]": "7205R[Z]",
    "7206L PQC-3[X]": "7206L_743_543[X]",
    "7206L PQC-3[Y]": "7206L_743_543[Y]",
    "7206L PQC-3[Z]": "7206L_743_543[Z]",
    "7206R PQC-3[X]": "7206R_743_543[X]",
    "7206R PQC-3[Y]": "7206R_743_543[Y]",
    "7206R PQC-3[Z]": "7206R_743_543[Z]",
    "7207L PQC-3[X]": "7207L_743_543[X]",
    "7207L PQC-3[Y]": "7207L_743_543[Y]",
    "7207L PQC-3[Z]": "7207L_743_543[Z]",
    "7207R PQC-3[X]": "7207R_743_543[X]",
    "7207R PQC-3[Y]": "7207R_743_543[Y]",
    "7207R PQC-3[Z]": "7207R_743_543[Z]",
    "7208L PQC-3[X]": "7208L[X]",
    "7208L PQC-3[Y]": "7208L[Y]",
    "7208L PQC-3[Z]": "7208L[Z]",
    "7208R PQC-3[X]": "7208R[X]",
    "7208R PQC-3[Y]": "7208R[Y]",
    "7208R PQC-3[Z]": "7208R[Z]",
    "9906XLD[X]": "9906XLD[X]",
    "9906XLD[Y]": "9906XLD[Y]",
    "9906XLD[Z]": "9906XLD[Z]",
    "9906XRD PQC-3[X]": "9906XRD[X]",
    "9906XRD PQC-3[Y]": "9906XRD[Y]",
    "9906XRD PQC-3[Z]": "9906XRD[Z]",
    "9907XLD PQC-3[X]": "9907XLD[X]",
    "9907XLD PQC-3[Y]": "9907XLD[Y]",
    "9907XLD PQC-3[Z]": "9907XLD[Z]",
    "9907XRD PQC-3[X]": "9907XRD[X]",
    "9907XRD PQC-3[Y]": "9907XRD[Y]",
    "9907XRD PQC-3[Z]": "9907XRD[Z]",
    "9908XLD PQC-3[X]": "9908XLD[X]",
    "9908XLD PQC-3[Y]": "9908XLD[Y]",
    "9908XLD PQC-3[Z]": "9908XLD[Z]",
    "9908XRD PQC-3[X]": "9908XRD[X]",
    "9908XRD PQC-3[Y]": "9908XRD[Y]",
    "9908XRD PQC-3[Z]": "9908XRD[Z]"
}


def map_axis(perceptron_axis):

    if perceptron_axis in FORCED_MAP:
        return FORCED_MAP[perceptron_axis]

    match = re.match(
        r"^(1100)([LR]\[[XYZ]\])$",
        perceptron_axis,
        re.IGNORECASE
    )

    if match:
        return f"3125{match.group(2)}"

    return perceptron_axis

# ============================================================
# ESTADISTICAS TIPO PERCEPTRON
# ============================================================

def calcular_estadisticas_perceptron(perceptron_vals, cmm_vals):

    perceptron_vals = np.asarray(perceptron_vals, dtype=float)
    cmm_vals = np.asarray(cmm_vals, dtype=float)

    # Eliminar NaN / infinitos
    mask = (
        np.isfinite(perceptron_vals) &
        np.isfinite(cmm_vals)
    )

    perceptron_vals = perceptron_vals[mask]
    cmm_vals = cmm_vals[mask]

    n = len(perceptron_vals)

    if n == 0:
        return None

    # --------------------------------------------------------
    # DIFFERENCE
    # CMM - PIERCE
    # --------------------------------------------------------

    difference = cmm_vals - perceptron_vals

    # --------------------------------------------------------
    # ADJUSTMENT
    # Promedio de las diferencias
    # --------------------------------------------------------

    adjustment = np.mean(difference)

    # --------------------------------------------------------
    # ADJUSTED PIERCE
    # PIERCE + adjustment
    # --------------------------------------------------------

    adjusted_perceptron = perceptron_vals + adjustment

    # --------------------------------------------------------
    # ADJUSTED DIFFERENCE
    # --------------------------------------------------------

    adjusted_difference = cmm_vals - adjusted_perceptron

    # --------------------------------------------------------
    # FUNCION PARA ESTADISTICAS
    # --------------------------------------------------------

    def estadisticas(valores):

        valores = np.asarray(valores, dtype=float)

        if len(valores) == 0:
            return {
                "Mean": np.nan,
                "6 Sigma": np.nan,
                "Minimum": np.nan,
                "Maximum": np.nan,
                "Range": np.nan
            }

        # STDEV.S -> ddof=1
        if len(valores) >= 2:
            std = np.std(valores, ddof=1)
        else:
            std = np.nan

        return {
            "Mean": np.mean(valores),
            "6 Sigma": std * 6 if not np.isnan(std) else np.nan,
            "Minimum": np.min(valores),
            "Maximum": np.max(valores),
            "Range": np.max(valores) - np.min(valores)
        }

    stats_perceptron = estadisticas(perceptron_vals)
    stats_cmm = estadisticas(cmm_vals)
    stats_difference = estadisticas(difference)
    stats_adjusted = estadisticas(adjusted_perceptron)
    stats_adjusted_difference = estadisticas(adjusted_difference)

    # --------------------------------------------------------
    # CORRELACION
    # --------------------------------------------------------

    if (
        n >= 2
        and np.std(perceptron_vals, ddof=1) > 0
        and np.std(cmm_vals, ddof=1) > 0
    ):
        correlation = np.corrcoef(
            perceptron_vals,
            cmm_vals
        )[0, 1]
    else:
        correlation = np.nan

    # --------------------------------------------------------
    # T-TEST PAREADO
    #
    # PIERCE vs CMM
    # --------------------------------------------------------

    if n >= 2:

        try:
            t_test = stats.ttest_rel(
                perceptron_vals,
                cmm_vals
            ).pvalue
        except Exception:
            t_test = np.nan

    else:
        t_test = np.nan

    # --------------------------------------------------------
    # F-TEST
    #
    # Compara las varianzas PIERCE vs CMM
    # --------------------------------------------------------

    if n >= 2:

        try:

            var_p = np.var(
                perceptron_vals,
                ddof=1
            )

            var_c = np.var(
                cmm_vals,
                ddof=1
            )

            if var_p > 0 and var_c > 0:

                # Ponemos la mayor varianza arriba
                # para que F >= 1

                if var_p >= var_c:
                    f_value = var_p / var_c
                else:
                    f_value = var_c / var_p

                df1 = n - 1
                df2 = n - 1

                # p-value bilateral
                p_one_tail = 1 - stats.f.cdf(
                    f_value,
                    df1,
                    df2
                )

                f_test = min(
                    1.0,
                    2 * p_one_tail
                )

            else:
                f_test = np.nan

        except Exception:
            f_test = np.nan

    else:
        f_test = np.nan

    return {
        "n": n,

        "PIERCE": stats_perceptron,
        "CMM": stats_cmm,
        "Difference": stats_difference,
        "Adjusted PIERCE": stats_adjusted,
        "Adjusted Difference": stats_adjusted_difference,

        "Adjustment": adjustment,
        "Correlation": correlation,
        "T-Test": t_test,
        "F-Test": f_test
    }

# --- SUBIDA DE ARCHIVOS ---
st.subheader("📤 Archivos PERCEPTRON")
archivos_perceptron = st.file_uploader(
    "Carga hasta 9 archivos TXT Perceptron",
    type=["txt"],
    accept_multiple_files=True,
    key="perceptron"
)

st.subheader("📤 Archivos CMM")
archivos_cmm = st.file_uploader(
    "Carga hasta 9 archivos TXT CMM",
    type=["txt"],
    accept_multiple_files=True,
    key="cmm"
)

# --- VALIDACIÓN DE CANTIDAD ---
if len(archivos_perceptron) > 9:
    st.error("⚠️ Puedes cargar máximo 9 archivos PERCEPTRON.")
    st.stop()

if len(archivos_cmm) > 9:
    st.error("⚠️ Puedes cargar máximo 9 archivos CMM.")
    st.stop()

# --- PROCESAMIENTO ---
if archivos_perceptron and archivos_cmm:

    # Procesar todos los archivos PERCEPTRON
    perceptron_dfs = []
    for archivo in archivos_perceptron:
        df_tmp, _ = procesar_archivo(archivo, "perceptron")
        if df_tmp is not None and not df_tmp.empty:
            perceptron_dfs.append(df_tmp)

    # Procesar todos los archivos CMM
    cmm_dfs = []
    for archivo in archivos_cmm:
        df_tmp, _ = procesar_archivo(archivo, "cmm")
        if df_tmp is not None and not df_tmp.empty:
            cmm_dfs.append(df_tmp)

    if not perceptron_dfs:
        st.error("⚠️ Ningún archivo PERCEPTRON contiene mediciones válidas.")
        st.stop()

    if not cmm_dfs:
        st.error("⚠️ Ningún archivo CMM contiene mediciones válidas.")
        st.stop()

    df_perceptron = pd.concat(perceptron_dfs, ignore_index=True, sort=False)
    df_cmm = pd.concat(cmm_dfs, ignore_index=True, sort=False)

    st.success(
        f"✅ Procesados {len(perceptron_dfs)} archivos PERCEPTRON y "
        f"{len(cmm_dfs)} archivos CMM."
    )

    # ============================================================
    # MATCH: PERCEPTRON JSN == CMM TRACEFIELD JSN
    # El PSN NO se utiliza para relacionar los archivos.
    # ============================================================
    df_perceptron["JSN"] = df_perceptron["JSN"].astype(str).str.strip()
    df_cmm["JSN"] = df_cmm["JSN"].astype(str).str.strip()

    jsn_validos = sorted(
        set(df_perceptron["JSN"]).intersection(set(df_cmm["JSN"]))
    )

    if not jsn_validos:
        st.error(
            "⚠️ No se encontró coincidencia de JSN entre Perceptron y CMM. "
            "El CMM se identifica mediante TRACEFIELD JSN."
        )
        st.stop()

    df_match = pd.DataFrame({"JSN": jsn_validos})

    df_perceptron = df_perceptron[
        df_perceptron["JSN"].isin(jsn_validos)
    ].reset_index(drop=True)

    df_cmm = df_cmm[
        df_cmm["JSN"].isin(jsn_validos)
    ].reset_index(drop=True)

    # ---------------------------------
    # EXTRAER STATION Y MODEL DINÁMICO
    # ---------------------------------
    nombre_archivo = archivos_perceptron[0].name.replace(".txt", "")

    if "_" in nombre_archivo:
        nombre_sin_fecha = nombre_archivo.split("_", 1)[1]
    else:
        nombre_sin_fecha = nombre_archivo

    partes = nombre_sin_fecha.split("_")

    if "Front" in partes:
        idx_front = partes.index("Front")
        station_name = "_".join(partes[:idx_front + 2])
        model_name = (
            "_".join(partes[idx_front + 2:])
            if len(partes) > idx_front + 2
            else "UNKNOWN"
        )
    else:
        station_name = nombre_sin_fecha
        model_name = "UNKNOWN"

    # ---------------------------------
    # MAPEOS: SOLO LOS PUNTOS DE LA LISTA
    # ---------------------------------
    #
    # La lista FORCED_MAP es la lista maestra de puntos Perceptron.
    # Para cada punto:
    #   1) se intenta usar el CMM indicado por FORCED_MAP;
    #   2) si ese nombre no existe en el CMM nuevo, se busca el
    #      mismo punto como DIM del reporte CMM.
    #
    # Esto permite usar el CMM real con formato:
    #   DIM 3000L PQC-15
    #   X <MEAS>
    #   Y <MEAS>
    #   Z <MEAS>
    # ============================================================
    # MAPEO PERCEPTRON -> CMM
    # ============================================================

    cmm_columnas = list(df_cmm.columns)
    perceptron_columnas = list(df_perceptron.columns)

    ejes_mapeados = []

    for cmm_eje, perceptron_eje in FORCED_MAP.items():

        # El punto debe existir en PERCEPTRON
        if perceptron_eje not in perceptron_columnas:
            continue

        # Buscar nombre EXACTO en CMM
        if cmm_eje in cmm_columnas:

            ejes_mapeados.append(
                (
                    perceptron_eje,
                    cmm_eje
                )
            )

        else:

            # ----------------------------------------------------
            # Si no existe exactamente, intentar por checkpoint
            # ----------------------------------------------------

            match = re.match(
                r"^(\d+[A-Z]+)\[([XYZ])\]$",
                perceptron_eje.strip(),
                re.IGNORECASE
            )

            if match:

                checkpoint = match.group(1).upper()
                axis = match.group(2).upper()

                patron = re.compile(
                    rf"^{re.escape(checkpoint)}\s+.*\[{axis}\]$",
                    re.IGNORECASE
                )

                candidatos = [
                    columna
                    for columna in cmm_columnas
                    if patron.match(str(columna).strip())
                ]

                if candidatos:

                    ejes_mapeados.append(
                        (
                            perceptron_eje,
                            candidatos[0]
                        )
                    )
   


    # ============================================================
    # CREAR TABLA DE MAPEOS
    # ============================================================

    df_axes = pd.DataFrame(
        ejes_mapeados,
        columns=[
            "Perceptron-Axis",
            "CMM-Axis"
        ]
    ).drop_duplicates()


    # ============================================================
    # VALIDACIÓN
    # ============================================================

    if df_axes.empty:

        st.error(
            "⚠️ No se encontró ningún punto de la lista en el CMM."
        )

        st.write("### Columnas encontradas en PERCEPTRON")
        st.write(list(df_perceptron.columns))

        st.write("### Columnas encontradas en CMM")
        st.write(list(df_cmm.columns))

        st.stop()


    # ============================================================
    # MOSTRAR MAPEOS EN PANTALLA PARA DEPURAR
    # ============================================================

    st.subheader("🔗 Mapeo Perceptron → CMM")

    st.dataframe(
        df_axes,
        use_container_width=True,
        hide_index=True
    )
 

    df_axes = pd.DataFrame(
        ejes_mapeados,
        columns=["Perceptron-Axis", "CMM-Axis"]
    ).drop_duplicates()

    if df_axes.empty:
        st.error(
            "⚠️ No se encontró ningún punto de la lista en el CMM. "
            "Revisa que los nombres de DIM del CMM correspondan a la lista."
        )
        st.stop()

    # Mantener SOLO los puntos de la lista que sí existen en ambos archivos.
    puntos_perceptron = df_axes["Perceptron-Axis"].tolist()
    puntos_cmm = df_axes["CMM-Axis"].tolist()

    columnas_perceptron = ["PSN", "JSN", "Fecha", "Hora"] + [
        p for p in puntos_perceptron
        if p in df_perceptron.columns
    ]

    columnas_cmm = ["PSN", "JSN", "Fecha", "Hora"] + [
        p for p in puntos_cmm
        if p in df_cmm.columns
    ]

    # Quitar duplicados de nombres base antes de seleccionar.
    columnas_perceptron = list(dict.fromkeys(columnas_perceptron))
    columnas_cmm = list(dict.fromkeys(columnas_cmm))

    df_perceptron = df_perceptron[columnas_perceptron].copy()
    df_cmm = df_cmm[columnas_cmm].copy()

    # --- Cálculo de correlaciones ---

    df_merge = pd.merge(
        df_perceptron, df_cmm,
        on="JSN", suffixes=("_perceptron", "_cmm")
    )

    correlacion_data = []

    for perceptron_eje, cmm_eje in df_axes.values:

        # Intentar primero columnas sin sufijo
        col_perceptron_1 = perceptron_eje
        col_cmm_1 = cmm_eje

        # Intentar columnas con sufijo del merge
        col_perceptron_2 = f"{perceptron_eje}_perceptron"
        col_cmm_2 = f"{cmm_eje}_cmm"

        # Determinar cuáles columnas existen en df_merge
        if col_perceptron_1 in df_merge.columns and col_cmm_1 in df_merge.columns:
            col_perceptron = col_perceptron_1
            col_cmm = col_cmm_1
        elif col_perceptron_2 in df_merge.columns and col_cmm_2 in df_merge.columns:
            col_perceptron = col_perceptron_2
            col_cmm = col_cmm_2
        else:
            # Si no existe ninguna forma, saltar eje
            continue

        tmp = df_merge[[col_perceptron, col_cmm]].copy()

        tmp[col_perceptron] = pd.to_numeric(tmp[col_perceptron], errors="coerce")
        tmp[col_cmm] = pd.to_numeric(tmp[col_cmm], errors="coerce")

        tmp = tmp.dropna()

        if len(tmp) < 1:
            continue

        perceptron_vals = tmp[col_perceptron].values
        cmm_vals = tmp[col_cmm].values

        perceptron_mean = np.mean(perceptron_vals)
        cmm_mean = np.mean(cmm_vals)

        if len(tmp) < 2 or np.std(perceptron_vals) == 0 or np.std(cmm_vals) == 0:
            correlation = np.nan
        else:
            correlation = np.corrcoef(
                perceptron_vals, cmm_vals
            )[0, 1]
        #sigma6 = np.std(perceptron_vals) * 6
        sigma6 = np.std(perceptron_vals, ddof=1) * 6
        offset_calc = cmm_mean - perceptron_mean

        resultado_stats = calcular_estadisticas_perceptron(
            perceptron_vals,
            cmm_vals
        )

        correlacion_data.append([
            perceptron_eje,
            cmm_eje,
            round(perceptron_mean, 3),
            round(cmm_mean, 3),
            round(correlation, 3),
            round(sigma6, 3),
            round(offset_calc, 3),
            round(resultado_stats["T-Test"], 4)
                if not np.isnan(resultado_stats["T-Test"])
                else np.nan,
            round(resultado_stats["F-Test"], 4)
                if not np.isnan(resultado_stats["F-Test"])
                else np.nan
])

 
 
    df_correlacion = pd.DataFrame(
        correlacion_data,
        columns=[
        "Perceptron-Axis",
        "CMM-Axis",
        "Perceptron-Mean",
        "CMM-Mean",
        "Correlation",
        "6Sigma",
        "Calculated-Offset",
        "T-Test",
        "F-Test"
        ]
    )
 
    def colorear_correlacion(val):
        if isinstance(val, (int, float)):
            if val >= 0.7:
                return 'background-color: #47FF47; color: #000000; font-weight: 600;'
            elif val >= 0.69:
                return 'background-color: #FFFD00; color: #000000; font-weight: 600;'
        return 'color: #FFFFFF;'
 
    def colorear_offset(val):
        if isinstance(val, (int, float)):
            if abs(val) > 1:
                return 'background-color: #FF0000; color: #FFFFFF; font-weight: 600;'
            elif abs(val) > 0.5:
                return 'background-color: #FFFD00; color: #000000; font-weight: 600;'
        return 'color: #FFFFFF;'
    
    df_correlacion_styled = (
        df_correlacion.style
        .apply(lambda col: col.map(colorear_correlacion) if col.name == "Correlation" else [""]*len(col), axis=0)
        .apply(lambda col: col.map(colorear_offset) if col.name == "Calculated-Offset" else [""]*len(col), axis=0)
        .set_table_styles([
            {'selector': 'th', 'props': [('background-color', '#2b2b2b'),
                                        ('color', '#FFFFFF'),
                                        ('font-weight', 'bold'),
                                        ('text-align', 'center'),
                                        ('padding', '8px')]},
            {'selector': 'td', 'props': [('background-color', '#1e1e1e'),
                                        ('color', '#FFFFFF'),
                                        ('text-align', 'center'),
                                        ('padding', '8px')]},
            {'selector': 'tbody tr:hover', 'props': [('background-color', '#333333')]},
            {'selector': 'table', 'props': [('border-radius', '10px'),
                                        ('overflow', 'hidden'),
                                        ('border', '1px solid #444')]}
        ])
    )

    st.subheader("📈 Correlación")
    st.dataframe(df_correlacion_styled, use_container_width=True)

    # ============================================================
    # ESTADISTICAS DETALLADAS TIPO PERCEPTRON
    # ============================================================

    st.subheader("📊 Statistics")

    # Lista de puntos disponibles
    puntos_disponibles = df_axes["Perceptron-Axis"].tolist()

    if puntos_disponibles:

        punto_seleccionado = st.selectbox(
            "Selecciona el punto para ver sus estadísticas:",
            puntos_disponibles
        )

        # Buscar el CMM correspondiente
        fila_mapping = df_axes[
            df_axes["Perceptron-Axis"] == punto_seleccionado
        ]

        if not fila_mapping.empty:

            cmm_eje = fila_mapping.iloc[0]["CMM-Axis"]

            # ----------------------------------------------------
            # Buscar columnas en el merge
            # ----------------------------------------------------

            col_p = None
            col_c = None

            posibles_p = [
                punto_seleccionado,
                f"{punto_seleccionado}_perceptron"
            ]

            posibles_c = [
                cmm_eje,
                f"{cmm_eje}_cmm"
            ]

            for columna in posibles_p:
                if columna in df_merge.columns:
                    col_p = columna
                    break

            for columna in posibles_c:
                if columna in df_merge.columns:
                    col_c = columna
                    break

            if col_p and col_c:

                tmp_stats = df_merge[
                    [col_p, col_c]
                ].copy()

                tmp_stats[col_p] = pd.to_numeric(
                    tmp_stats[col_p],
                    errors="coerce"
                )

                tmp_stats[col_c] = pd.to_numeric(
                    tmp_stats[col_c],
                    errors="coerce"
                )

                tmp_stats = tmp_stats.dropna()

                perceptron_vals = tmp_stats[col_p].to_numpy()
                cmm_vals = tmp_stats[col_c].to_numpy()

                resultado_stats = calcular_estadisticas_perceptron(
                    perceptron_vals,
                    cmm_vals
                )

                if resultado_stats:

                    # ------------------------------------------------
                    # TABLA DE ESTADISTICAS
                    # ------------------------------------------------

                    df_stats = pd.DataFrame({

                        "Statistics": [
                            "Mean",
                            "6 Sigma",
                            "Minimum",
                            "Maximum",
                            "Range"
                        ],

                        "PIERCE": [
                            resultado_stats["PIERCE"]["Mean"],
                            resultado_stats["PIERCE"]["6 Sigma"],
                            resultado_stats["PIERCE"]["Minimum"],
                            resultado_stats["PIERCE"]["Maximum"],
                            resultado_stats["PIERCE"]["Range"]
                        ],

                        "CMM": [
                            resultado_stats["CMM"]["Mean"],
                            resultado_stats["CMM"]["6 Sigma"],
                            resultado_stats["CMM"]["Minimum"],
                            resultado_stats["CMM"]["Maximum"],
                            resultado_stats["CMM"]["Range"]
                        ],

                        "Difference": [
                            resultado_stats["Difference"]["Mean"],
                            resultado_stats["Difference"]["6 Sigma"],
                            resultado_stats["Difference"]["Minimum"],
                            resultado_stats["Difference"]["Maximum"],
                            resultado_stats["Difference"]["Range"]
                        ],

                        "Adjusted PIERCE": [
                            resultado_stats["Adjusted PIERCE"]["Mean"],
                            resultado_stats["Adjusted PIERCE"]["6 Sigma"],
                            resultado_stats["Adjusted PIERCE"]["Minimum"],
                            resultado_stats["Adjusted PIERCE"]["Maximum"],
                            resultado_stats["Adjusted PIERCE"]["Range"]
                        ],

                        "Adjusted Difference": [
                            resultado_stats["Adjusted Difference"]["Mean"],
                            resultado_stats["Adjusted Difference"]["6 Sigma"],
                            resultado_stats["Adjusted Difference"]["Minimum"],
                            resultado_stats["Adjusted Difference"]["Maximum"],
                            resultado_stats["Adjusted Difference"]["Range"]
                        ]
                    })

                    # ------------------------------------------------
                    # MOSTRAR TABLA
                    # ------------------------------------------------

                    st.dataframe(
                        df_stats.style.format(
                            {
                                "PIERCE": "{:.4f}",
                                "CMM": "{:.4f}",
                                "Difference": "{:.4f}",
                                "Adjusted PIERCE": "{:.4f}",
                                "Adjusted Difference": "{:.4f}"
                            }
                        ),
                        use_container_width=True,
                        hide_index=True
                    )

                    # ------------------------------------------------
                    # DATOS ESTADISTICOS
                    # ------------------------------------------------

                    col1, col2, col3 = st.columns(3)

                    with col1:

                        st.metric(
                            "Correlation Coefficient",
                            f"{resultado_stats['Correlation']:.4f}"
                            if not np.isnan(resultado_stats["Correlation"])
                            else "N/A"
                        )

                    with col2:

                        st.metric(
                            "T-Test",
                            f"{resultado_stats['T-Test']:.4f}"
                            if not np.isnan(resultado_stats["T-Test"])
                            else "N/A"
                        )

                    with col3:

                        st.metric(
                            "F-Test",
                            f"{resultado_stats['F-Test']:.4f}"
                            if not np.isnan(resultado_stats["F-Test"])
                            else "N/A"
                        )

                    # ------------------------------------------------
                    # ADJUSTMENT
                    # ------------------------------------------------

                    st.info(
                        f"**Adjustment:** "
                        f"{resultado_stats['Adjustment']:.4f}   |   "
                        f"**N:** {resultado_stats['n']} mediciones"
                    )

                    # ------------------------------------------------
                    # TABLA DE VALORES USADOS
                    # ------------------------------------------------

                    with st.expander(
                        "🔎 Ver valores utilizados para el cálculo"
                    ):

                        df_detalle = pd.DataFrame({

                            "PIERCE": perceptron_vals,

                            "CMM": cmm_vals,

                            "Difference":
                                cmm_vals - perceptron_vals,

                            "Adjusted PIERCE":
                                perceptron_vals +
                                resultado_stats["Adjustment"],

                            "Adjusted Difference":
                                cmm_vals -
                                (
                                    perceptron_vals +
                                    resultado_stats["Adjustment"]
                                )
                        })

                        st.dataframe(
                            df_detalle.style.format(
                                "{:.4f}"
                            ),
                            use_container_width=True,
                            hide_index=True
                        )
 
 
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_perceptron.to_excel(writer, index=False, sheet_name="Perceptron")
        df_cmm.to_excel(writer, index=False, sheet_name="CMM")
        df_match.to_excel(writer, index=False, sheet_name="Match_JSN")
        df_axes.to_excel(writer, index=False, sheet_name="Eje-Mapping")
        df_correlacion.to_excel(writer, index=False, sheet_name="Correlacion")

    buffer.seek(0)
    wb = openpyxl.load_workbook(buffer)
    ws = wb["Correlacion"]

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=5, max_col=5):
        for cell in row:
            if cell.value is not None:
                if cell.value >= 0.7:
                    cell.fill = PatternFill(start_color="47FF47", end_color="47FF47", fill_type="solid")
                elif cell.value >= 0.69:
                    cell.fill = PatternFill(start_color="FFFD00", end_color="FFFD00", fill_type="solid")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=7, max_col=7):
        for cell in row:
            if cell.value is not None:
                if abs(cell.value) > 1:
                    cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                elif abs(cell.value) > 0.5:
                    cell.fill = PatternFill(start_color="FFFD00", end_color="FFFD00", fill_type="solid")

    mean_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    for col in [3, 4]:
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=col, max_col=col):
            for cell in row:
                if cell.value is not None:
                    cell.fill = mean_fill

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
 
    st.download_button(
        label="📥 Descargar Excel completo coloreado",
        data=excel_buffer,
        file_name="Mediciones_Percepton_Completo_Coloreado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
 
    # ---------------- Preparar datos ----------------
    df_correlacion["Checkpoint"] = df_correlacion["Perceptron-Axis"].str.extract(r"(^\d+[LR])")
    df_correlacion["Axis"] = df_correlacion["Perceptron-Axis"].str.extract(r"\[([XYZ])\]")

    tabla_offsets = (
        df_correlacion[["Checkpoint", "Axis", "Calculated-Offset"]]
        .rename(columns={"Calculated-Offset": "OFFSET"})
    )

    tabla_offsets["OFFSET"] = tabla_offsets["OFFSET"].round(3)


    # ---------------- Tabla editable ----------------
    st.subheader("Editar OFFSET")

    tabla_editada = st.data_editor(
        tabla_offsets,
        column_config={
            "Checkpoint": st.column_config.TextColumn(disabled=True),
            "Axis": st.column_config.TextColumn(disabled=True),
            "OFFSET": st.column_config.NumberColumn(
                format="%.3f",
                step=0.001
            )
        },
        hide_index=True,
        use_container_width=True
    )


    # ---------------- Generar XML ----------------
    def generar_xml(tabla, station_name, model_name):

        gauge = ET.Element("GAUGE")
        station = ET.SubElement(gauge, "STATION")
        ET.SubElement(station, "NAME").text = station_name

        model = ET.SubElement(station, "MODEL")
        ET.SubElement(model, "NAME").text = model_name

        for checkpoint_name, group in tabla.groupby("Checkpoint"):

            checkpoint = ET.SubElement(model, "CHECKPOINT")
            ET.SubElement(checkpoint, "NAME").text = checkpoint_name

            for _, row in group.iterrows():

                axis_node = ET.SubElement(checkpoint, "AXIS")
                ET.SubElement(axis_node, "NAME").text = row["Axis"]
                ET.SubElement(axis_node, "OFFSET").text = str(row["OFFSET"])

            axis_node = ET.SubElement(checkpoint, "AXIS")
            ET.SubElement(axis_node, "NAME").text = "Diameter"
            ET.SubElement(axis_node, "OFFSET").text = "0"

        return ET.tostring(gauge, encoding="utf-8").decode("utf-8")


    xml_data = generar_xml(
        tabla_editada,
        station_name,
        model_name
    )


    # Vista previa
    st.code(xml_data, language="xml")


    # Descargar
    st.download_button(
        "📥 Descargar XML Editado",
        xml_data,
        "Comparacion_Perceptron_Editado.xml",
        "application/xml"
    )

import base64
import glob
import os
import re
import unicodedata

import pandas as pd
import plotly.express as px
import streamlit as st


# =========================================================
# CONFIGURACION GENERAL
# =========================================================

st.set_page_config(
    page_title="Estados de Pago | Contrato Mulchén",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

NOMBRE_BASE_EXCEL = "RESUMEN ESTADOS DE PAGO"

LOGO_SUPERIOR_BASE = "logo1"
SELLO_AGUA_BASE = "camion"
SELFIE_BASE = "selfie"

AUTOR = "Ricardo Grez"
CARGO = "Administrador de Contrato"
EMPRESA = "SAIVAM"
VERSION = "1.0"

CONTRATO = "CW2307646"
NOMBRE_CONTRATO = "Aseo Industrial y Gestión de Residuos"
CENTRO_OPERACIONAL = "CMPC Maderas | Planta Mulchén"

MESES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

MESES_NUMERO = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}

COLUMNAS_MONTO = [
    "Servicio_fijo",
    "Transporte_residuos",
    "Adicionales",
    "Total_neto",
    "IVA_19",
    "Total_bruto",
]

COLUMNAS_TONELADAS = [
    "Ton_RAD",
    "Ton_Corteza_G3",
    "Ton_Escoria",
    "Ton_Cenizas",
    "Ton_Total",
]

COLORES_RESIDUOS = {
    "RAD": "#9333ea",
    "Corteza G3": "#ea580c",
    "Escoria": "#059669",
    "Cenizas": "#2563eb",
}


# =========================================================
# UTILIDADES GENERALES
# =========================================================

def archivo_a_base64(ruta):
    if not ruta or not os.path.exists(ruta):
        return ""

    with open(ruta, "rb") as archivo:
        return base64.b64encode(
            archivo.read()
        ).decode("utf-8")


def obtener_tipo_mime(ruta):
    extension = os.path.splitext(str(ruta))[1].lower().replace(".", "")

    tipos = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "webp": "image/webp",
    }

    return tipos.get(extension, "image/png")


def buscar_imagen_por_nombre(nombre_base):
    extensiones = [
        "png",
        "jpg",
        "jpeg",
        "webp",
    ]

    nombre_base_normalizado = normalizar(nombre_base)

    candidatos = []

    for carpeta in ["."]:
        for extension in extensiones:
            candidatos.extend(glob.glob(os.path.join(carpeta, f"{nombre_base}.{extension}")))
            candidatos.extend(glob.glob(os.path.join(carpeta, f"**/*{nombre_base}*.{extension}"), recursive=True))

    for archivo in glob.glob("**/*", recursive=True):
        if os.path.isfile(archivo):
            extension = os.path.splitext(archivo)[1].lower().replace(".", "")

            if extension in extensiones:
                nombre_archivo = normalizar(os.path.splitext(os.path.basename(archivo))[0])

                if nombre_base_normalizado in nombre_archivo:
                    candidatos.append(archivo)

    candidatos_unicos = sorted(set(candidatos))

    if candidatos_unicos:
        return candidatos_unicos[0]

    return None


def buscar_archivo_excel():
    extensiones = [
        "xlsx",
        "xlsm",
        "xls",
    ]

    for extension in extensiones:
        ruta = f"{NOMBRE_BASE_EXCEL}.{extension}"

        if os.path.exists(ruta):
            return ruta

    archivos = []

    for extension in extensiones:
        archivos.extend(
            glob.glob(f"*{NOMBRE_BASE_EXCEL}*.{extension}")
        )

    if archivos:
        return sorted(archivos)[0]

    return None


def normalizar(texto):
    texto = unicodedata.normalize(
        "NFD",
        str(texto).strip().lower(),
    )

    texto = "".join(
        caracter
        for caracter in texto
        if unicodedata.category(caracter) != "Mn"
    )

    texto = texto.replace("_", " ")

    return re.sub(
        r"\s+",
        " ",
        texto,
    )


def buscar_columna(columnas, nombre):
    mapa = {
        normalizar(columna): columna
        for columna in columnas
    }

    return mapa.get(
        normalizar(nombre)
    )


def limpiar_numero(valor):
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor)

    texto = texto.replace("$", "")
    texto = texto.replace("ton", "")
    texto = texto.replace("t", "")
    texto = texto.replace(" ", "")

    texto = re.sub(
        r"[^0-9,.-]",
        "",
        texto,
    )

    if texto in ["", "-"]:
        return 0.0

    if "," in texto:
        texto = texto.replace(".", "")
        texto = texto.replace(",", ".")

    elif "." in texto and len(texto.split(".")[-1]) == 3:
        texto = texto.replace(".", "")

    try:
        return float(texto)

    except ValueError:
        return 0.0


def formato_monto(valor):
    try:
        return f"{int(round(float(valor))):,}".replace(",", ".")

    except (TypeError, ValueError):
        return "0"


def pesos_html(valor):
    return (
        f'<span class="monto-clp">'
        f'<span class="simbolo-peso">$</span>&nbsp;{formato_monto(valor)}'
        f'</span>'
    )


def pesos_texto(valor):
    return f"$ {formato_monto(valor)}"


def toneladas(valor):
    try:
        return (
            f"{float(valor):,.1f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
            + " t"
        )

    except (TypeError, ValueError):
        return "0,0 t"


def porcentaje(valor):
    try:
        return f"{float(valor):.1f}%".replace(".", ",")

    except (TypeError, ValueError):
        return "0,0%"


def convertir_fecha_desde_mes_periodo(mes, periodo):
    mes_texto = normalizar(mes)
    mes_numero = MESES_NUMERO.get(mes_texto)

    if not mes_numero:
        return pd.NaT

    coincidencia_anio = re.search(
        r"(\d{4})",
        str(periodo),
    )

    if not coincidencia_anio:
        return pd.NaT

    anio = int(coincidencia_anio.group(1))

    return pd.Timestamp(
        year=anio,
        month=mes_numero,
        day=1,
    )


def nombre_periodo(fecha):
    if pd.isna(fecha):
        return ""

    return f"{MESES[fecha.month]} {fecha.year}"


def reiniciar_app():
    try:
        st.rerun()
    except Exception:
        st.experimental_rerun()


# =========================================================
# ESTILO GENERAL CLARO
# =========================================================

def aplicar_estilo_general():
    st.markdown(
        """
<style>

[data-testid="stSidebar"],
[data-testid="collapsedControl"],
header[data-testid="stHeader"],
#MainMenu,
footer {
    display: none !important;
}

html,
body,
[data-testid="stAppViewContainer"],
.stApp {
    background: #f4f7fb !important;
    color: #0f172a !important;
}

.block-container {
    padding-top: 0.7rem !important;
    padding-bottom: 1rem !important;
    position: relative !important;
    z-index: 5 !important;
}

/* LOGIN */
.login-contenedor {
    max-width: 760px;
    margin: 40px auto 0 auto;
    padding: 30px 34px 26px 34px;

    background:
        linear-gradient(
            135deg,
            rgba(255, 255, 255, 0.98),
            rgba(241, 245, 249, 0.96)
        );

    border: 1px solid rgba(148, 163, 184, 0.65);
    border-radius: 22px;

    box-shadow:
        0 14px 34px
        rgba(15, 23, 42, 0.14);

    text-align: center;
}

.login-foto {
    width: 158px;
    height: 158px;
    margin: 0 auto 14px auto;
    border-radius: 50%;
    border: 6px solid #f59e0b;
    background-position: center;
    background-size: cover;
    background-repeat: no-repeat;

    box-shadow:
        0 8px 20px
        rgba(15, 23, 42, 0.18);
}

.login-titulo {
    font-size: 40px;
    font-weight: 950;
    color: #0f172a;
    margin: 8px 0 6px 0;
}

.login-subtitulo {
    font-size: 17px;
    font-weight: 650;
    color: #334155;
    margin-bottom: 12px;
}

.login-pie {
    border-top: 1px solid #cbd5e1;
    margin-top: 26px;
    padding-top: 18px;
    color: #64748b;
    font-size: 14px;
    line-height: 1.65;
    text-align: center;
}

/* Inputs login */
div[data-baseweb="input"] > div {
    background: #ffffff !important;
    color: #0f172a !important;
    border: 1px solid #94a3b8 !important;
    border-radius: 10px !important;
}

div[data-baseweb="input"] input {
    color: #0f172a !important;
    font-weight: 600 !important;
}

/* Botón ingresar */
div[data-testid="stFormSubmitButton"] button {
    background: linear-gradient(135deg, #16a34a, #15803d) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 12px !important;
    height: 48px !important;
    font-weight: 900 !important;
    width: 100% !important;
    box-shadow: 0 8px 18px rgba(22, 163, 74, 0.28) !important;
}

div[data-testid="stFormSubmitButton"] button:hover {
    background: linear-gradient(135deg, #15803d, #166534) !important;
    color: #ffffff !important;
}

/* Encabezado profesional */
.encabezado-panel {
    background:
        linear-gradient(
            135deg,
            rgba(255, 255, 255, 0.98),
            rgba(226, 232, 240, 0.92)
        );

    border: 1px solid rgba(148, 163, 184, 0.60);
    border-left: 6px solid #16a34a;
    border-radius: 18px;
    padding: 24px 26px 22px 26px;

    box-shadow:
        0 12px 28px
        rgba(15, 23, 42, 0.12);

    margin-bottom: 20px;
}

.titulo {
    font-size: 40px;
    font-weight: 900;
    color: #0f172a;
    line-height: 1.15;
    margin-bottom: 8px;
}

.subtitulo {
    font-size: 17px;
    color: #1e293b;
    margin: 8px 0 14px;
    line-height: 1.45;
}

.metadata {
    color: #64748b;
    font-size: 13px;
    margin-top: 10px;
}

.badges-header {
    display: flex;
    flex-wrap: wrap;
    gap: 9px;
    margin-top: 13px;
}

.badge-header {
    background: rgba(255, 255, 255, 0.88);
    border: 1px solid rgba(37, 99, 235, 0.28);
    color: #1e40af;
    padding: 7px 11px;
    border-radius: 999px;
    font-size: 13px;
    font-weight: 700;
}

.badge-verde {
    border-color: rgba(22, 163, 74, 0.42);
    color: #166534;
}

.badge-amarillo {
    border-color: rgba(202, 138, 4, 0.42);
    color: #854d0e;
}

.seccion {
    font-size: 23px;
    font-weight: 850;
    color: #0f172a;
    margin: 32px 0 15px;
    border-left: 5px solid #16a34a;
    padding-left: 11px;
}

/* Tarjetas KPI - estilo ejecutivo */
.tarjeta {
    background:
        linear-gradient(
            180deg,
            rgba(255, 255, 255, 0.99),
            rgba(248, 250, 252, 0.98)
        );

    border: 1px solid rgba(148, 163, 184, 0.58);
    border-top: 4px solid #334155;
    border-radius: 15px;
    padding: 18px 12px 16px 12px;
    min-height: 116px;
    text-align: center;

    box-shadow:
        0 8px 18px
        rgba(15, 23, 42, 0.08);
}

.tarjeta-verde {
    border-top-color: #047857;
}

.tarjeta-azul {
    border-top-color: #1d4ed8;
}

.tarjeta-naranjo {
    border-top-color: #b45309;
}

.tarjeta-amarillo {
    border-top-color: #92400e;
}

.tarjeta-morado {
    border-top-color: #5b21b6;
}

.tarjeta-titulo {
    font-size: 12px;
    color: #475569;
    font-weight: 850;
    letter-spacing: 0.25px;
    text-transform: uppercase;
    min-height: 20px;
}

.tarjeta-valor {
    font-size: 23px;
    font-weight: 950;
    margin-top: 8px;
    direction: ltr !important;
    unicode-bidi: isolate !important;
    white-space: nowrap !important;
}

.monto-clp {
    direction: ltr !important;
    unicode-bidi: isolate !important;
    white-space: nowrap !important;
}

.simbolo-peso {
    display: inline-block;
}

.tarjeta-subtitulo {
    color: #64748b;
    font-size: 12px;
    margin-top: 7px;
    line-height: 1.35;
}

.verde {
    color: #047857;
}

.azul {
    color: #1d4ed8;
}

.naranjo {
    color: #b45309;
}

.amarillo {
    color: #92400e;
}

.morado {
    color: #5b21b6;
}

.resumen {
    background:
        linear-gradient(
            135deg,
            rgba(255, 255, 255, 0.98),
            rgba(241, 245, 249, 0.98)
        );

    border: 1px solid rgba(148, 163, 184, 0.65);
    border-left: 5px solid #16a34a;
    border-radius: 13px;
    padding: 16px 18px;
    color: #0f172a;
    line-height: 1.65;
    font-size: 14px;

    box-shadow:
        0 6px 18px
        rgba(15, 23, 42, 0.08);
}

.pie-pagina {
    border-top: 1px solid #cbd5e1;
    color: #64748b;
    font-size: 12px;
    line-height: 1.65;
    margin-top: 38px;
    padding: 18px 0 8px;
    text-align: center;
}

div[data-baseweb="select"] > div {
    background: #ffffff !important;
    color: #111827 !important;
    border-radius: 8px !important;
    border: 1px solid #cbd5e1 !important;
}

div[data-baseweb="select"] span,
div[data-baseweb="select"] input {
    color: #111827 !important;
}

h1,
h2,
h3,
h4,
p,
label {
    color: #0f172a;
}

</style>
        """,
        unsafe_allow_html=True,
    )


def agregar_sello_agua_panel():
    ruta_sello = buscar_imagen_por_nombre(SELLO_AGUA_BASE)

    if not ruta_sello:
        return

    sello = archivo_a_base64(ruta_sello)
    tipo_mime = obtener_tipo_mime(ruta_sello)

    st.markdown(
        (
            "<style>"

            ".marca-agua-saivam {"
            "position: fixed;"
            "inset: 0;"
            "width: 100vw;"
            "height: 100vh;"
            f"background-image: url('data:{tipo_mime};base64,{sello}');"
            "background-repeat: no-repeat;"
            "background-position: center center;"
            "background-size: cover;"
            "opacity: 0.055;"
            "z-index: 1;"
            "pointer-events: none;"
            "}"

            "</style>"

            "<div class='marca-agua-saivam'></div>"
        ),
        unsafe_allow_html=True,
    )


# =========================================================
# LOGIN CON STREAMLIT SECRETS
# Formato esperado:
#
# [usuarios]
# Ricardo = "clave"
# Alexis = "clave"
# Moises = "clave"
# Victor = "clave"
# =========================================================

def obtener_usuarios_autorizados():
    try:
        usuarios = dict(st.secrets["usuarios"])

        if not usuarios:
            st.error(
                "La sección [usuarios] existe, pero no tiene usuarios configurados."
            )
            st.stop()

        return usuarios

    except Exception:
        st.error(
            "No se encontraron usuarios en Streamlit Secrets. "
            "Configura las credenciales con este formato: "
            "[usuarios] Ricardo = \"clave\""
        )
        st.stop()


def credenciales_validas(usuario_ingresado, clave_ingresada):
    usuarios = obtener_usuarios_autorizados()

    usuario_ingresado_norm = normalizar(usuario_ingresado)

    for usuario_guardado, clave_guardada in usuarios.items():
        if normalizar(usuario_guardado) == usuario_ingresado_norm:
            return str(clave_ingresada) == str(clave_guardada), usuario_guardado

    return False, None


def mostrar_login():
    ruta_selfie = buscar_imagen_por_nombre(SELFIE_BASE)

    if ruta_selfie:
        selfie_base64 = archivo_a_base64(ruta_selfie)
        foto_css = f"background-image: url('data:image/png;base64,{selfie_base64}');"
    else:
        foto_css = "background: #e2e8f0;"

    st.markdown(
        (
            '<div class="login-contenedor">'
            f'<div class="login-foto" style="{foto_css}"></div>'
            '<div class="login-titulo">🔐 Acceso restringido</div>'
            '<div class="login-subtitulo">'
            'Ingresa tu usuario y contraseña para visualizar el panel.'
            '</div>'
            '</div>'
        ),
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1.35, 2.3, 1.35])

    with col2:
        with st.form("formulario_login"):
            usuario = st.text_input("Usuario")
            clave = st.text_input("Contraseña", type="password")
            ingresar = st.form_submit_button("Ingresar", use_container_width=True)

        if ingresar:
            usuario = usuario.strip()

            acceso, usuario_validado = credenciales_validas(usuario, clave)

            if acceso:
                st.session_state["autenticado"] = True
                st.session_state["usuario"] = usuario_validado
                reiniciar_app()
            else:
                st.error("Usuario o contraseña incorrectos.")

        st.markdown(
            (
                '<div class="login-pie">'
                f'<b>Panel desarrollado por {AUTOR}</b><br>'
                f'{CARGO} | {EMPRESA}<br>'
                'Acceso restringido para usuarios autorizados'
                '</div>'
            ),
            unsafe_allow_html=True,
        )


def validar_acceso():
    if "autenticado" not in st.session_state:
        st.session_state["autenticado"] = False

    if "usuario" not in st.session_state:
        st.session_state["usuario"] = ""

    if not st.session_state["autenticado"]:
        mostrar_login()
        st.stop()


# =========================================================
# CARGA Y PREPARACION DE DATOS
# =========================================================

@st.cache_data
def cargar_datos(ruta_excel):
    datos = pd.read_excel(ruta_excel)

    datos = (
        datos
        .dropna(how="all")
        .dropna(axis=1, how="all")
    )

    datos.columns = (
        datos.columns
        .astype(str)
        .str.strip()
    )

    datos = datos.loc[
        :,
        ~datos.columns.str.contains(
            "Unnamed",
            case=False,
        ),
    ]

    renombres = {}

    # Se aceptan distintos nombres de encabezado para evitar errores
    # cuando la planilla usa títulos más descriptivos.
    equivalencias = {
        "Mes": [
            "Mes",
        ],
        "Periodo": [
            "Periodo",
            "Período",
        ],
        "Servicio_fijo": [
            "Servicio_fijo",
            "Servicio fijo",
        ],
        "Transporte_residuos": [
            "Transporte_residuos",
            "Transporte residuos",
            "Transporte de residuos",
        ],
        "Adicionales": [
            "Adicionales",
        ],
        "Total_neto": [
            "Total_neto",
            "Total neto",
            "Total neto acumulado",
        ],
        "IVA_19": [
            "IVA_19",
            "IVA 19",
            "IVA 19%",
        ],
        "Total_bruto": [
            "Total_bruto",
            "Total bruto",
            "Total bruto acumulado",
            "Valor total acumulado",
            "Total acumulado",
            "Monto final estados de pago",
        ],
        "Ton_RAD": [
            "Ton_RAD",
            "Ton RAD",
        ],
        "Ton_Corteza_G3": [
            "Ton_Corteza_G3",
            "Ton Corteza G3",
        ],
        "Ton_Escoria": [
            "Ton_Escoria",
            "Ton Escoria",
        ],
        "Ton_Cenizas": [
            "Ton_Cenizas",
            "Ton Cenizas",
        ],
        "Ton_Total": [
            "Ton_Total",
            "Ton Total",
            "Toneladas totales",
        ],
    }

    for columna_objetivo, nombres_posibles in equivalencias.items():
        for nombre_posible in nombres_posibles:
            columna_encontrada = buscar_columna(
                datos.columns,
                nombre_posible,
            )

            if columna_encontrada:
                renombres[columna_encontrada] = columna_objetivo
                break

    datos = datos.rename(columns=renombres)

    # Si la planilla no contiene una columna de total bruto,
    # se calcula automáticamente como total neto + IVA.
    if (
        "Total_bruto" not in datos.columns
        and "Total_neto" in datos.columns
        and "IVA_19" in datos.columns
    ):
        datos["Total_bruto"] = (
            datos["Total_neto"].apply(limpiar_numero)
            + datos["IVA_19"].apply(limpiar_numero)
        )

    columnas_requeridas = [
        "Mes",
        "Periodo",
        "Servicio_fijo",
        "Transporte_residuos",
        "Adicionales",
        "Total_neto",
        "IVA_19",
        "Total_bruto",
        "Ton_RAD",
        "Ton_Corteza_G3",
        "Ton_Escoria",
        "Ton_Cenizas",
        "Ton_Total",
    ]

    faltantes = [
        columna
        for columna in columnas_requeridas
        if columna not in datos.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas en la planilla: "
            + ", ".join(faltantes)
        )

    datos = datos[
        datos["Mes"].notna()
    ].copy()

    datos = datos[
        ~datos["Mes"]
        .astype(str)
        .str.lower()
        .str.contains("total")
    ].copy()

    for columna in COLUMNAS_MONTO:
        datos[columna] = datos[columna].apply(limpiar_numero)

    for columna in COLUMNAS_TONELADAS:
        datos[columna] = datos[columna].apply(limpiar_numero)

    datos["Fecha"] = datos.apply(
        lambda fila: convertir_fecha_desde_mes_periodo(
            fila["Mes"],
            fila["Periodo"],
        ),
        axis=1,
    )

    datos = datos[
        datos["Fecha"].notna()
    ].copy()

    datos["Año"] = datos["Fecha"].dt.year
    datos["Mes_Num"] = datos["Fecha"].dt.month
    datos["Mes_Nombre"] = datos["Mes_Num"].map(MESES)

    datos["Periodo_Texto"] = (
        datos["Mes_Nombre"]
        + " "
        + datos["Año"].astype(str)
    )

    datos["Periodo_Orden"] = datos["Fecha"].dt.strftime("%Y-%m")

    datos["Total_bruto_MM"] = datos["Total_bruto"] / 1_000_000
    datos["Total_neto_MM"] = datos["Total_neto"] / 1_000_000
    datos["Servicio_fijo_MM"] = datos["Servicio_fijo"] / 1_000_000
    datos["Transporte_residuos_MM"] = datos["Transporte_residuos"] / 1_000_000
    datos["Adicionales_MM"] = datos["Adicionales"] / 1_000_000

    datos["Costo_transporte_por_ton"] = datos.apply(
        lambda fila: (
            fila["Transporte_residuos"] / fila["Ton_Total"]
            if fila["Ton_Total"] > 0
            else 0
        ),
        axis=1,
    )

    datos["Participacion_transporte"] = datos.apply(
        lambda fila: (
            fila["Transporte_residuos"] / fila["Total_neto"] * 100
            if fila["Total_neto"] > 0
            else 0
        ),
        axis=1,
    )

    datos["Participacion_adicionales"] = datos.apply(
        lambda fila: (
            fila["Adicionales"] / fila["Total_neto"] * 100
            if fila["Total_neto"] > 0
            else 0
        ),
        axis=1,
    )

    datos = datos.sort_values("Fecha").reset_index(drop=True)

    return datos


def preparar_residuos_largo(datos):
    registros = []

    mapa_residuos = {
        "RAD": "Ton_RAD",
        "Corteza G3": "Ton_Corteza_G3",
        "Escoria": "Ton_Escoria",
        "Cenizas": "Ton_Cenizas",
    }

    for residuo, columna in mapa_residuos.items():
        temporal = datos[
            [
                "Fecha",
                "Año",
                "Mes_Num",
                "Mes_Nombre",
                "Periodo_Texto",
                "Periodo_Orden",
            ]
        ].copy()

        temporal["Residuo"] = residuo
        temporal["Toneladas"] = datos[columna]

        registros.append(temporal)

    residuos_largo = pd.concat(
        registros,
        ignore_index=True,
    )

    residuos_largo = residuos_largo[
        residuos_largo["Toneladas"] > 0
    ].copy()

    return residuos_largo


def preparar_componentes_ep_largo(datos):
    registros = []

    componentes = {
        "Servicio fijo": "Servicio_fijo_MM",
        "Transporte residuos": "Transporte_residuos_MM",
        "Adicionales": "Adicionales_MM",
    }

    for componente, columna in componentes.items():
        temporal = datos[
            [
                "Fecha",
                "Año",
                "Mes_Num",
                "Mes_Nombre",
                "Periodo_Texto",
                "Periodo_Orden",
            ]
        ].copy()

        temporal["Componente"] = componente
        temporal["Monto_MM"] = datos[columna]

        registros.append(temporal)

    componentes_largo = pd.concat(
        registros,
        ignore_index=True,
    )

    return componentes_largo


# =========================================================
# COMPONENTES VISUALES
# =========================================================

def seccion(titulo):
    st.markdown(
        f'<div class="seccion">{titulo}</div>',
        unsafe_allow_html=True,
    )


def tarjeta(titulo, valor, subtitulo, color):
    st.markdown(
        (
            f'<div class="tarjeta tarjeta-{color}">'
            f'<div class="tarjeta-titulo">{titulo}</div>'
            f'<div class="tarjeta-valor {color}">{valor}</div>'
            f'<div class="tarjeta-subtitulo">{subtitulo}</div>'
            "</div>"
        ),
        unsafe_allow_html=True,
    )


def formato_grafico(figura, altura=400):
    figura.update_layout(
        height=altura,
        plot_bgcolor="#f8fafc",
        paper_bgcolor="#f8fafc",

        font=dict(
            color="#0f172a",
            size=13,
        ),

        title_font=dict(
            size=20,
            color="#0f172a",
        ),

        legend=dict(
            font=dict(
                color="#0f172a",
                size=15,
            ),

            title_font=dict(
                color="#0f172a",
                size=14,
            ),

            bgcolor="rgba(255,255,255,0)",
        ),

        margin=dict(
            l=35,
            r=35,
            t=70,
            b=55,
        ),

        hoverlabel=dict(
            bgcolor="#ffffff",
            font_color="#0f172a",
        ),
    )

    figura.update_xaxes(
        title_text="",
        tickfont=dict(
            color="#334155",
            size=11,
        ),
        gridcolor="rgba(100,116,139,.22)",
    )

    figura.update_yaxes(
        title_text="",
        tickfont=dict(
            color="#334155",
            size=11,
        ),
        gridcolor="rgba(100,116,139,.22)",
    )

    return figura


def aplicar_eje_millones_clp(figura):
    figura.update_yaxes(
        title_text="",
        ticksuffix=" Millones CLP",
        separatethousands=True,
    )

    return figura


def aplicar_eje_toneladas(figura):
    figura.update_yaxes(
        title_text="",
        ticksuffix=" t",
        separatethousands=True,
    )

    return figura


def filtrar_datos(datos, filtro_anios, filtro_meses):
    salida = datos.copy()

    if filtro_anios:
        salida = salida[
            salida["Año"].isin(filtro_anios)
        ]

    if filtro_meses:
        salida = salida[
            salida["Mes_Nombre"].isin(filtro_meses)
        ]

    return salida.copy()


# =========================================================
# PANEL PRINCIPAL
# =========================================================

def mostrar_panel():
    ruta_excel = buscar_archivo_excel()

    if not ruta_excel:
        raise FileNotFoundError(
            "No se encontró el archivo Excel RESUMEN ESTADOS DE PAGO."
        )

    datos = cargar_datos(ruta_excel)

    fecha_minima = datos["Fecha"].min()
    fecha_maxima = datos["Fecha"].max()

    # -----------------------------------------------------
    # ENCABEZADO
    # -----------------------------------------------------

    col_titulo, col_logo = st.columns(
        [
            5,
            1,
        ]
    )

    with col_titulo:
        usuario_actual = st.session_state.get("usuario", "")

        st.markdown(
            (
                '<div class="encabezado-panel">'
                '<div class="titulo">'
                "📊 Resumen Ejecutivo de Estados de Pago"
                "</div>"

                '<div class="subtitulo">'
                f"<b>Contrato {CONTRATO}</b> | {NOMBRE_CONTRATO} | {CENTRO_OPERACIONAL}"
                "</div>"

                '<div class="badges-header">'
                f'<span class="badge-header badge-verde">Contrato {CONTRATO}</span>'
                '<span class="badge-header">Estados de Pago</span>'
                '<span class="badge-header">Transporte de Residuos</span>'
                '<span class="badge-header">Toneladas Gestionadas</span>'
                '<span class="badge-header badge-amarillo">Planta Mulchén</span>'
                f'<span class="badge-header">Usuario: {usuario_actual}</span>'
                "</div>"

                '<div class="metadata">'
                f"<b>Período analizado:</b> {nombre_periodo(fecha_minima)} a {nombre_periodo(fecha_maxima)}"
                " &nbsp;|&nbsp; "
                f"<b>Última actualización:</b> {nombre_periodo(fecha_maxima)}"
                " &nbsp;|&nbsp; "
                "<b>Fuente:</b> Resumen mensual de Estados de Pago"
                "</div>"
                "</div>"
            ),
            unsafe_allow_html=True,
        )

    with col_logo:
        ruta_logo = buscar_imagen_por_nombre(LOGO_SUPERIOR_BASE)

        if ruta_logo:
            st.image(
                ruta_logo,
                width=220,
            )

    # -----------------------------------------------------
    # FILTROS
    # -----------------------------------------------------

    seccion("🔎 Filtros de análisis")

    filtro1, filtro2 = st.columns(2)

    with filtro1:
        filtro_anios = st.multiselect(
            "Año",
            sorted(datos["Año"].unique()),
        )

    with filtro2:
        filtro_meses = st.multiselect(
            "Mes",
            [
                mes
                for mes in MESES.values()
                if mes in datos["Mes_Nombre"].unique()
            ],
        )

    datos_filtrados = filtrar_datos(
        datos,
        filtro_anios,
        filtro_meses,
    )

    if datos_filtrados.empty:
        st.warning("No existen datos para los filtros seleccionados.")
        return

    residuos_largo = preparar_residuos_largo(datos_filtrados)
    componentes_largo = preparar_componentes_ep_largo(datos_filtrados)

    total_servicio_fijo = float(datos_filtrados["Servicio_fijo"].sum())
    total_transporte = float(datos_filtrados["Transporte_residuos"].sum())
    total_adicionales = float(datos_filtrados["Adicionales"].sum())
    total_neto = float(datos_filtrados["Total_neto"].sum())
    total_iva = float(datos_filtrados["IVA_19"].sum())
    total_bruto = float(datos_filtrados["Total_bruto"].sum())
    total_toneladas = float(datos_filtrados["Ton_Total"].sum())
    total_rad = float(datos_filtrados["Ton_RAD"].sum())
    total_corteza = float(datos_filtrados["Ton_Corteza_G3"].sum())
    total_escoria = float(datos_filtrados["Ton_Escoria"].sum())
    total_cenizas = float(datos_filtrados["Ton_Cenizas"].sum())

    cantidad_meses = datos_filtrados["Periodo_Orden"].nunique()

    promedio_bruto_mensual = (
        total_bruto / cantidad_meses
        if cantidad_meses
        else 0
    )

    promedio_neto_mensual = (
        total_neto / cantidad_meses
        if cantidad_meses
        else 0
    )

    promedio_ton_mensual = (
        total_toneladas / cantidad_meses
        if cantidad_meses
        else 0
    )

    costo_transporte_por_ton = (
        total_transporte / total_toneladas
        if total_toneladas
        else 0
    )

    participacion_transporte = (
        total_transporte / total_neto * 100
        if total_neto
        else 0
    )

    participacion_adicionales = (
        total_adicionales / total_neto * 100
        if total_neto
        else 0
    )

    # -----------------------------------------------------
    # INDICADORES
    # -----------------------------------------------------

    seccion("📌 Indicadores ejecutivos")

    indicadores = [
        (
            "Total neto acumulado",
            pesos_html(total_neto),
            "Base antes de IVA",
            "verde",
        ),
        (
            "IVA 19%",
            pesos_html(total_iva),
            "Impuesto acumulado",
            "amarillo",
        ),
        (
            "Valor total acumulado",
            pesos_html(total_bruto),
            "Monto final estados de pago",
            "azul",
        ),
        (
            "Valor promedio mensual<br>(Neto sin IVA)",
            pesos_html(promedio_neto_mensual),
            f"{cantidad_meses} meses considerados",
            "morado",
        ),
        (
            "Servicio fijo",
            pesos_html(total_servicio_fijo),
            "Monto contractual acumulado",
            "verde",
        ),
        (
            "Transporte residuos",
            pesos_html(total_transporte),
            f"{porcentaje(participacion_transporte)} del neto",
            "naranjo",
        ),
        (
            "Adicionales",
            pesos_html(total_adicionales),
            f"{porcentaje(participacion_adicionales)} del neto",
            "amarillo",
        ),
        (
            "Toneladas gestionadas",
            toneladas(total_toneladas),
            "Total operacional del período",
            "azul",
        ),
    ]

    for inicio in range(0, 8, 4):
        columnas = st.columns(4)

        for columna, datos_tarjeta in zip(
            columnas,
            indicadores[inicio:inicio + 4],
        ):
            with columna:
                tarjeta(*datos_tarjeta)

    # -----------------------------------------------------
    # RESUMEN
    # -----------------------------------------------------

    seccion("📝 Resumen ejecutivo")

    st.markdown(
        (
            '<div class="resumen">'
            f"Entre <b>{nombre_periodo(fecha_minima)}</b> y "
            f"<b>{nombre_periodo(fecha_maxima)}</b>, el contrato registra "
            f"un total neto acumulado de <b>{pesos_html(total_neto)}</b>, "
            f"un IVA acumulado de <b>{pesos_html(total_iva)}</b> "
            f"y un total bruto de <b>{pesos_html(total_bruto)}</b>.<br><br>"

            f"El servicio fijo representa <b>{pesos_html(total_servicio_fijo)}</b>, "
            f"el transporte de residuos alcanza <b>{pesos_html(total_transporte)}</b> "
            f"y los adicionales acumulan <b>{pesos_html(total_adicionales)}</b>.<br><br>"

            f"Durante el período seleccionado se gestionaron "
            f"<b>{toneladas(total_toneladas)}</b>, con un costo promedio de transporte "
            f"de <b>{pesos_html(costo_transporte_por_ton)}</b> por tonelada.<br><br>"

            "<b>Hito relevante:</b> En <b>octubre 2023</b> se registra "
            "un <b>retroactivo por modificación de tarifas de traslado de residuos</b>, "
            "correspondiente al período informado, generando un aumento en los montos "
            "asociados a transporte de residuos."
            "</div>"
        ),
        unsafe_allow_html=True,
    )

    # -----------------------------------------------------
    # RESIDUOS
    # -----------------------------------------------------

    seccion("⚖️ Toneladas acumuladas por tipo de residuo")

    tarjetas_residuos = [
        (
            "RAD",
            toneladas(total_rad),
            "Residuo aserrín / derivados",
            "morado",
        ),
        (
            "Corteza G3",
            toneladas(total_corteza),
            "Corteza gestionada",
            "naranjo",
        ),
        (
            "Escoria",
            toneladas(total_escoria),
            "Escoria gestionada",
            "verde",
        ),
        (
            "Cenizas",
            toneladas(total_cenizas),
            "Cenizas gestionadas",
            "azul",
        ),
    ]

    columnas_residuos = st.columns(4)

    for columna, datos_tarjeta in zip(
        columnas_residuos,
        tarjetas_residuos,
    ):
        with columna:
            tarjeta(*datos_tarjeta)

    # -----------------------------------------------------
    # COSTO POR TONELADA
    # -----------------------------------------------------

    seccion("💰 Indicadores de costo por tonelada")

    col_costo1, col_costo2, col_costo3, col_costo4 = st.columns(4)

    with col_costo1:
        tarjeta(
            "Costo transporte por tonelada",
            pesos_html(costo_transporte_por_ton),
            "Transporte residuos / Ton_Total",
            "naranjo",
        )

    with col_costo2:
        tarjeta(
            "Transporte residuos",
            pesos_html(total_transporte),
            "Costo total asociado a transporte",
            "verde",
        )

    with col_costo3:
        tarjeta(
            "Toneladas promedio mensuales",
            toneladas(promedio_ton_mensual),
            f"{cantidad_meses} meses considerados",
            "azul",
        )

    with col_costo4:
        tarjeta(
            "Promedio mensual neto",
            pesos_html(promedio_neto_mensual),
            f"{cantidad_meses} meses considerados",
            "morado",
        )

    # -----------------------------------------------------
    # EVOLUCION ESTADOS DE PAGO
    # -----------------------------------------------------

    seccion("📈 Evolución mensual de Estados de Pago")

    figura_total_bruto = px.line(
        datos_filtrados,
        x="Periodo_Texto",
        y="Total_bruto_MM",
        markers=True,
        title="Evolución mensual del total bruto",
        template="plotly_white",
    )

    figura_total_bruto.update_traces(
        line=dict(
            width=3,
            color="#16a34a",
        ),
        marker=dict(
            size=6,
        ),
        customdata=datos_filtrados["Total_bruto"].apply(pesos_texto),
        hovertemplate=(
            "<b>%{x}</b>"
            "<br>Total bruto: %{customdata}"
            "<extra></extra>"
        ),
    )

    figura_total_bruto = formato_grafico(figura_total_bruto, 380)
    figura_total_bruto = aplicar_eje_millones_clp(figura_total_bruto)

    st.plotly_chart(
        figura_total_bruto,
        use_container_width=True,
    )

    figura_total_neto = px.line(
        datos_filtrados,
        x="Periodo_Texto",
        y="Total_neto_MM",
        markers=True,
        title="Evolución mensual del total neto",
        template="plotly_white",
    )

    figura_total_neto.update_traces(
        line=dict(
            width=3,
            color="#0284c7",
        ),
        marker=dict(
            size=6,
        ),
        customdata=datos_filtrados["Total_neto"].apply(pesos_texto),
        hovertemplate=(
            "<b>%{x}</b>"
            "<br>Total neto: %{customdata}"
            "<extra></extra>"
        ),
    )

    figura_total_neto = formato_grafico(figura_total_neto, 380)
    figura_total_neto = aplicar_eje_millones_clp(figura_total_neto)

    st.plotly_chart(
        figura_total_neto,
        use_container_width=True,
    )

    figura_adicionales = px.line(
        datos_filtrados,
        x="Periodo_Texto",
        y="Adicionales_MM",
        markers=True,
        title="Evolución mensual de adicionales",
        template="plotly_white",
    )

    figura_adicionales.update_traces(
        line=dict(
            width=3,
            color="#ca8a04",
        ),
        marker=dict(
            size=6,
        ),
        customdata=datos_filtrados["Adicionales"].apply(pesos_texto),
        hovertemplate=(
            "<b>%{x}</b>"
            "<br>Adicionales: %{customdata}"
            "<extra></extra>"
        ),
    )

    figura_adicionales = formato_grafico(figura_adicionales, 380)
    figura_adicionales = aplicar_eje_millones_clp(figura_adicionales)

    st.plotly_chart(
        figura_adicionales,
        use_container_width=True,
    )

    # -----------------------------------------------------
    # COMPOSICION ESTADO DE PAGO
    # -----------------------------------------------------

    seccion("📊 Composición mensual del Estado de Pago")

    figura_componentes = px.bar(
        componentes_largo,
        x="Periodo_Texto",
        y="Monto_MM",
        color="Componente",
        barmode="stack",
        title="Servicio fijo, transporte de residuos y adicionales",
        template="plotly_white",
    )

    figura_componentes.update_layout(
        colorway=[
            "#4f46e5",
            "#ea580c",
            "#059669",
        ]
    )

    figura_componentes = formato_grafico(figura_componentes, 420)
    figura_componentes = aplicar_eje_millones_clp(figura_componentes)

    st.plotly_chart(
        figura_componentes,
        use_container_width=True,
    )

    # -----------------------------------------------------
    # TONELADAS
    # -----------------------------------------------------

    seccion("⚖️ Evolución mensual de toneladas gestionadas")

    figura_ton_total = px.line(
        datos_filtrados,
        x="Periodo_Texto",
        y="Ton_Total",
        markers=True,
        title="Toneladas totales gestionadas por mes",
        template="plotly_white",
    )

    figura_ton_total.update_traces(
        line=dict(
            width=3,
            color="#0284c7",
        ),
        marker=dict(
            size=6,
        ),
        customdata=datos_filtrados["Ton_Total"].apply(toneladas),
        hovertemplate=(
            "<b>%{x}</b>"
            "<br>Toneladas: %{customdata}"
            "<extra></extra>"
        ),
    )

    figura_ton_total = formato_grafico(figura_ton_total, 380)
    figura_ton_total = aplicar_eje_toneladas(figura_ton_total)

    st.plotly_chart(
        figura_ton_total,
        use_container_width=True,
    )

    figura_residuos = px.bar(
        residuos_largo,
        x="Periodo_Texto",
        y="Toneladas",
        color="Residuo",
        barmode="stack",
        title="Composición mensual de toneladas por tipo de residuo",
        template="plotly_white",
        color_discrete_map=COLORES_RESIDUOS,
    )

    figura_residuos = formato_grafico(figura_residuos, 420)
    figura_residuos = aplicar_eje_toneladas(figura_residuos)

    st.plotly_chart(
        figura_residuos,
        use_container_width=True,
    )

    # -----------------------------------------------------
    # DISTRIBUCION
    # -----------------------------------------------------

    seccion("🥧 Distribución porcentual")

    col_pie1, col_pie2 = st.columns(2)

    with col_pie1:
        resumen_residuos = (
            residuos_largo
            .groupby("Residuo", as_index=False)["Toneladas"]
            .sum()
        )

        figura_pie_residuos = px.pie(
            resumen_residuos,
            names="Residuo",
            values="Toneladas",
            hole=0.56,
            title="Participación por tipo de residuo",
            template="plotly_white",
            color="Residuo",
            color_discrete_map=COLORES_RESIDUOS,
        )

        figura_pie_residuos.update_traces(
            textposition="inside",
            textinfo="percent+label",
        )

        st.plotly_chart(
            formato_grafico(figura_pie_residuos, 430),
            use_container_width=True,
        )

    with col_pie2:
        resumen_componentes = pd.DataFrame(
            {
                "Componente": [
                    "Servicio fijo",
                    "Transporte residuos",
                    "Adicionales",
                ],
                "Monto": [
                    total_servicio_fijo,
                    total_transporte,
                    total_adicionales,
                ],
            }
        )

        figura_pie_componentes = px.pie(
            resumen_componentes,
            names="Componente",
            values="Monto",
            hole=0.56,
            title="Participación sobre el total neto",
            template="plotly_white",
        )

        figura_pie_componentes.update_traces(
            textposition="inside",
            textinfo="percent+label",
        )

        st.plotly_chart(
            formato_grafico(figura_pie_componentes, 430),
            use_container_width=True,
        )

    # -----------------------------------------------------
    # COMPARATIVO ANUAL
    # -----------------------------------------------------

    seccion("📆 Comparativo anual")

    resumen_anual = (
        datos_filtrados
        .groupby("Año", as_index=False)
        .agg(
            Total_bruto=("Total_bruto", "sum"),
        )
    )

    resumen_anual["Total_bruto_MM"] = resumen_anual["Total_bruto"] / 1_000_000

    figura_anual = px.bar(
        resumen_anual,
        x="Año",
        y="Total_bruto_MM",
        title="Total bruto acumulado por año",
        template="plotly_white",
        text=resumen_anual["Total_bruto"].apply(pesos_texto),
    )

    figura_anual.update_traces(
        marker_color="#16a34a",
        textposition="outside",
    )

    figura_anual.update_xaxes(
        dtick=1,
    )

    figura_anual = formato_grafico(figura_anual, 420)
    figura_anual = aplicar_eje_millones_clp(figura_anual)

    st.plotly_chart(
        figura_anual,
        use_container_width=True,
    )

    # -----------------------------------------------------
    # PIE
    # -----------------------------------------------------

    st.markdown(
        (
            '<div class="pie-pagina">'
            f"<b>Panel desarrollado por {AUTOR}</b>"
            "<br>"
            f"{CARGO} | {EMPRESA}"
            "<br>"
            f"Versión {VERSION} | "
            f"Última actualización: "
            f"{nombre_periodo(fecha_maxima)}"
            "</div>"
        ),
        unsafe_allow_html=True,
    )


# =========================================================
# EJECUCION
# =========================================================

aplicar_estilo_general()
agregar_sello_agua_panel()
validar_acceso()

try:
    mostrar_panel()

except FileNotFoundError:
    st.error("No se encontró la planilla Excel.")

    st.write(
        "Verifica que el archivo esté en la misma carpeta que streamlit_app.py "
        "y que tenga uno de estos nombres:"
    )

    st.code(
        "RESUMEN ESTADOS DE PAGO.xlsx\n"
        "RESUMEN ESTADOS DE PAGO.xlsm\n"
        "RESUMEN ESTADOS DE PAGO.xls"
    )

except Exception as error:
    st.error("Ocurrió un error al cargar la planilla.")
    st.exception(error)
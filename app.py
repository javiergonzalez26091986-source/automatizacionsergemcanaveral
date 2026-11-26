import io
from datetime import date, datetime, time, timedelta
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. ZONA DE DATOS MAESTROS (Reemplazar aquí con la base real de SERGEM)
# ==============================================================================

# Festivos oficiales en Colombia (2026 - ampliar con 2027 al cierre de año)
FESTIVOS_COLOMBIA = {
    "2026-01-01", "2026-01-12", "2026-03-23", "2026-04-02", "2026-04-03",
    "2026-05-01", "2026-05-18", "2026-06-08", "2026-06-15", "2026-06-29",
    "2026-07-20", "2026-08-07", "2026-08-17", "2026-10-12", "2026-11-02",
    "2026-11-16", "2026-12-08", "2026-12-25",
}

# Credenciales administrativas y de supervisión (puedes cambiarlas aquí)
CREDENCIALES_ADMIN = {
    "admin_sergem": {"usuario": "admin", "clave": "sergem2026", "nombre": "Administración SERGEM"},
    "encargado_punto": {"usuario": "supervisor", "clave": "punto2026", "nombre": "Encargado de Punto"},
    "auditor_canaveral": {"usuario": "fabian", "clave": "canaveral2026", "nombre": "Auditoría Cañaveral (Sr. Fabián)"},
}

# Matriz inicial de tarifas por Sede y Vehículo (100% editable también desde la app)
DATOS_INICIALES_TARIFAS = [
    {"Sede": "Cali Norte", "Vehículo": "Moto", "Tarifa Día ($)": 138000, "Valor Hora Extra ($)": 11500, "Tarifa Fija Dom/Fest ($)": 165000, "Valor Hora Recargo Dom/Fest ($)": 14500},
    {"Sede": "Cali Norte", "Vehículo": "Camioneta", "Tarifa Día ($)": 255000, "Valor Hora Extra ($)": 19200, "Tarifa Fija Dom/Fest ($)": 298000, "Valor Hora Recargo Dom/Fest ($)": 24000},
    {"Sede": "Cali Sur", "Vehículo": "Moto", "Tarifa Día ($)": 138000, "Valor Hora Extra ($)": 11500, "Tarifa Fija Dom/Fest ($)": 165000, "Valor Hora Recargo Dom/Fest ($)": 14500},
    {"Sede": "Cali Sur", "Vehículo": "Chopper", "Tarifa Día ($)": 172000, "Valor Hora Extra ($)": 14200, "Tarifa Fija Dom/Fest ($)": 205000, "Valor Hora Recargo Dom/Fest ($)": 17800},
    {"Sede": "Yumbo", "Vehículo": "Camioneta", "Tarifa Día ($)": 268000, "Valor Hora Extra ($)": 20100, "Tarifa Fija Dom/Fest ($)": 312000, "Valor Hora Recargo Dom/Fest ($)": 25200},
    {"Sede": "Yumbo", "Vehículo": "Moto", "Tarifa Día ($)": 144000, "Valor Hora Extra ($)": 12000, "Tarifa Fija Dom/Fest ($)": 172000, "Valor Hora Recargo Dom/Fest ($)": 15000},
    {"Sede": "Palmira", "Vehículo": "Chopper", "Tarifa Día ($)": 179000, "Valor Hora Extra ($)": 14800, "Tarifa Fija Dom/Fest ($)": 214000, "Valor Hora Recargo Dom/Fest ($)": 18500},
    {"Sede": "Jamundí", "Vehículo": "Moto", "Tarifa Día ($)": 146000, "Valor Hora Extra ($)": 12200, "Tarifa Fija Dom/Fest ($)": 175000, "Valor Hora Recargo Dom/Fest ($)": 15300},
    {"Sede": "Tuluá", "Vehículo": "Camioneta", "Tarifa Día ($)": 282000, "Valor Hora Extra ($)": 21400, "Tarifa Fija Dom/Fest ($)": 330000, "Valor Hora Recargo Dom/Fest ($)": 26800},
    {"Sede": "Buga", "Vehículo": "Moto", "Tarifa Día ($)": 149000, "Valor Hora Extra ($)": 12500, "Tarifa Fija Dom/Fest ($)": 178000, "Valor Hora Recargo Dom/Fest ($)": 15600},
]

# Base maestra de trabajadores (el PIN permite que cada trabajador entre desde su celular)
DATOS_INICIALES_PERSONAL = [
    {"ID": "SG-101", "Cédula": "1144082310", "PIN": "2310", "Nombre Completo": "Carlos Andrés Mejía", "Sede": "Cali Norte", "Vehículo": "Camioneta", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-01", "Fecha Corte": ""},
    {"ID": "SG-102", "Cédula": "1107094512", "PIN": "4512", "Nombre Completo": "Jhonatan Stiven Ortiz", "Sede": "Cali Sur", "Vehículo": "Moto", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-01", "Fecha Corte": ""},
    {"ID": "SG-103", "Cédula": "1085312889", "PIN": "2889", "Nombre Completo": "Luis Fernando Caicedo", "Sede": "Cali Sur", "Vehículo": "Chopper", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-15", "Fecha Corte": ""},
    {"ID": "SG-104", "Cédula": "1112456781", "PIN": "6781", "Nombre Completo": "Diego Alejandro Rincón", "Sede": "Yumbo", "Vehículo": "Camioneta", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-01", "Fecha Corte": ""},
    {"ID": "SG-105", "Cédula": "1006234990", "PIN": "4990", "Nombre Completo": "Brayan Alexis Zapata", "Sede": "Yumbo", "Vehículo": "Moto", "Vinculación": "Prestación de Servicios", "Estado": "Activo", "Fecha Ingreso": "2026-10-03", "Fecha Corte": ""},
    {"ID": "SG-106", "Cédula": "1113678214", "PIN": "8214", "Nombre Completo": "Héctor Fabio Valencia", "Sede": "Palmira", "Vehículo": "Chopper", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-01", "Fecha Corte": ""},
    {"ID": "SG-107", "Cédula": "1143890115", "PIN": "0115", "Nombre Completo": "Kevin Santiago Mosquera", "Sede": "Jamundí", "Vehículo": "Moto", "Vinculación": "Prestación de Servicios", "Estado": "Activo", "Fecha Ingreso": "2026-11-02", "Fecha Corte": ""},
    {"ID": "SG-108", "Cédula": "1116254908", "PIN": "4908", "Nombre Completo": "Óscar Iván Restrepo", "Sede": "Tuluá", "Vehículo": "Camioneta", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-01", "Fecha Corte": ""},
    {"ID": "SG-109", "Cédula": "1115078432", "PIN": "8432", "Nombre Completo": "Sebastián Gómez Vélez", "Sede": "Buga", "Vehículo": "Moto", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-09-01", "Fecha Corte": ""},
    {"ID": "SG-110", "Cédula": "1144199654", "PIN": "9654", "Nombre Completo": "Andrés Felipe Quintero", "Sede": "Cali Norte", "Vehículo": "Moto", "Vinculación": "Prestación de Servicios", "Estado": "Activo", "Fecha Ingreso": "2026-11-16", "Fecha Corte": ""},
    {"ID": "SG-111", "Cédula": "1061782340", "PIN": "2340", "Nombre Completo": "Mauricio Paredes León", "Sede": "Palmira", "Vehículo": "Chopper", "Vinculación": "Prestación de Servicios", "Estado": "Liquidado / Retirado", "Fecha Ingreso": "2026-09-10", "Fecha Corte": "2026-11-12"},
]

DIAS_SEMANA_ES = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

# ==============================================================================
# 2. CONFIGURACIÓN DE PÁGINA Y ESTILOS CORPORATIVOS
# ==============================================================================
st.set_page_config(
    page_title="SERGEM | Control Operativo y Prefacturación Cañaveral",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .stApp {
        background-color: #F4F7FA !important;
        color: #1A202C !important;
    }
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0 !important;
    }
    section[data-testid="stSidebar"] * {
        color: #1A202C !important;
    }
    h1, h2, h3, h4, p, label, .stMarkdown, .stCaption {
        color: #1A202C !important;
    }
    .main-banner {
        background: linear-gradient(135deg, #0F2942 0%, #1E4E79 100%);
        padding: 1.35rem 1.75rem;
        border-radius: 12px;
        margin-bottom: 1.2rem;
        box-shadow: 0 4px 12px rgba(15, 41, 66, 0.15);
    }
    .main-banner h1 {
        margin: 0 !important;
        font-size: 1.45rem !important;
        font-weight: 700 !important;
        color: #FFFFFF !important;
    }
    .main-banner p {
        margin: 0.35rem 0 0 0 !important;
        font-size: 0.9rem !important;
        color: #E2E8F0 !important;
    }
    .badge-pill {
        display: inline-block;
        background-color: rgba(255, 255, 255, 0.16);
        border: 1px solid rgba(255, 255, 255, 0.35);
        color: #FFFFFF !important;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 0.78rem;
        font-weight: 600;
        margin-right: 6px;
        margin-top: 8px;
    }
    .login-card {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E0;
        border-top: 5px solid #0F2942;
        border-radius: 14px;
        padding: 1.8rem;
        box-shadow: 0 8px 24px rgba(15, 41, 66, 0.08);
    }
    .role-strip {
        background-color: #EBF8FF;
        border-left: 4px solid #2B6CB0;
        padding: 0.65rem 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        color: #1A365D !important;
        font-size: 0.9rem;
    }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-left: 4px solid #1E4E79 !important;
        padding: 0.9rem 1.1rem !important;
        border-radius: 10px !important;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04) !important;
    }
    div[data-testid="stMetric"] label,
    div[data-testid="stMetric"] label * {
        color: #4A5568 !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"],
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] * {
        color: #0F2942 !important;
        font-size: 1.4rem !important;
        font-weight: 700 !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #E2E8F0 !important;
        padding: 6px;
        border-radius: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: 600;
        background-color: transparent !important;
    }
    .stTabs [data-baseweb="tab"] p {
        color: #2D3748 !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0F2942 !important;
    }
    .stTabs [aria-selected="true"] p {
        color: #FFFFFF !important;
    }
    button[kind="primary"] p {
        color: #FFFFFF !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==============================================================================
# 3. MOTOR DE CÁLCULO HORARIO, CALENDARIO COLOMBIANO Y CONSOLIDACIÓN
# ==============================================================================
def calcular_jornada_automatica(
    fecha_op: date,
    h_ingreso: time,
    h_sal_almuerzo: time,
    h_reg_almuerzo: time,
    h_salida: time,
    jornada_base: float = 8.0,
) -> dict:
    """
    Calcula automáticamente:
    - Horas netas descontando almuerzo
    - Horas extras exactas en decimales
    - Día de la semana en español (Lunes a Domingo)
    - Detección automática de Domingo o Festivo en Colombia
    """
    dt_ing = datetime.combine(fecha_op, h_ingreso)
    dt_sal_alm = datetime.combine(fecha_op, h_sal_almuerzo)
    dt_reg_alm = datetime.combine(fecha_op, h_reg_almuerzo)
    dt_sal = datetime.combine(fecha_op, h_salida)

    if dt_sal < dt_ing:
        dt_sal += timedelta(days=1)
    if dt_reg_alm < dt_sal_alm:
        dt_reg_alm += timedelta(days=1)

    min_almuerzo = max(0.0, (dt_reg_alm - dt_sal_alm).total_seconds() / 60.0)
    min_brutos = max(0.0, (dt_sal - dt_ing).total_seconds() / 60.0)
    horas_netas = round(max(0.0, (min_brutos - min_almuerzo) / 60.0), 2)
    horas_extras = round(max(0.0, horas_netas - jornada_base), 2)

    dia_semana = DIAS_SEMANA_ES[fecha_op.weekday()]
    es_dom_fest = (fecha_op.weekday() == 6) or (fecha_op.strftime("%Y-%m-%d") in FESTIVOS_COLOMBIA)
    tipo_dia = "Domingo / Festivo" if es_dom_fest else "Ordinario"
    horas_recargo_df = horas_netas if es_dom_fest else 0.0

    return {
        "dia_semana": dia_semana,
        "tipo_dia": tipo_dia,
        "es_dom_fest": es_dom_fest,
        "horas_netas": horas_netas,
        "horas_extras": horas_extras,
        "horas_recargo_df": horas_recargo_df,
    }


def obtener_quincena_label(fecha_str: str) -> str:
    dt = datetime.strptime(fecha_str, "%Y-%m-%d")
    meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    rango = "01-15" if dt.day <= 15 else "16-31"
    return f"{rango} {meses[dt.month - 1]} {dt.year}"


def init_state():
    if "auth" not in st.session_state:
        st.session_state.auth = {"autenticado": False, "rol": None, "nombre": None, "cedula": None}

    if "jornada_base" not in st.session_state:
        st.session_state.jornada_base = 8.0

    if "tarifas" not in st.session_state:
        st.session_state.tarifas = pd.DataFrame(DATOS_INICIALES_TARIFAS)

    if "personal" not in st.session_state:
        st.session_state.personal = pd.DataFrame(DATOS_INICIALES_PERSONAL)

    if "registros_diarios" not in st.session_state:
        # Registros diarios de ejemplo (incluye pendientes por validar y aprobados)
        st.session_state.registros_diarios = pd.DataFrame(
            [
                {
                    "ID_Reg": "REG-001",
                    "Fecha": "2026-11-25",
                    "Día": 25,
                    "Día Semana": "Miércoles",
                    "Tipo Día": "Ordinario",
                    "Sede": "Cali Norte",
                    "Cédula": "1144082310",
                    "Mensajero": "Carlos Andrés Mejía",
                    "Vehículo": "Camioneta",
                    "Ingreso": "07:00",
                    "Salida Almuerzo": "12:00",
                    "Regreso Almuerzo": "13:00",
                    "Salida Turno": "18:15",
                    "Horas Netas": 10.25,
                    "Horas Extras": 2.25,
                    "Horas Dom/Fest": 0.00,
                    "Validación": "Pendiente Aprobación",
                    "Observación / Soporte": "Refuerzo despacho ruta norte - planilla #512",
                },
                {
                    "ID_Reg": "REG-002",
                    "Fecha": "2026-11-25",
                    "Día": 25,
                    "Día Semana": "Miércoles",
                    "Tipo Día": "Ordinario",
                    "Sede": "Yumbo",
                    "Cédula": "1006234990",
                    "Mensajero": "Brayan Alexis Zapata",
                    "Vehículo": "Moto",
                    "Ingreso": "07:30",
                    "Salida Almuerzo": "12:30",
                    "Regreso Almuerzo": "13:30",
                    "Salida Turno": "17:30",
                    "Horas Netas": 9.00,
                    "Horas Extras": 1.00,
                    "Horas Dom/Fest": 0.00,
                    "Validación": "Pendiente Aprobación",
                    "Observación / Soporte": "Apoyo entregas zona industrial",
                },
                {
                    "ID_Reg": "REG-003",
                    "Fecha": "2026-11-22",
                    "Día": 22,
                    "Día Semana": "Domingo",
                    "Tipo Día": "Domingo / Festivo",
                    "Sede": "Cali Norte",
                    "Cédula": "1144082310",
                    "Mensajero": "Carlos Andrés Mejía",
                    "Vehículo": "Camioneta",
                    "Ingreso": "08:00",
                    "Salida Almuerzo": "12:00",
                    "Regreso Almuerzo": "13:00",
                    "Salida Turno": "16:00",
                    "Horas Netas": 7.00,
                    "Horas Extras": 0.00,
                    "Horas Dom/Fest": 7.00,
                    "Validación": "Aprobado",
                    "Observación / Soporte": "Operación dominical autorizada por jefe de bodega",
                },
                {
                    "ID_Reg": "REG-004",
                    "Fecha": "2026-11-18",
                    "Día": 18,
                    "Día Semana": "Miércoles",
                    "Tipo Día": "Ordinario",
                    "Sede": "Cali Sur",
                    "Cédula": "1107094512",
                    "Mensajero": "Jhonatan Stiven Ortiz",
                    "Vehículo": "Moto",
                    "Ingreso": "07:00",
                    "Salida Almuerzo": "12:00",
                    "Regreso Almuerzo": "13:00",
                    "Salida Turno": "17:30",
                    "Horas Netas": 9.50,
                    "Horas Extras": 1.50,
                    "Horas Dom/Fest": 0.00,
                    "Validación": "Aprobado",
                    "Observación / Soporte": "Ruta sur completa",
                },
                {
                    "ID_Reg": "REG-005",
                    "Fecha": "2026-11-16",
                    "Día": 16,
                    "Día Semana": "Lunes",
                    "Tipo Día": "Domingo / Festivo",
                    "Sede": "Palmira",
                    "Cédula": "1113678214",
                    "Mensajero": "Héctor Fabio Valencia",
                    "Vehículo": "Chopper",
                    "Ingreso": "08:00",
                    "Salida Almuerzo": "12:00",
                    "Regreso Almuerzo": "13:00",
                    "Salida Turno": "17:00",
                    "Horas Netas": 8.00,
                    "Horas Extras": 0.00,
                    "Horas Dom/Fest": 8.00,
                    "Validación": "Aprobado",
                    "Observación / Soporte": "Festivo Independencia de Cartagena",
                },
                {
                    "ID_Reg": "REG-006",
                    "Fecha": "2026-11-17",
                    "Día": 17,
                    "Día Semana": "Martes",
                    "Tipo Día": "Ordinario",
                    "Sede": "Tuluá",
                    "Cédula": "1116254908",
                    "Mensajero": "Óscar Iván Restrepo",
                    "Vehículo": "Camioneta",
                    "Ingreso": "07:00",
                    "Salida Almuerzo": "12:00",
                    "Regreso Almuerzo": "13:00",
                    "Salida Turno": "18:00",
                    "Horas Netas": 10.00,
                    "Horas Extras": 2.00,
                    "Horas Dom/Fest": 0.00,
                    "Validación": "Aprobado",
                    "Observación / Soporte": "Refuerzo despacho martes",
                },
            ]
        )

    if "prefactura_q1" not in st.session_state:
        st.session_state.prefactura_q1 = pd.DataFrame(
            [
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Cali Norte",
                    "Cédula": "1144082310",
                    "Mensajero": "Carlos Andrés Mejía",
                    "Vehículo": "Camioneta",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,8,9,10,11,12,13,15",
                    "Cant. Días": 13,
                    "Horas Extras": 26.77,
                    "Días Dom/Fest": 2,
                    "Horas Dom/Fest": 16.00,
                    "Valor Liquidado ($)": 4031984,
                    "Observación Cañaveral (Fabián)": "Validar 2,5 HE del día 08 en ruta norte.",
                    "Respuesta / Ajuste SERGEM": "Soportado con marcación horaria y planilla #408.",
                    "Estado Auditoría": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Cali Sur",
                    "Cédula": "1107094512",
                    "Mensajero": "Jhonatan Stiven Ortiz",
                    "Vehículo": "Moto",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,7,8,9,10,11,12,13,14",
                    "Cant. Días": 14,
                    "Horas Extras": 12.50,
                    "Días Dom/Fest": 1,
                    "Horas Dom/Fest": 8.00,
                    "Valor Liquidado ($)": 2158750,
                    "Observación Cañaveral (Fabián)": "OK conforme.",
                    "Respuesta / Ajuste SERGEM": "Sin novedad.",
                    "Estado Auditoría": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Cali Sur",
                    "Cédula": "1085312889",
                    "Mensajero": "Luis Fernando Caicedo",
                    "Vehículo": "Chopper",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "2,3,4,5,6,8,9,10,11,12,13,15",
                    "Cant. Días": 12,
                    "Horas Extras": 18.25,
                    "Días Dom/Fest": 1,
                    "Horas Dom/Fest": 8.00,
                    "Valor Liquidado ($)": 2436550,
                    "Observación Cañaveral (Fabián)": "Revisar turno dominical del día 08.",
                    "Respuesta / Ajuste SERGEM": "Validado con marcación de ingreso 08:00 y salida 17:00.",
                    "Estado Auditoría": "Ajustado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Yumbo",
                    "Cédula": "1112456781",
                    "Mensajero": "Diego Alejandro Rincón",
                    "Vehículo": "Camioneta",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,8,9,10,11,12,13,14,15",
                    "Cant. Días": 14,
                    "Horas Extras": 21.40,
                    "Días Dom/Fest": 2,
                    "Horas Dom/Fest": 16.00,
                    "Valor Liquidado ($)": 4395340,
                    "Observación Cañaveral (Fabián)": "",
                    "Respuesta / Ajuste SERGEM": "",
                    "Estado Auditoría": "Pendiente",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Yumbo",
                    "Cédula": "1006234990",
                    "Mensajero": "Brayan Alexis Zapata",
                    "Vehículo": "Moto",
                    "Vinculación": "Prestación de Servicios",
                    "Días Trabajados": "3,4,5,10,11,12",
                    "Cant. Días": 6,
                    "Horas Extras": 6.50,
                    "Días Dom/Fest": 0,
                    "Horas Dom/Fest": 0.00,
                    "Valor Liquidado ($)": 942000,
                    "Observación Cañaveral (Fabián)": "Confirmar si cubrió incapacidad en Yumbo.",
                    "Respuesta / Ajuste SERGEM": "Sí, reemplazo por incapacidad de ruta Yumbo-02.",
                    "Estado Auditoría": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Palmira",
                    "Cédula": "1113678214",
                    "Mensajero": "Héctor Fabio Valencia",
                    "Vehículo": "Chopper",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,8,9,10,11,12,13,15",
                    "Cant. Días": 13,
                    "Horas Extras": 15.80,
                    "Días Dom/Fest": 2,
                    "Horas Dom/Fest": 16.00,
                    "Valor Liquidado ($)": 2708840,
                    "Observación Cañaveral (Fabián)": "",
                    "Respuesta / Ajuste SERGEM": "",
                    "Estado Auditoría": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Jamundí",
                    "Cédula": "1143890115",
                    "Mensajero": "Kevin Santiago Mosquera",
                    "Vehículo": "Moto",
                    "Vinculación": "Prestación de Servicios",
                    "Días Trabajados": "2,4,5,9,11,12,14",
                    "Cant. Días": 7,
                    "Horas Extras": 8.25,
                    "Días Dom/Fest": 0,
                    "Horas Dom/Fest": 0.00,
                    "Valor Liquidado ($)": 1122650,
                    "Observación Cañaveral (Fabián)": "Pendiente validar 1,25 HE del día 14.",
                    "Respuesta / Ajuste SERGEM": "",
                    "Estado Auditoría": "Con Observación",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Tuluá",
                    "Cédula": "1116254908",
                    "Mensajero": "Óscar Iván Restrepo",
                    "Vehículo": "Camioneta",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,8,9,10,11,12,13,14,15",
                    "Cant. Días": 14,
                    "Horas Extras": 19.60,
                    "Días Dom/Fest": 2,
                    "Horas Dom/Fest": 16.00,
                    "Valor Liquidado ($)": 4612240,
                    "Observación Cañaveral (Fabián)": "OK conforme.",
                    "Respuesta / Ajuste SERGEM": "Sin novedad.",
                    "Estado Auditoría": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Buga",
                    "Cédula": "1115078432",
                    "Mensajero": "Sebastián Gómez Vélez",
                    "Vehículo": "Moto",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,8,9,10,11,12,13",
                    "Cant. Días": 12,
                    "Horas Extras": 10.00,
                    "Días Dom/Fest": 1,
                    "Horas Dom/Fest": 8.00,
                    "Valor Liquidado ($)": 1996800,
                    "Observación Cañaveral (Fabián)": "",
                    "Respuesta / Ajuste SERGEM": "",
                    "Estado Auditoría": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Palmira",
                    "Cédula": "1061782340",
                    "Mensajero": "Mauricio Paredes León (Retirado 12-Nov)",
                    "Vehículo": "Chopper",
                    "Vinculación": "Prestación de Servicios",
                    "Días Trabajados": "1,2,3,4,5,8,9,10,11,12",
                    "Cant. Días": 10,
                    "Horas Extras": 4.50,
                    "Días Dom/Fest": 1,
                    "Horas Dom/Fest": 8.00,
                    "Valor Liquidado ($)": 1968600,
                    "Observación Cañaveral (Fabián)": "Liquidación proporcional hasta fecha de corte 12/11.",
                    "Respuesta / Ajuste SERGEM": "Confirmado, solo se cobran los 10 días laborados antes del retiro.",
                    "Estado Auditoría": "Autorizado",
                },
            ]
        )


init_state()


def consolidar_desde_registros_aprobados():
    """Cruza las marcaciones aprobadas con la matriz de tarifas para generar líneas de prefactura."""
    df_reg = st.session_state.registros_diarios
    aprobados = df_reg[df_reg["Validación"] == "Aprobado"].copy()
    if aprobados.empty:
        return pd.DataFrame()

    aprobados["Quincena"] = aprobados["Fecha"].apply(obtener_quincena_label)
    df_tarifas = st.session_state.tarifas
    df_pers = st.session_state.personal

    filas = []
    for (quincena, cedula, mensajero, sede, vehiculo), grupo in aprobados.groupby(
        ["Quincena", "Cédula", "Mensajero", "Sede", "Vehículo"]
    ):
        dias_unicos = sorted(grupo["Día"].unique().tolist())
        dias_str = ",".join(str(d) for d in dias_unicos)
        cant_dias = len(dias_unicos)
        he_total = round(grupo["Horas Extras"].sum(), 2)
        dias_df = len(grupo[grupo["Tipo Día"] == "Domingo / Festivo"])
        dias_ord = cant_dias - dias_df
        horas_df = round(grupo["Horas Dom/Fest"].sum(), 2)

        # Buscar modalidad de vinculación
        match_p = df_pers[df_pers["Cédula"] == cedula]
        vinculacion = match_p.iloc[0]["Vinculación"] if not match_p.empty else "Fijo / Planta"

        # Buscar tarifa vigente
        match_t = df_tarifas[(df_tarifas["Sede"] == sede) & (df_tarifas["Vehículo"] == vehiculo)]
        if not match_t.empty:
            t_dia = float(match_t.iloc[0]["Tarifa Día ($)"])
            t_he = float(match_t.iloc[0]["Valor Hora Extra ($)"])
            t_df_dia = float(match_t.iloc[0]["Tarifa Fija Dom/Fest ($)"])
            t_df_hora = float(match_t.iloc[0]["Valor Hora Recargo Dom/Fest ($)"])
        else:
            t_dia, t_he, t_df_dia, t_df_hora = 140000.0, 12000.0, 170000.0, 15000.0

        valor_total = int(
            round((dias_ord * t_dia) + (dias_df * t_df_dia) + (he_total * t_he) + (horas_df * t_df_hora))
        )

        filas.append(
            {
                "Quincena": quincena,
                "Sede": sede,
                "Cédula": cedula,
                "Mensajero": mensajero,
                "Vehículo": vehiculo,
                "Vinculación": vinculacion,
                "Días Trabajados": dias_str,
                "Cant. Días": cant_dias,
                "Horas Extras": he_total,
                "Días Dom/Fest": dias_df,
                "Horas Dom/Fest": horas_df,
                "Valor Liquidado ($)": valor_total,
                "Observación Cañaveral (Fabián)": "",
                "Respuesta / Ajuste SERGEM": "Liquidado automáticamente desde marcaciones aprobadas.",
                "Estado Auditoría": "Pendiente",
            }
        )
    return pd.DataFrame(filas)


# ==============================================================================
# 4. GENERADOR DE EXCEL OFICIAL (FORMATO EXACTO CAÑAVERAL)
# ==============================================================================
def generar_excel_prefactura(df_prefactura: pd.DataFrame, df_diario: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_prefactura.to_excel(writer, index=False, sheet_name="Prefactura_Quincenal", startrow=3)
        df_diario.to_excel(writer, index=False, sheet_name="Respaldo_Auditoria_Diaria", startrow=2)

        wb = writer.book
        ws = wb["Prefactura_Quincenal"]

        ws.merge_cells("A1:O1")
        ws["A1"] = "SERGEM S.A.S. — PREFACTURA DE OPERACIÓN LOGÍSTICA CAÑAVERAL"
        ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[1].height = 28

        ws.merge_cells("A2:O2")
        ws["A2"] = f"Generado el {date.today().strftime('%d/%m/%Y')}  |  Liquidación automática por marcación horaria y calendario de festivos"
        ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="1E4E79")
        ws.row_dimensions[2].height = 20

        header_fill = PatternFill(start_color="1E4E79", end_color="1E4E79", fill_type="solid")
        audit_fill = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid")
        thin_border = Border(
            left=Side(style="thin", color="CBD5E0"),
            right=Side(style="thin", color="CBD5E0"),
            top=Side(style="thin", color="CBD5E0"),
            bottom=Side(style="thin", color="CBD5E0"),
        )

        for col_idx in range(1, len(df_prefactura.columns) + 1):
            cell = ws.cell(row=4, column=col_idx)
            cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            cell.fill = audit_fill if col_idx >= 13 else header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[4].height = 28

        for row_idx in range(5, 5 + len(df_prefactura)):
            estado_val = str(ws.cell(row=row_idx, column=15).value)
            for col_idx in range(1, len(df_prefactura.columns) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = Font(name="Calibri", size=10)
                cell.border = thin_border
                if col_idx == 12:
                    cell.number_format = '"$"#,##0'
                    cell.font = Font(name="Calibri", size=10, bold=True)
                elif col_idx in (9, 11):
                    cell.number_format = "0.00"
                elif col_idx == 15:
                    cell.alignment = Alignment(horizontal="center")
                    if estado_val == "Autorizado":
                        cell.fill = PatternFill(start_color="C6F6D5", end_color="C6F6D5", fill_type="solid")
                        cell.font = Font(name="Calibri", size=10, bold=True, color="22543D")
                    elif estado_val == "Con Observación":
                        cell.fill = PatternFill(start_color="FED7D7", end_color="FED7D7", fill_type="solid")
                        cell.font = Font(name="Calibri", size=10, bold=True, color="822727")
                    elif estado_val == "Ajustado":
                        cell.fill = PatternFill(start_color="FEFCBF", end_color="FEFCBF", fill_type="solid")
                        cell.font = Font(name="Calibri", size=10, bold=True, color="744210")

        total_row = 5 + len(df_prefactura)
        ws.cell(row=total_row, column=11, value="TOTAL PREFACTURA:").font = Font(bold=True)
        total_cell = ws.cell(row=total_row, column=12, value=f"=SUM(L5:L{total_row - 1})")
        total_cell.font = Font(bold=True, size=11, color="0F2942")
        total_cell.number_format = '"$"#,##0'
        total_cell.fill = PatternFill(start_color="EBF8FF", end_color="EBF8FF", fill_type="solid")

        for col_idx in range(1, len(df_prefactura.columns) + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 17
        ws.column_dimensions["D"].width = 28
        ws.column_dimensions["G"].width = 26
        ws.column_dimensions["M"].width = 34
        ws.column_dimensions["N"].width = 34

    return output.getvalue()


# ==============================================================================
# 5. PANTALLA DE INICIO DE SESIÓN (CONTROL DE ACCESO POR PERFIL Y CONTRASEÑA)
# ==============================================================================
def render_login():
    st.markdown(
        """
        <div class="main-banner">
            <h1>🚚 SERGEM S.A.S. — Portal de Control Horario y Prefacturación Cañaveral</h1>
            <p>Ingreso seguro por perfil: Marcación diaria de trabajadores, validación de punto y administración gerencial.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_esp1, col_centro, col_esp2 = st.columns([1, 1.45, 1])
    with col_centro:
        st.markdown("### 🔐 Iniciar Sesión")
        perfil_login = st.selectbox(
            "Selecciona tu perfil de acceso:",
            [
                "👷 Trabajador / Mensajero (Marcación de Turno)",
                "📋 Encargado de Punto (Validación Diaria)",
                "🔐 Administración SERGEM (Contraseña Maestra)",
                "👨‍💼 Auditoría Cañaveral (Sr. Fabián)",
            ],
        )

        if perfil_login.startswith("👷"):
            with st.form("login_trabajador"):
                cedula_in = st.text_input("Número de Cédula (sin puntos):", placeholder="Ej. 1144082310")
                pin_in = st.text_input("PIN / Contraseña personal:", type="password", placeholder="4 dígitos")
                if st.form_submit_button("Ingresar a mi Turno", use_container_width=True, type="primary"):
                    df_p = st.session_state.personal
                    ced_limpia = cedula_in.replace(".", "").strip()
                    match = df_p[
                        (df_p["Cédula"].astype(str).str.replace(".", "", regex=False) == ced_limpia)
                        & (df_p["PIN"].astype(str) == pin_in.strip())
                    ]
                    if match.empty:
                        st.error("Cédula o PIN incorrectos. Verifica con Administración SERGEM.")
                    elif match.iloc[0]["Estado"] != "Activo":
                        st.warning("Este usuario figura como Liquidado / Retirado y no puede registrar turnos.")
                    else:
                        st.session_state.auth = {
                            "autenticado": True,
                            "rol": "trabajador",
                            "nombre": match.iloc[0]["Nombre Completo"],
                            "cedula": match.iloc[0]["Cédula"],
                        }
                        st.rerun()

        else:
            mapa_cred = {
                "📋 Encargado de Punto (Validación Diaria)": ("encargado_punto", CREDENCIALES_ADMIN["encargado_punto"]),
                "🔐 Administración SERGEM (Contraseña Maestra)": ("admin_sergem", CREDENCIALES_ADMIN["admin_sergem"]),
                "👨‍💼 Auditoría Cañaveral (Sr. Fabián)": ("auditor_canaveral", CREDENCIALES_ADMIN["auditor_canaveral"]),
            }
            rol_key, datos_cred = mapa_cred[perfil_login]

            with st.form("login_admin"):
                user_in = st.text_input("Usuario:", value=datos_cred["usuario"])
                pass_in = st.text_input("Contraseña:", type="password", placeholder="Digita la contraseña del perfil")
                if st.form_submit_button("Ingresar al Panel", use_container_width=True, type="primary"):
                    if user_in.strip() == datos_cred["usuario"] and pass_in.strip() == datos_cred["clave"]:
                        st.session_state.auth = {
                            "autenticado": True,
                            "rol": rol_key,
                            "nombre": datos_cred["nombre"],
                            "cedula": None,
                        }
                        st.rerun()
                    else:
                        st.error("Usuario o contraseña incorrectos.")

        with st.expander("🔑 Credenciales de Prueba (Prototipo)"):
            st.markdown(
                """
                * **Trabajador (Carlos Andrés Mejía):** Cédula `1144082310` | PIN `2310`
                * **Encargado de Punto:** Usuario `supervisor` | Clave `punto2026`
                * **Administración SERGEM:** Usuario `admin` | Clave `sergem2026`
                * **Auditoría Cañaveral (Sr. Fabián):** Usuario `fabian` | Clave `canaveral2026`
                """
            )


# ==============================================================================
# 6. MÓDULO TRABAJADOR: MARCACIÓN DE 4 HORAS Y CÁLCULO AUTOMÁTICO
# ==============================================================================
def render_vista_trabajador():
    cedula_usr = st.session_state.auth["cedula"]
    df_p = st.session_state.personal
    info_usr = df_p[df_p["Cédula"] == cedula_usr].iloc[0]

    st.markdown(
        f'<div class="role-strip">👷 <b>Bienvenido, {info_usr["Nombre Completo"]}</b> | '
        f'📍 Sede: <b>{info_usr["Sede"]}</b> | 🚚 Vehículo: <b>{info_usr["Vehículo"]}</b> | '
        f'📋 {info_usr["Vinculación"]}</div>',
        unsafe_allow_html=True,
    )

    col_form, col_hist = st.columns([1.15, 1.85])

    with col_form:
        st.subheader("⏱️ Registrar Horario del Día")
        st.caption(
            "Digita únicamente las 4 horas de tu jornada. El sistema calcula automáticamente "
            "el tiempo de almuerzo, tus horas extras y si aplica recargo dominical o festivo."
        )

        fecha_op = st.date_input("Fecha del turno:", value=date(2026, 11, 25))
        c1, c2 = st.columns(2)
        h_ing = c1.time_input("1. Hora de Ingreso:", value=time(7, 0))
        h_sal_alm = c2.time_input("2. Salida a Almuerzo:", value=time(12, 0))
        c3, c4 = st.columns(2)
        h_reg_alm = c3.time_input("3. Regreso de Almuerzo:", value=time(13, 0))
        h_sal = c4.time_input("4. Salida de Turno:", value=time(17, 30))

        # Vista previa automática antes de guardar
        calc = calcular_jornada_automatica(
            fecha_op, h_ing, h_sal_alm, h_reg_alm, h_sal, st.session_state.jornada_base
        )
        st.info(
            f"📅 **{calc['dia_semana']} ({calc['tipo_dia']})**  \n"
            f"⏱️ Horas netas laboradas: **{calc['horas_netas']:.2f} h** | "
            f"⚡ Horas extras calculadas: **{calc['horas_extras']:.2f} h**"
        )

        obs_txt = st.text_input(
            "Observación / Soporte (opcional):",
            placeholder="Ej. Ruta norte completa, apoyo despacho o número de planilla",
        )

        if st.button("✅ Enviar Reporte de Horario", use_container_width=True, type="primary"):
            nuevo_id = f"REG-{len(st.session_state.registros_diarios) + 1:03d}"
            nuevo_reg = pd.DataFrame(
                [
                    {
                        "ID_Reg": nuevo_id,
                        "Fecha": fecha_op.strftime("%Y-%m-%d"),
                        "Día": fecha_op.day,
                        "Día Semana": calc["dia_semana"],
                        "Tipo Día": calc["tipo_dia"],
                        "Sede": info_usr["Sede"],
                        "Cédula": info_usr["Cédula"],
                        "Mensajero": info_usr["Nombre Completo"],
                        "Vehículo": info_usr["Vehículo"],
                        "Ingreso": h_ing.strftime("%H:%M"),
                        "Salida Almuerzo": h_sal_alm.strftime("%H:%M"),
                        "Regreso Almuerzo": h_reg_alm.strftime("%H:%M"),
                        "Salida Turno": h_sal.strftime("%H:%M"),
                        "Horas Netas": calc["horas_netas"],
                        "Horas Extras": calc["horas_extras"],
                        "Horas Dom/Fest": calc["horas_recargo_df"],
                        "Validación": "Pendiente Aprobación",
                        "Observación / Soporte": obs_txt if obs_txt else "Sin novedad",
                    }
                ]
            )
            st.session_state.registros_diarios = pd.concat(
                [nuevo_reg, st.session_state.registros_diarios], ignore_index=True
            )
            st.success("✅ Marcación enviada correctamente. Queda pendiente de visto bueno del encargado del punto.")
            st.rerun()

    with col_hist:
        st.subheader("📋 Mis Turnos Registrados")
        mis_regs = st.session_state.registros_diarios[
            st.session_state.registros_diarios["Cédula"] == cedula_usr
        ][
            [
                "Fecha", "Día Semana", "Tipo Día", "Ingreso", "Salida Almuerzo",
                "Regreso Almuerzo", "Salida Turno", "Horas Netas", "Horas Extras",
                "Validación", "Observación / Soporte",
            ]
        ]
        st.dataframe(mis_regs, use_container_width=True, hide_index=True)


# ==============================================================================
# 7. MÓDULO ENCARGADO DE PUNTO: VALIDACIÓN DE MARCACIONES DIARIAS
# ==============================================================================
def render_vista_encargado_punto():
    st.markdown(
        '<div class="role-strip">📋 <b>Panel Encargado de Punto</b> — Revisa las marcaciones enviadas por los trabajadores y aprueba los turnos del día antes de pasar a prefacturación.</div>',
        unsafe_allow_html=True,
    )

    df_reg = st.session_state.registros_diarios
    pendientes = df_reg[df_reg["Validación"] == "Pendiente Aprobación"]

    m1, m2, m3 = st.columns(3)
    m1.metric("⏳ Pendientes por Aprobar", f"{len(pendientes)} registros")
    m2.metric("✅ Registros Aprobados", f"{len(df_reg[df_reg['Validación'] == 'Aprobado'])} registros")
    m3.metric("⚡ Horas Extras en Revisión", f"{pendientes['Horas Extras'].sum():.2f} h")

    st.subheader("✅ Bandeja de Validación Diaria")
    st.caption("Cambia el estado en la columna **Validación** a *Aprobado* o *Rechazado* y guarda los cambios.")

    edit_reg = st.data_editor(
        st.session_state.registros_diarios,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Validación": st.column_config.SelectboxColumn(
                "Validación",
                options=["Pendiente Aprobación", "Aprobado", "Rechazado"],
                required=True,
            ),
            "Horas Netas": st.column_config.NumberColumn(format="%.2f"),
            "Horas Extras": st.column_config.NumberColumn(format="%.2f"),
            "Horas Dom/Fest": st.column_config.NumberColumn(format="%.2f"),
        },
        disabled=[
            "ID_Reg", "Fecha", "Día", "Día Semana", "Tipo Día", "Sede", "Cédula",
            "Mensajero", "Vehículo", "Ingreso", "Salida Almuerzo", "Regreso Almuerzo",
            "Salida Turno", "Horas Netas", "Horas Extras", "Horas Dom/Fest",
        ],
        key="editor_validacion_punto",
    )
    st.session_state.registros_diarios = edit_reg

    if st.button("💾 Confirmar Validaciones del Punto", type="primary"):
        st.success("✅ Validaciones actualizadas. Los turnos aprobados ya están listos para consolidar en Administración.")


# ==============================================================================
# 8. MÓDULOS ADMINISTRACIÓN SERGEM Y AUDITORÍA CAÑAVERAL
# ==============================================================================
def render_kpis_gerenciales():
    df_tot = st.session_state.prefactura_q1
    total_pres = df_tot["Valor Liquidado ($)"].sum()
    total_aut = df_tot[df_tot["Estado Auditoría"] == "Autorizado"]["Valor Liquidado ($)"].sum()
    total_rev = df_tot[df_tot["Estado Auditoría"] != "Autorizado"]["Valor Liquidado ($)"].sum()
    activos = len(st.session_state.personal[st.session_state.personal["Estado"] == "Activo"])

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("💰 Total Prefactura", f"${total_pres:,.0f}")
    pct = (total_aut / total_pres * 100) if total_pres > 0 else 0
    k2.metric("✅ Autorizado Facturar", f"${total_aut:,.0f}", f"{pct:.1f}% aprobado")
    k3.metric("🔍 En Revisión / Glosas", f"${total_rev:,.0f}", f"{len(df_tot[df_tot['Estado Auditoría'] != 'Autorizado'])} ítems")
    k4.metric("👥 Personal Activo", f"{activos} colaboradores", "Sin listas rígidas")


def render_tab_prefactura(modo_auditor=False):
    col_f1, col_f2, col_f3 = st.columns([1.2, 1.2, 1.8])
    with col_f1:
        filtro_sede = st.selectbox(
            "Filtrar por Sede:",
            ["Todas las sedes"] + sorted(st.session_state.prefactura_q1["Sede"].unique().tolist()),
            key=f"f_sede_{modo_auditor}",
        )
    with col_f2:
        filtro_estado = st.selectbox(
            "Filtrar por Estado:",
            ["Todos", "Autorizado", "Ajustado", "Con Observación", "Pendiente"],
            key=f"f_est_{modo_auditor}",
        )
    with col_f3:
        if not modo_auditor:
            if st.button("🔄 Agregar / Recalcular desde Marcaciones Aprobadas (16-31 Nov)", use_container_width=True):
                df_nuevos = consolidar_desde_registros_aprobados()
                if not df_nuevos.empty:
                    base_q1 = st.session_state.prefactura_q1[
                        st.session_state.prefactura_q1["Quincena"] == "01-15 Nov 2026"
                    ]
                    st.session_state.prefactura_q1 = pd.concat([base_q1, df_nuevos], ignore_index=True)
                    st.success("✅ Corte 16-31 Nov actualizado automáticamente desde las marcaciones aprobadas.")
                    st.rerun()
                else:
                    st.warning("No hay registros aprobados pendientes por consolidar.")

    df_view = st.session_state.prefactura_q1.copy()
    if filtro_sede != "Todas las sedes":
        df_view = df_view[df_view["Sede"] == filtro_sede]
    if filtro_estado != "Todos":
        df_view = df_view[df_view["Estado Auditoría"] == filtro_estado]

    edited_df = st.data_editor(
        df_view,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Valor Liquidado ($)": st.column_config.NumberColumn(format="$ %d"),
            "Horas Extras": st.column_config.NumberColumn(format="%.2f"),
            "Horas Dom/Fest": st.column_config.NumberColumn(format="%.2f"),
            "Estado Auditoría": st.column_config.SelectboxColumn(
                "Estado Auditoría",
                options=["Pendiente", "Con Observación", "Ajustado", "Autorizado"],
                required=True,
            ),
            "Observación Cañaveral (Fabián)": st.column_config.TextColumn(width="large"),
            "Respuesta / Ajuste SERGEM": st.column_config.TextColumn(width="large"),
        },
        disabled=[
            "Quincena", "Sede", "Cédula", "Mensajero", "Vehículo", "Vinculación",
            "Días Trabajados", "Cant. Días", "Horas Extras", "Días Dom/Fest",
            "Horas Dom/Fest", "Valor Liquidado ($)",
        ],
        key=f"ed_pref_{modo_auditor}",
    )

    if filtro_sede == "Todas las sedes" and filtro_estado == "Todos":
        st.session_state.prefactura_q1 = edited_df.copy()

    st.divider()
    c_b1, c_b2 = st.columns([1.4, 2.6])
    with c_b1:
        excel_bytes = generar_excel_prefactura(
            st.session_state.prefactura_q1, st.session_state.registros_diarios
        )
        st.download_button(
            label="📥 Descargar Excel Oficial Cañaveral (.xlsx)",
            data=excel_bytes,
            file_name="PREFACTURA_SERGEM_CANAVERAL_NOV2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            key=f"dl_{modo_auditor}",
        )
    with c_b2:
        st.info(
            "Incluye hoja **Prefactura_Quincenal** (días separados por comas `1,2,3...`, horas extras decimales, "
            "dominicales/festivos y glosas) + hoja **Respaldo_Auditoria_Diaria** con marcaciones horarias y día de la semana."
        )


def render_tab_estadisticas_dias():
    st.subheader("📊 Estadística Operativa por Día de la Semana")
    st.caption(
        "Permite identificar qué días de la semana (Lunes a Domingo) presentan mayor demanda de turnos, "
        "apoyos y horas extras sin limitar los refuerzos únicamente a los miércoles."
    )

    df_reg = st.session_state.registros_diarios.copy()
    resumen_dias = (
        df_reg.groupby("Día Semana", as_index=False)
        .agg(
            Turnos_Registrados=("ID_Reg", "count"),
            Total_Horas_Netas=("Horas Netas", "sum"),
            Total_Horas_Extras=("Horas Extras", "sum"),
        )
    )
    resumen_dias["Orden"] = resumen_dias["Día Semana"].apply(lambda d: DIAS_SEMANA_ES.index(d))
    resumen_dias = resumen_dias.sort_values("Orden").drop(columns=["Orden"])

    c1, c2 = st.columns([1.2, 1.8])
    with c1:
        st.dataframe(resumen_dias, use_container_width=True, hide_index=True)
    with c2:
        st.bar_chart(resumen_dias.set_index("Día Semana")[["Turnos_Registrados", "Total_Horas_Extras"]])


def render_tab_personal():
    st.subheader("👥 Gestión Dinámica de Personal y Accesos (PIN)")
    st.caption(
        "Al crear un colaborador se le asigna su PIN de acceso celular. Si alguien se retira, "
        "se liquida con fecha de corte sin alterar los históricos anteriores."
    )

    col_alta, col_retiro = st.columns(2)

    with col_alta:
        with st.expander("➕ Registrar Nuevo Ingreso (Fijo o Servicios)", expanded=True):
            with st.form("form_nuevo_ingreso", clear_on_submit=True):
                c1, c2 = st.columns(2)
                cedula_in = c1.text_input("Cédula (sin puntos):", placeholder="Ej. 1144555222")
                nombre_in = c2.text_input("Nombre Completo:", placeholder="Ej. Juan Camilo Torres")
                c3, c4 = st.columns(2)
                sede_in = c3.selectbox("Sede:", sorted(st.session_state.tarifas["Sede"].unique().tolist()))
                veh_in = c4.selectbox("Vehículo:", ["Moto", "Camioneta", "Chopper"])
                c5, c6, c7 = st.columns(3)
                vinc_in = c5.selectbox("Modalidad:", ["Fijo / Planta", "Prestación de Servicios"])
                pin_in = c6.text_input("PIN Celular (4 dígitos):", placeholder="Ej. 5222")
                fecha_in = c7.date_input("Fecha Ingreso:", value=date(2026, 11, 25))

                if st.form_submit_button("💾 Activar Colaborador", use_container_width=True, type="primary"):
                    if cedula_in and nombre_in:
                        ced_limpia = cedula_in.replace(".", "").strip()
                        pin_final = pin_in.strip() if pin_in.strip() else ced_limpia[-4:]
                        nuevo_id = f"SG-{101 + len(st.session_state.personal)}"
                        fila_nueva = pd.DataFrame(
                            [
                                {
                                    "ID": nuevo_id,
                                    "Cédula": ced_limpia,
                                    "PIN": pin_final,
                                    "Nombre Completo": nombre_in.strip(),
                                    "Sede": sede_in,
                                    "Vehículo": veh_in,
                                    "Vinculación": vinc_in,
                                    "Estado": "Activo",
                                    "Fecha Ingreso": fecha_in.strftime("%Y-%m-%d"),
                                    "Fecha Corte": "",
                                }
                            ]
                        )
                        st.session_state.personal = pd.concat(
                            [fila_nueva, st.session_state.personal], ignore_index=True
                        )
                        st.success(f"✅ {nombre_in} activado con PIN `{pin_final}`.")
                        st.rerun()
                    else:
                        st.error("Ingresa cédula y nombre completo.")

    with col_retiro:
        with st.expander("🚪 Liquidar / Retirar Personal con Fecha de Corte", expanded=True):
            activos_lista = st.session_state.personal[st.session_state.personal["Estado"] == "Activo"]["Nombre Completo"].tolist()
            with st.form("form_retiro"):
                persona_ret = st.selectbox("Colaborador activo:", activos_lista)
                fecha_corte_ret = st.date_input("Fecha de corte / liquidación:", value=date(2026, 11, 25))
                st.text_input("Motivo de retiro (opcional):", placeholder="Ej. Fin de reemplazo")

                if st.form_submit_button("🔒 Marcar como Liquidado / Retirado", use_container_width=True):
                    idx = st.session_state.personal[
                        st.session_state.personal["Nombre Completo"] == persona_ret
                    ].index
                    st.session_state.personal.loc[idx, "Estado"] = "Liquidado / Retirado"
                    st.session_state.personal.loc[idx, "Fecha Corte"] = fecha_corte_ret.strftime("%Y-%m-%d")
                    st.warning(f"🔒 {persona_ret} retirado con corte al {fecha_corte_ret}.")
                    st.rerun()

    st.markdown("**Base Maestra Actualizada (Activos e Históricos):**")
    st.dataframe(st.session_state.personal, use_container_width=True, hide_index=True)


def render_tab_tarifas():
    st.subheader("💲 Matriz Autónoma de Tarifas y Jornada Base")
    c_j1, c_j2 = st.columns([1, 2.5])
    with c_j1:
        st.session_state.jornada_base = st.number_input(
            "Jornada Ordinaria Base (horas/día):",
            min_value=4.0,
            max_value=12.0,
            value=float(st.session_state.jornada_base),
            step=0.5,
        )
    with c_j2:
        st.caption(
            "A partir de estas horas netas diarias (ya descontado el almuerzo), el sistema empieza a liquidar "
            "automáticamente **Horas Extras**. Los domingos y festivos aplican tarifa fija + recargo horario."
        )

    tarifas_editadas = st.data_editor(
        st.session_state.tarifas,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Tarifa Día ($)": st.column_config.NumberColumn(format="$ %d"),
            "Valor Hora Extra ($)": st.column_config.NumberColumn(format="$ %d"),
            "Tarifa Fija Dom/Fest ($)": st.column_config.NumberColumn(format="$ %d"),
            "Valor Hora Recargo Dom/Fest ($)": st.column_config.NumberColumn(format="$ %d"),
        },
        key="editor_tarifas_admin",
    )
    st.session_state.tarifas = tarifas_editadas


# ==============================================================================
# 9. CONTROLADOR PRINCIPAL DE SESIÓN Y VISTAS
# ==============================================================================
if not st.session_state.auth["autenticado"]:
    render_login()
else:
    with st.sidebar:
        st.markdown("### 🏢 SERGEM — Cañaveral")
        st.success(f"👤 **{st.session_state.auth['nombre']}**")
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            st.session_state.auth = {"autenticado": False, "rol": None, "nombre": None, "cedula": None}
            st.rerun()
        st.divider()
        st.info(
            "⏱️ **Cálculo automático:** El trabajador solo marca sus 4 horas; el sistema liquida extras, "
            "festivos y día de la semana."
        )
        st.caption("Desarrollado por **Gestión Logik**")

    st.markdown(
        """
        <div class="main-banner">
            <h1>🚚 Sistema de Control Horario, Rotación y Prefacturación — SERGEM & Cañaveral</h1>
            <p>Marcación diaria en campo, validación por punto, calendario automático de festivos y conciliación quincenal.</p>
            <div>
                <span class="badge-pill">⏱️ 4 Marcaciones Horarias</span>
                <span class="badge-pill">📅 Festivos Automáticos CO</span>
                <span class="badge-pill">📊 Estadística por Día</span>
                <span class="badge-pill">🔒 Acceso por Contraseña</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    rol_actual = st.session_state.auth["rol"]

    if rol_actual == "trabajador":
        render_vista_trabajador()

    elif rol_actual == "encargado_punto":
        render_vista_encargado_punto()

    elif rol_actual == "admin_sergem":
        st.markdown(
            '<div class="role-strip">🔐 <b>Panel Administración SERGEM</b> — Control total de prefacturación, validación de turnos, personal dinámico, tarifas y estadísticas.</div>',
            unsafe_allow_html=True,
        )
        render_kpis_gerenciales()
        st.write("")
        t1, t2, t3, t4, t5 = st.tabs(
            [
                "📊 1. Prefactura y Glosas",
                "✅ 2. Marcaciones Diarias",
                "📈 3. Estadística por Día",
                "👥 4. Personal y PIN",
                "💲 5. Tarifas y Jornada",
            ]
        )
        with t1:
            render_tab_prefactura(modo_auditor=False)
        with t2:
            render_vista_encargado_punto()
        with t3:
            render_tab_estadisticas_dias()
        with t4:
            render_tab_personal()
        with t5:
            render_tab_tarifas()

    elif rol_actual == "auditor_canaveral":
        st.markdown(
            '<div class="role-strip">👨‍💼 <b>Auditoría Cañaveral (Sr. Fabián)</b> — Revisión quincenal, glosas en columnas paralelas, respaldo de marcaciones y descarga oficial.</div>',
            unsafe_allow_html=True,
        )
        render_kpis_gerenciales()
        st.write("")
        ta1, ta2 = st.tabs(
            [
                "📊 1. Revisión de Prefactura y Autorización",
                "📋 2. Respaldo de Marcaciones Horarias (Solo Lectura)",
            ]
        )
        with ta1:
            render_tab_prefactura(modo_auditor=True)
        with ta2:
            st.subheader("📋 Bitácora Horaria Auditada")
            st.dataframe(st.session_state.registros_diarios, use_container_width=True, hide_index=True)

import io
from datetime import date
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA Y ESTILOS CORPORATIVOS
# ==============================================================================
st.set_page_config(
    page_title="SERGEM | Gestión Dinámica y Prefacturación Cañaveral",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    /* Fondo y tipografía general */
    .stApp {
        background-color: #F4F7FA;
    }
    /* Encabezado ejecutivo */
    .main-banner {
        background: linear-gradient(135deg, #0F2942 0%, #1E4E79 100%);
        padding: 1.4rem 1.8rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.2rem;
        box-shadow: 0 4px 12px rgba(15, 41, 66, 0.15);
    }
    .main-banner h1 {
        margin: 0;
        font-size: 1.55rem;
        font-weight: 700;
        letter-spacing: -0.3px;
        color: #FFFFFF;
    }
    .main-banner p {
        margin: 0.35rem 0 0 0;
        font-size: 0.92rem;
        color: #D9E6F2;
    }
    .badge-pill {
        display: inline-block;
        background-color: rgba(255, 255, 255, 0.16);
        border: 1px solid rgba(255, 255, 255, 0.3);
        color: #FFFFFF;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 0.78rem;
        font-weight: 600;
        margin-right: 6px;
        margin-top: 8px;
    }
    /* Tarjetas KPI */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 4px solid #1E4E79;
        padding: 0.9rem 1.1rem;
        border-radius: 10px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
    }
    div[data-testid="stMetric"] label {
        color: #4A5568 !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #0F2942 !important;
        font-size: 1.45rem !important;
        font-weight: 700 !important;
    }
    /* Pestañas */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #E2E8F0;
        padding: 6px;
        border-radius: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: 600;
        color: #2D3748;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0F2942 !important;
        color: #FFFFFF !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==============================================================================
# INICIALIZACIÓN DE DATOS FICTICIOS REALISTAS (SERGEM - CAÑAVERAL)
# ==============================================================================
def init_state():
    if "tarifas" not in st.session_state:
        st.session_state.tarifas = pd.DataFrame(
            [
                {"Sede": "Cali Norte", "Vehículo": "Moto", "Tarifa Día ($)": 138000, "Valor Hora Extra ($)": 11500, "Recargo Dominical ($)": 35000, "Bono Apoyo Miércoles ($)": 22000},
                {"Sede": "Cali Norte", "Vehículo": "Camioneta", "Tarifa Día ($)": 255000, "Valor Hora Extra ($)": 19200, "Recargo Dominical ($)": 55000, "Bono Apoyo Miércoles ($)": 35000},
                {"Sede": "Cali Sur", "Vehículo": "Moto", "Tarifa Día ($)": 138000, "Valor Hora Extra ($)": 11500, "Recargo Dominical ($)": 35000, "Bono Apoyo Miércoles ($)": 22000},
                {"Sede": "Cali Sur", "Vehículo": "Chopper", "Tarifa Día ($)": 172000, "Valor Hora Extra ($)": 14200, "Recargo Dominical ($)": 42000, "Bono Apoyo Miércoles ($)": 28000},
                {"Sede": "Yumbo", "Vehículo": "Camioneta", "Tarifa Día ($)": 268000, "Valor Hora Extra ($)": 20100, "Recargo Dominical ($)": 58000, "Bono Apoyo Miércoles ($)": 38000},
                {"Sede": "Yumbo", "Vehículo": "Moto", "Tarifa Día ($)": 144000, "Valor Hora Extra ($)": 12000, "Recargo Dominical ($)": 36000, "Bono Apoyo Miércoles ($)": 24000},
                {"Sede": "Palmira", "Vehículo": "Chopper", "Tarifa Día ($)": 179000, "Valor Hora Extra ($)": 14800, "Recargo Dominical ($)": 44000, "Bono Apoyo Miércoles ($)": 30000},
                {"Sede": "Jamundí", "Vehículo": "Moto", "Tarifa Día ($)": 146000, "Valor Hora Extra ($)": 12200, "Recargo Dominical ($)": 37000, "Bono Apoyo Miércoles ($)": 25000},
                {"Sede": "Tuluá", "Vehículo": "Camioneta", "Tarifa Día ($)": 282000, "Valor Hora Extra ($)": 21400, "Recargo Dominical ($)": 62000, "Bono Apoyo Miércoles ($)": 40000},
                {"Sede": "Buga", "Vehículo": "Moto", "Tarifa Día ($)": 149000, "Valor Hora Extra ($)": 12500, "Recargo Dominical ($)": 38000, "Bono Apoyo Miércoles ($)": 25000},
            ]
        )

    if "personal" not in st.session_state:
        st.session_state.personal = pd.DataFrame(
            [
                {"ID": "SG-101", "Cédula": "1.144.082.310", "Nombre Completo": "Carlos Andrés Mejía", "Sede": "Cali Norte", "Vehículo": "Camioneta", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-01", "Fecha Corte": ""},
                {"ID": "SG-102", "Cédula": "1.107.094.512", "Nombre Completo": "Jhonatan Stiven Ortiz", "Sede": "Cali Sur", "Vehículo": "Moto", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-01", "Fecha Corte": ""},
                {"ID": "SG-103", "Cédula": "1.085.312.889", "Nombre Completo": "Luis Fernando Caicedo", "Sede": "Cali Sur", "Vehículo": "Chopper", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-15", "Fecha Corte": ""},
                {"ID": "SG-104", "Cédula": "1.112.456.781", "Nombre Completo": "Diego Alejandro Rincón", "Sede": "Yumbo", "Vehículo": "Camioneta", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-01", "Fecha Corte": ""},
                {"ID": "SG-105", "Cédula": "1.006.234.990", "Nombre Completo": "Brayan Alexis Zapata", "Sede": "Yumbo", "Vehículo": "Moto", "Vinculación": "Prestación de Servicios", "Estado": "Activo", "Fecha Ingreso": "2026-10-03", "Fecha Corte": ""},
                {"ID": "SG-106", "Cédula": "1.113.678.214", "Nombre Completo": "Héctor Fabio Valencia", "Sede": "Palmira", "Vehículo": "Chopper", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-01", "Fecha Corte": ""},
                {"ID": "SG-107", "Cédula": "1.143.890.115", "Nombre Completo": "Kevin Santiago Mosquera", "Sede": "Jamundí", "Vehículo": "Moto", "Vinculación": "Prestación de Servicios", "Estado": "Activo", "Fecha Ingreso": "2026-11-02", "Fecha Corte": ""},
                {"ID": "SG-108", "Cédula": "1.116.254.908", "Nombre Completo": "Óscar Iván Restrepo", "Sede": "Tuluá", "Vehículo": "Camioneta", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-08-01", "Fecha Corte": ""},
                {"ID": "SG-109", "Cédula": "1.115.078.432", "Nombre Completo": "Sebastián Gómez Vélez", "Sede": "Buga", "Vehículo": "Moto", "Vinculación": "Fijo / Planta", "Estado": "Activo", "Fecha Ingreso": "2026-09-01", "Fecha Corte": ""},
                {"ID": "SG-110", "Cédula": "1.144.199.654", "Nombre Completo": "Andrés Felipe Quintero", "Sede": "Cali Norte", "Vehículo": "Moto", "Vinculación": "Prestación de Servicios", "Estado": "Activo", "Fecha Ingreso": "2026-11-16", "Fecha Corte": ""},
                {"ID": "SG-111", "Cédula": "1.061.782.340", "Nombre Completo": "Mauricio Paredes León", "Sede": "Palmira", "Vehículo": "Chopper", "Vinculación": "Prestación de Servicios", "Estado": "Liquidado / Retirado", "Fecha Ingreso": "2026-09-10", "Fecha Corte": "2026-11-12"},
            ]
        )

    if "prefactura_q1" not in st.session_state:
        st.session_state.prefactura_q1 = pd.DataFrame(
            [
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Cali Norte",
                    "Cédula": "1.144.082.310",
                    "Mensajero": "Carlos Andrés Mejía",
                    "Vehículo": "Camioneta",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,8,9,10,11,12,13,15",
                    "Cant. Días": 13,
                    "Horas Extras": 26.77,
                    "Dominicales": 2,
                    "Apoyo Miércoles": 2,
                    "Valor Liquidado ($)": 3982984,
                    "Observación Cañaveral (Fabián)": "Validar 2,5 HE del día 08 en ruta norte.",
                    "Respuesta / Ajuste SERGEM": "Soportado con planilla firmada de despacho #408.",
                    "Estado": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Cali Sur",
                    "Cédula": "1.107.094.512",
                    "Mensajero": "Jhonatan Stiven Ortiz",
                    "Vehículo": "Moto",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,7,8,9,10,11,12,13,14",
                    "Cant. Días": 14,
                    "Horas Extras": 12.50,
                    "Dominicales": 1,
                    "Apoyo Miércoles": 2,
                    "Valor Liquidado ($)": 2154750,
                    "Observación Cañaveral (Fabián)": "OK conforme.",
                    "Respuesta / Ajuste SERGEM": "Sin novedad.",
                    "Estado": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Cali Sur",
                    "Cédula": "1.085.312.889",
                    "Mensajero": "Luis Fernando Caicedo",
                    "Vehículo": "Chopper",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "2,3,4,5,6,8,9,10,11,12,13,15",
                    "Cant. Días": 12,
                    "Horas Extras": 18.25,
                    "Dominicales": 1,
                    "Apoyo Miércoles": 2,
                    "Valor Liquidado ($)": 2421150,
                    "Observación Cañaveral (Fabián)": "Revisar apoyo miércoles día 11.",
                    "Respuesta / Ajuste SERGEM": "Se anexa soporte de refuerzo pedido por jefe de bodega.",
                    "Estado": "Ajustado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Yumbo",
                    "Cédula": "1.112.456.781",
                    "Mensajero": "Diego Alejandro Rincón",
                    "Vehículo": "Camioneta",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,8,9,10,11,12,13,14,15",
                    "Cant. Días": 14,
                    "Horas Extras": 21.40,
                    "Dominicales": 2,
                    "Apoyo Miércoles": 2,
                    "Valor Liquidado ($)": 4374140,
                    "Observación Cañaveral (Fabián)": "",
                    "Respuesta / Ajuste SERGEM": "",
                    "Estado": "Pendiente",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Yumbo",
                    "Cédula": "1.006.234.990",
                    "Mensajero": "Brayan Alexis Zapata",
                    "Vehículo": "Moto",
                    "Vinculación": "Prestación de Servicios",
                    "Días Trabajados": "3,4,5,10,11,12",
                    "Cant. Días": 6,
                    "Horas Extras": 6.50,
                    "Dominicales": 0,
                    "Apoyo Miércoles": 2,
                    "Valor Liquidado ($)": 990000,
                    "Observación Cañaveral (Fabián)": "Confirmar si cubrió incapacidad en Yumbo.",
                    "Respuesta / Ajuste SERGEM": "Sí, reemplazo por incapacidad de ruta Yumbo-02.",
                    "Estado": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Palmira",
                    "Cédula": "1.113.678.214",
                    "Mensajero": "Héctor Fabio Valencia",
                    "Vehículo": "Chopper",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,8,9,10,11,12,13,15",
                    "Cant. Días": 13,
                    "Horas Extras": 15.80,
                    "Dominicales": 2,
                    "Apoyo Miércoles": 1,
                    "Valor Liquidado ($)": 2678840,
                    "Observación Cañaveral (Fabián)": "",
                    "Respuesta / Ajuste SERGEM": "",
                    "Estado": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Jamundí",
                    "Cédula": "1.143.890.115",
                    "Mensajero": "Kevin Santiago Mosquera",
                    "Vehículo": "Moto",
                    "Vinculación": "Prestación de Servicios",
                    "Días Trabajados": "2,4,5,9,11,12,14",
                    "Cant. Días": 7,
                    "Horas Extras": 8.25,
                    "Dominicales": 0,
                    "Apoyo Miércoles": 2,
                    "Valor Liquidado ($)": 1172650,
                    "Observación Cañaveral (Fabián)": "Pendiente validar 1,25 HE del día 14.",
                    "Respuesta / Ajuste SERGEM": "",
                    "Estado": "Con Observación",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Tuluá",
                    "Cédula": "1.116.254.908",
                    "Mensajero": "Óscar Iván Restrepo",
                    "Vehículo": "Camioneta",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,8,9,10,11,12,13,14,15",
                    "Cant. Días": 14,
                    "Horas Extras": 19.60,
                    "Dominicales": 2,
                    "Apoyo Miércoles": 2,
                    "Valor Liquidado ($)": 4571440,
                    "Observación Cañaveral (Fabián)": "OK conforme.",
                    "Respuesta / Ajuste SERGEM": "Sin novedad.",
                    "Estado": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Buga",
                    "Cédula": "1.115.078.432",
                    "Mensajero": "Sebastián Gómez Vélez",
                    "Vehículo": "Moto",
                    "Vinculación": "Fijo / Planta",
                    "Días Trabajados": "1,2,3,4,5,6,8,9,10,11,12,13",
                    "Cant. Días": 12,
                    "Horas Extras": 10.00,
                    "Dominicales": 1,
                    "Apoyo Miércoles": 2,
                    "Valor Liquidado ($)": 2001000,
                    "Observación Cañaveral (Fabián)": "",
                    "Respuesta / Ajuste SERGEM": "",
                    "Estado": "Autorizado",
                },
                {
                    "Quincena": "01-15 Nov 2026",
                    "Sede": "Palmira",
                    "Cédula": "1.061.782.340",
                    "Mensajero": "Mauricio Paredes León (Retirado 12-Nov)",
                    "Vehículo": "Chopper",
                    "Vinculación": "Prestación de Servicios",
                    "Días Trabajados": "1,2,3,4,5,8,9,10,11,12",
                    "Cant. Días": 10,
                    "Horas Extras": 4.50,
                    "Dominicales": 1,
                    "Apoyo Miércoles": 2,
                    "Valor Liquidado ($)": 1960600,
                    "Observación Cañaveral (Fabián)": "Liquidación proporcional hasta fecha de corte 12/11.",
                    "Respuesta / Ajuste SERGEM": "Confirmado, solo se cobran los 10 días laborados antes del retiro.",
                    "Estado": "Autorizado",
                },
            ]
        )

    if "registros_diarios" not in st.session_state:
        st.session_state.registros_diarios = pd.DataFrame(
            [
                {"Fecha": "2026-11-16", "Día": 16, "Sede": "Cali Norte", "Mensajero": "Carlos Andrés Mejía", "Vehículo": "Camioneta", "Asistió": "Sí", "Horas Extras": 2.25, "Dominical": "No", "Apoyo Miércoles": "No", "Novedad": "Ruta completa sin novedad"},
                {"Fecha": "2026-11-16", "Día": 16, "Sede": "Cali Norte", "Mensajero": "Andrés Felipe Quintero", "Vehículo": "Moto", "Asistió": "Sí", "Horas Extras": 1.50, "Dominical": "No", "Apoyo Miércoles": "No", "Novedad": "Ingreso nuevo activo desde hoy"},
                {"Fecha": "2026-11-18", "Día": 18, "Sede": "Yumbo", "Mensajero": "Brayan Alexis Zapata", "Vehículo": "Moto", "Asistió": "Sí", "Horas Extras": 0.00, "Dominical": "No", "Apoyo Miércoles": "Sí", "Novedad": "Refuerzo despacho miércoles"},
            ]
        )


init_state()


# ==============================================================================
# FUNCIÓN GENERADORA DE EXCEL OFICIAL (FORMATO EXACTO CAÑAVERAL)
# ==============================================================================
def generar_excel_prefactura(df_prefactura: pd.DataFrame, df_diario: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_prefactura.to_excel(writer, index=False, sheet_name="Prefactura_Quincenal", startrow=3)
        df_diario.to_excel(writer, index=False, sheet_name="Respaldo_Auditoria_Diaria", startrow=2)

        wb = writer.book
        ws = wb["Prefactura_Quincenal"]

        # Encabezado institucional en Excel
        ws.merge_cells("A1:P1")
        ws["A1"] = "SERGEM S.A.S. — PREFACTURA DE OPERACIÓN LOGÍSTICA CAÑAVERAL"
        ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[1].height = 28

        ws.merge_cells("A2:P2")
        ws["A2"] = f"Generado automáticamente el {date.today().strftime('%d/%m/%Y')}  |  Control de Glosas y Autorización de Facturación"
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

        # Formatear encabezados fila 4
        for col_idx in range(1, len(df_prefactura.columns) + 1):
            cell = ws.cell(row=4, column=col_idx)
            cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            cell.fill = audit_fill if col_idx >= 13 else header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[4].height = 28

        # Formatear filas de datos
        for row_idx in range(5, 5 + len(df_prefactura)):
            estado_val = str(ws.cell(row=row_idx, column=16).value)
            for col_idx in range(1, len(df_prefactura.columns) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = Font(name="Calibri", size=10)
                cell.border = thin_border
                if col_idx == 12:
                    cell.number_format = '"$"#,##0'
                    cell.font = Font(name="Calibri", size=10, bold=True)
                elif col_idx == 9:
                    cell.number_format = "0.00"
                elif col_idx == 16:
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

        # Fila Total al final
        total_row = 5 + len(df_prefactura)
        ws.cell(row=total_row, column=11, value="TOTAL PREFACTURA:").font = Font(bold=True)
        total_cell = ws.cell(row=total_row, column=12, value=f"=SUM(L5:L{total_row - 1})")
        total_cell.font = Font(bold=True, size=11, color="0F2942")
        total_cell.number_format = '"$"#,##0'
        total_cell.fill = PatternFill(start_color="EBF8FF", end_color="EBF8FF", fill_type="solid")

        # Autoajuste de columnas
        for col_idx in range(1, len(df_prefactura.columns) + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 18
        ws.column_dimensions["D"].width = 28
        ws.column_dimensions["G"].width = 26
        ws.column_dimensions["M"].width = 34
        ws.column_dimensions["N"].width = 34

    return output.getvalue()


# ==============================================================================
# BARRA LATERAL (PERFIL DE ACCESO Y RESUMEN DE VALOR)
# ==============================================================================
with st.sidebar:
    st.markdown("### 🏢 SERGEM — Cañaveral")
    rol_activo = st.radio(
        "Simular vista de usuario:",
        [
            "👩‍💼 Facturación / Gerencia SERGEM",
            "👨‍💼 Auditoría Cañaveral (Sr. Fabián)",
            "📱 Supervisor Operativo (Celular)",
        ],
    )
    st.divider()
    st.markdown("**Ventajas clave del sistema:**")
    st.success(
        "✅ **Cero carga para Cañaveral:** No digitan nada; reciben el Excel exacto y revisan glosas."
    )
    st.info(
        "🔄 **Rotación sin listas fijas:** Altas y retiros diarios sin alterar quincenas pasadas."
    )
    st.warning(
        "⚡ **Cierre en 1 clic:** Días separados por comas (1,2,3...) y horas extras decimales."
    )
    st.caption("Prototipo funcional desarrollado por **Gestión Logik**")


# ==============================================================================
# BANNER PRINCIPAL
# ==============================================================================
st.markdown(
    """
    <div class="main-banner">
        <h1>🚚 Plataforma Integral de Prefacturación, Rotación y Auditoría — SERGEM & Cañaveral</h1>
        <p>Control operativo diario en tiempo real, liquidación automática por sede/vehículo y conciliación quincenal sin reprocesos.</p>
        <div>
            <span class="badge-pill">📍 7 Municipios Activos</span>
            <span class="badge-pill">🏍️ Moto • Camioneta • Chopper</span>
            <span class="badge-pill">📅 Cortes 01-15 y 16-31</span>
            <span class="badge-pill">🔒 Trazabilidad de Glosas</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Contenedor superior para KPIs dinámicos
kpi_container = st.container()

# ==============================================================================
# PESTAÑAS PRINCIPALES DE LA APLICACIÓN
# ==============================================================================
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📊 1. Prefactura y Control de Glosas",
        "📱 2. Registro Diario (Vista Celular)",
        "👥 3. Maestro Dinámico de Personal",
        "💲 4. Matriz Autónoma de Tarifas",
    ]
)

# ------------------------------------------------------------------------------
# TAB 1: PREFACTURA QUINCENAL Y AUDITORÍA CON CAÑAVERAL
# ------------------------------------------------------------------------------
with tab1:
    col_f1, col_f2, col_f3 = st.columns([1.2, 1.2, 2])
    with col_f1:
        filtro_sede = st.selectbox(
            "Filtrar por Sede / Municipio:",
            ["Todas las sedes"] + sorted(st.session_state.prefactura_q1["Sede"].unique().tolist()),
        )
    with col_f2:
        filtro_estado = st.selectbox(
            "Filtrar por Estado de Auditoría:",
            ["Todos", "Autorizado", "Ajustado", "Con Observación", "Pendiente"],
        )
    with col_f3:
        st.caption(
            "💡 **Interactivo en vivo:** Edita directamente las columnas **Observación Cañaveral (Fabián)**, "
            "**Respuesta / Ajuste SERGEM** o cambia el **Estado** a *Autorizado* para ver cómo se actualizan los totales arriba."
        )

    df_view = st.session_state.prefactura_q1.copy()
    if filtro_sede != "Todas las sedes":
        df_view = df_view[df_view["Sede"] == filtro_sede]
    if filtro_estado != "Todos":
        df_view = df_view[df_view["Estado"] == filtro_estado]

    edited_df = st.data_editor(
        df_view,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Valor Liquidado ($)": st.column_config.NumberColumn(format="$ %d"),
            "Horas Extras": st.column_config.NumberColumn(format="%.2f"),
            "Estado": st.column_config.SelectboxColumn(
                "Estado Auditoría",
                options=["Pendiente", "Con Observación", "Ajustado", "Autorizado"],
                required=True,
            ),
            "Observación Cañaveral (Fabián)": st.column_config.TextColumn(width="large"),
            "Respuesta / Ajuste SERGEM": st.column_config.TextColumn(width="large"),
        },
        disabled=[
            "Quincena",
            "Sede",
            "Cédula",
            "Mensajero",
            "Vehículo",
            "Vinculación",
            "Días Trabajados",
            "Cant. Días",
            "Horas Extras",
            "Dominicales",
            "Apoyo Miércoles",
            "Valor Liquidado ($)",
        ],
        key="editor_prefactura",
    )

    # Sincronizar cambios si no hay filtro activo
    if filtro_sede == "Todas las sedes" and filtro_estado == "Todos":
        st.session_state.prefactura_q1 = edited_df.copy()

    st.divider()
    col_btn1, col_btn2 = st.columns([1.4, 2.6])
    with col_btn1:
        excel_bytes = generar_excel_prefactura(
            st.session_state.prefactura_q1, st.session_state.registros_diarios
        )
        st.download_button(
            label="📥 Descargar Excel Oficial Cañaveral (.xlsx)",
            data=excel_bytes,
            file_name="PREFACTURA_SERGEM_CANAVERAL_Q1_NOV2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    with col_btn2:
        st.info(
            "El archivo descargado conserva **exactamente** la agrupación quincenal (`01-15` / `16-31`), "
            "los días trabajados separados por comas (`1,2,3...`), decimales exactos en horas extras y una segunda hoja "
            "con el respaldo diario para auditoría."
        )

# ------------------------------------------------------------------------------
# LLENAR KPIs SUPERIORES CON LOS DATOS ACTUALIZADOS EN VIVO
# ------------------------------------------------------------------------------
df_totales = st.session_state.prefactura_q1
total_presentado = df_totales["Valor Liquidado ($)"].sum()
total_autorizado = df_totales[df_totales["Estado"] == "Autorizado"]["Valor Liquidado ($)"].sum()
total_revision = df_totales[df_totales["Estado"] != "Autorizado"]["Valor Liquidado ($)"].sum()
activos_count = len(st.session_state.personal[st.session_state.personal["Estado"] == "Activo"])

with kpi_container:
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("💰 Total Prefactura Quincenal", f"${total_presentado:,.0f}")
    k2.metric("✅ Autorizado para Facturar", f"${total_autorizado:,.0f}", f"{(total_autorizado/total_presentado)*100:.1f}% aprobado")
    k3.metric("🔍 En Revisión / Glosas", f"${total_revision:,.0f}", f"{len(df_totales[df_totales['Estado'] != 'Autorizado'])} ítems")
    k4.metric("👥 Personal Activo Hoy", f"{activos_count} colaboradores", "Sin listas rígidas")

# ------------------------------------------------------------------------------
# TAB 2: REGISTRO DIARIO OPERATIVO (PARA SUPERVISION EN CELULAR O PC)
# ------------------------------------------------------------------------------
with tab2:
    st.subheader("📱 Reporte Diario de Operación (Menos de 5 minutos al día)")
    st.caption(
        "El supervisor elige la sede y marca novedades del día sobre el personal activo. "
        "Cañaveral no tiene que digitar nada y SERGEM queda con respaldo diario."
    )

    col_form, col_hist = st.columns([1.15, 1.85])

    with col_form:
        with st.form("form_registro_diario", clear_on_submit=True):
            fecha_reg = st.date_input("Fecha de operación:", value=date(2026, 11, 25))
            activos_df = st.session_state.personal[st.session_state.personal["Estado"] == "Activo"]
            mensajero_sel = st.selectbox(
                "Mensajero Activo:",
                activos_df["Nombre Completo"].tolist(),
            )
            info_msg = activos_df[activos_df["Nombre Completo"] == mensajero_sel].iloc[0]
            st.caption(
                f"📍 Sede: **{info_msg['Sede']}** | 🚚 Vehículo: **{info_msg['Vehículo']}** | 📋 {info_msg['Vinculación']}"
            )

            c_a, c_b = st.columns(2)
            with c_a:
                horas_ext = st.number_input("Horas Extras (decimales):", min_value=0.0, max_value=12.0, value=0.0, step=0.25)
                es_dominical = st.selectbox("¿Turno Dominical?", ["No", "Sí"])
            with c_b:
                es_miercoles = st.selectbox("¿Apoyo Miércoles?", ["No", "Sí"])
                novedad_txt = st.text_input("Observación / Soporte:", placeholder="Ej. Ruta completa o refuerzo")

            submitted_dia = st.form_submit_button("✅ Guardar Novedad del Día", use_container_width=True, type="primary")
            if submitted_dia:
                nuevo_reg = pd.DataFrame(
                    [
                        {
                            "Fecha": fecha_reg.strftime("%Y-%m-%d"),
                            "Día": fecha_reg.day,
                            "Sede": info_msg["Sede"],
                            "Mensajero": mensajero_sel,
                            "Vehículo": info_msg["Vehículo"],
                            "Asistió": "Sí",
                            "Horas Extras": horas_ext,
                            "Dominical": es_dominical,
                            "Apoyo Miércoles": es_miercoles,
                            "Novedad": novedad_txt if novedad_txt else "Asistencia normal",
                        }
                    ]
                )
                st.session_state.registros_diarios = pd.concat(
                    [nuevo_reg, st.session_state.registros_diarios], ignore_index=True
                )
                st.toast(f"Registro guardado para {mensajero_sel}", icon="✅")

    with col_hist:
        st.markdown("**Últimos registros diarios capturados (Respaldo de Auditoría):**")
        st.dataframe(st.session_state.registros_diarios, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# TAB 3: MAESTRO DINÁMICO DE PERSONAL (ALTAS Y RETIROS SIN LISTAS FIJAS)
# ------------------------------------------------------------------------------
with tab3:
    st.subheader("👥 Gestión Dinámica de Personal (Fijo y Prestación de Servicios)")
    st.caption(
        "Ingresa nuevos colaboradores o liquida retiros con fecha de corte. "
        "Los retirados dejan de salir en los reportes futuros pero conservan intacto su historial."
    )

    col_alta, col_retiro = st.columns(2)

    with col_alta:
        with st.expander("➕ Registrar Nuevo Ingreso (Fijo o Servicios)", expanded=True):
            with st.form("form_nuevo_ingreso", clear_on_submit=True):
                c1, c2 = st.columns(2)
                cedula_in = c1.text_input("Cédula:", placeholder="Ej. 1.144.555.222")
                nombre_in = c2.text_input("Nombre Completo:", placeholder="Ej. Juan Camilo Torres")
                c3, c4 = st.columns(2)
                sede_in = c3.selectbox("Sede asignada:", sorted(st.session_state.tarifas["Sede"].unique().tolist()))
                veh_in = c4.selectbox("Vehículo:", ["Moto", "Camioneta", "Chopper"])
                c5, c6 = st.columns(2)
                vinc_in = c5.selectbox("Modalidad:", ["Fijo / Planta", "Prestación de Servicios"])
                fecha_in = c6.date_input("Fecha de Ingreso:", value=date(2026, 11, 25))

                if st.form_submit_button("💾 Activar Colaborador", use_container_width=True, type="primary"):
                    if cedula_in and nombre_in:
                        nuevo_id = f"SG-{101 + len(st.session_state.personal)}"
                        fila_nueva = pd.DataFrame(
                            [
                                {
                                    "ID": nuevo_id,
                                    "Cédula": cedula_in,
                                    "Nombre Completo": nombre_in,
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
                        st.success(f"✅ {nombre_in} quedó activo de inmediato en {sede_in}.")
                        st.rerun()
                    else:
                        st.error("Por favor ingresa cédula y nombre completo.")

    with col_retiro:
        with st.expander("🚪 Liquidar / Retirar Personal con Fecha de Corte", expanded=True):
            activos_lista = st.session_state.personal[st.session_state.personal["Estado"] == "Activo"]["Nombre Completo"].tolist()
            with st.form("form_retiro"):
                persona_ret = st.selectbox("Seleccionar colaborador activo:", activos_lista)
                fecha_corte_ret = st.date_input("Fecha de corte / liquidación:", value=date(2026, 11, 25))
                motivo_ret = st.text_input("Nota de retiro (opcional):", placeholder="Ej. Fin de reemplazo o retiro voluntario")

                if st.form_submit_button("🔒 Marcar como Liquidado / Retirado", use_container_width=True):
                    idx = st.session_state.personal[
                        st.session_state.personal["Nombre Completo"] == persona_ret
                    ].index
                    st.session_state.personal.loc[idx, "Estado"] = "Liquidado / Retirado"
                    st.session_state.personal.loc[idx, "Fecha Corte"] = fecha_corte_ret.strftime("%Y-%m-%d")
                    st.warning(f"🔒 {persona_ret} fue retirado con corte al {fecha_corte_ret}. Su historial anterior permanece intacto.")
                    st.rerun()

    st.markdown("**Base Maestra Actualizada (Activos e Históricos):**")
    st.dataframe(st.session_state.personal, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# TAB 4: MATRIZ AUTÓNOMA DE TARIFAS POR SEDE Y VEHÍCULO
# ------------------------------------------------------------------------------
with tab4:
    st.subheader("💲 Matriz Autónoma de Tarifas (100% Administrable por SERGEM)")
    st.caption(
        "Si cambia una tarifa, sube el valor de hora extra o abren un nuevo municipio, "
        "tu equipo administrativo lo actualiza aquí directamente sin depender de soporte técnico."
    )

    tarifas_editadas = st.data_editor(
        st.session_state.tarifas,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Tarifa Día ($)": st.column_config.NumberColumn(format="$ %d"),
            "Valor Hora Extra ($)": st.column_config.NumberColumn(format="$ %d"),
            "Recargo Dominical ($)": st.column_config.NumberColumn(format="$ %d"),
            "Bono Apoyo Miércoles ($)": st.column_config.NumberColumn(format="$ %d"),
        },
        key="editor_tarifas",
    )
    st.session_state.tarifas = tarifas_editadas

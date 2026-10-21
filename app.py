import streamlit as st
import pandas as pd
import os
import requests
from datetime import datetime
import urllib.parse
import time

st.set_page_config(page_title="LLEGANDO - Red Integral de Movilidad y Logística", page_icon="🚗", layout="wide")

# ==========================================
# SOPORTE PWA Y ESTILOS CSS
# ==========================================
st.markdown("""
    <link rel="manifest" href="data:application/manifest+json;charset=utf-8,{
      'name': 'LLEGANDO',
      'short_name': 'LLEGANDO',
      'start_url': '.',
      'display': 'standalone',
      'background_color': '#F8FAFC',
      'theme_color': '#1E3A8A',
      'icons': [{
        'src': 'https://cdn-icons-png.flaticon.com/512/3063/3063823.png',
        'sizes': '512x512',
        'type': 'image/png'
      }]
    }">
    <style>
        .main { background-color: #F8FAFC; }
        h1, h2, h3 { color: #0A2540; font-family: 'Helvetica Neue', sans-serif; }
        .stButton>button { background-color: #1E3A8A; color: white; border-radius: 6px; font-weight: 600; border: none; }
        .stButton>button:hover { background-color: #0F172A; color: white; }
        .cintillo-alerta {
            background-color: #0A2540; color: white; padding: 18px 24px;
            border-radius: 12px; box-shadow: 0 6px 20px rgba(10, 37, 64, 0.3);
            margin-bottom: 20px; border-left: 8px solid #3B82F6;
        }
        .banner-instalacion {
            background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
            color: white; padding: 20px; border-radius: 12px; margin-bottom: 25px;
        }
        .card-aliado {
            background-color: #FFFFFF; padding: 15px; border-radius: 8px;
            border-left: 5px solid #10B981; box-shadow: 0 2px 5px rgba(0,0,0,0.05); margin-bottom: 10px;
        }
    </style>
""", unsafe_allow_html=True)

DB_FILE = "base_datos_servicios.xlsx"
DB_PRESTADORES = "prestadores_autorizados.xlsx"
DB_LIQUIDACIONES = "liquidaciones_prestadores.xlsx"
CLAVE_ADMIN_SECRETA = "admin123"

@st.cache_data(ttl=3600)
def obtener_tasa_bcv():
    try:
        response = requests.get("https://ve.dolarapi.com/v1/dolares/oficial", timeout=5)
        if response.status_code == 200:
            return float(response.json().get("promedio", 36.50))
    except Exception:
        pass
    return 36.50

if "tasa_bcv" not in st.session_state:
    st.session_state.tasa_bcv = obtener_tasa_bcv()
if "porcentaje_prestador" not in st.session_state:
    st.session_state.porcentaje_prestador = 80.0

DATOS_PAGO_MOVIL = {
    "Banco": "BDV (0102)",
    "Telefono": "0412-3349463",
    "Cedula_RIF": "V-13609424",
    "Titular": "Servicios de Movilidad"
}

def cargar_datos():
    if os.path.exists(DB_FILE):
        df_temp = pd.read_excel(DB_FILE, dtype={"Ref_Pago_Movil": str, "Telefono_Cliente": str, "ID_Servicio": int})
        if "Pago_Verificado" not in df_temp.columns: df_temp["Pago_Verificado"] = "No"
        if "Liquidado" not in df_temp.columns: df_temp["Liquidado"] = "No"
        if "Valoracion" not in df_temp.columns: df_temp["Valoracion"] = "Sin valorar"
        if "Aliado_Comercial" not in df_temp.columns: df_temp["Aliado_Comercial"] = "Particular"
        return df_temp
    else:
        df_inicial = pd.DataFrame(columns=[
            "ID_Servicio", "Fecha_Hora", "Tipo_Servicio", "Cliente", 
            "Telefono_Cliente", "Origen", "Destino", "Conductor", 
            "Monto_USD", "Metodo_Pago", "Ref_Pago_Movil", "Pago_Verificado", "Estado", "Liquidado", "Valoracion", "Notas", "Aliado_Comercial"
        ])
        df_inicial.to_excel(DB_def cargar_datos():
    if os.path.exists(DB_FILE):
        df_temp = pd.read_excel(DB_FILE, dtype={"Ref_Pago_Movil": str, "Telefono_Cliente": str, "ID_Servicio": int})
        if "Pago_Verificado" not in df_temp.columns: df_temp["Pago_Verificado"] = "No"
        if "Liquidado" not in df_temp.columns: df_temp["Liquidado"] = "No"
        if "Valoracion" not in df_temp.columns: df_temp["Valoracion"] = "Sin valorar"
        if "Aliado_Comercial" not in df_temp.columns: df_temp["Aliado_Comercial"] = "Particular"
        if "Capture_Pago" not in df_temp.columns: df_temp["Capture_Pago"] = ""
        return df_temp
    else:
        df_inicial = pd.DataFrame(columns=[
            "ID_Servicio", "Fecha_Hora", "Tipo_Servicio", "Cliente", 
            "Telefono_Cliente", "Origen", "Destino", "Conductor", 
            "Monto_USD", "Metodo_Pago", "Ref_Pago_Movil", "Pago_Verificado", "Estado", "Liquidado", "Valoracion", "Notas", "Aliado_Comercial", "Capture_Pago"
        ])
        df_inicial.to_excel(DB_FILE, index=False)
        return df_inicial

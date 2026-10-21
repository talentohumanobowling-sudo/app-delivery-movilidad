import streamlit as st
import pandas as pd
import os
import requests
from datetime import datetime
import urllib.parse
import time

st.set_page_config(page_title="LLEGANDO - Red Integral de Movilidad y Logística", page_icon="🚗", layout="wide")

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
    columnas_necesarias = [
        "ID_Servicio", "Fecha_Hora", "Tipo_Servicio", "Cliente", 
        "Telefono_Cliente", "Origen", "Destino", "Conductor", 
        "Monto_USD", "Metodo_Pago", "Ref_Pago_Movil", 
        "Pago_Verificado", "Estado", "Liquidado", "Valoracion", "Notas", "Aliado_Comercial"
    ]
    if os.path.exists(DB_FILE):
        try:
            df_temp = pd.read_excel(DB_FILE, dtype=str)
        except Exception:
            df_temp = pd.DataFrame()
            
        for col in columnas_necesarias:
            if col not in df_temp.columns:
                df_temp[col] = ""
                
        df_temp = df_temp[columnas_necesarias]
        df_temp["ID_Servicio"] = pd.to_numeric(df_temp["ID_Servicio"], errors='coerce').fillna(1001).astype(int)
        df_temp["Monto_USD"] = pd.to_numeric(df_temp["Monto_USD"], errors='coerce').fillna(0.0)
        
        for col in columnas_necesarias:
            if col not in ["ID_Servicio", "Monto_USD"]:
                df_temp[col] = df_temp[col].astype(str).fillna("").replace("nan", "").str.strip()
        return df_temp
    else:
        df_inicial = pd.DataFrame(columns=columnas_necesarias)
        df_inicial.to_excel(DB_FILE, index=False)
        return df_inicial

def guardar_datos(df_a_guardar):
    df_a_guardar.to_excel(DB_FILE, index=False)

def cargar_prestadores():
    if os.path.exists(DB_PRESTADORES):
        return pd.read_excel(DB_PRESTADORES, dtype=str).fillna("")
    else:
        df_p = pd.DataFrame(columns=["Nombre", "Telefono", "Vehiculo", "Marca", "Modelo", "Placa", "Foto_Path", "Pago_Afiliacion_5USD", "Pago_Mensualidad_5USD", "Estatus", "Saldo_Wallet"])
        df_p.to_excel(DB_PRESTADORES, index=False)
        return df_p

def guardar_prestadores(df_p):
    df_p.to_excel(DB_PRESTADORES, index=False)

def cargar_liquidaciones():
    if os.path.exists(DB_LIQUIDACIONES):
        return pd.read_excel(DB_LIQUIDACIONES, dtype=str).fillna("")
    else:
        df_l = pd.DataFrame(columns=["ID_Liquidacion", "Fecha", "Prestador", "Monto_USD", "Referencia_Pago", "Estatus_Confirmacion"])
        df_l.to_excel(DB_LIQUIDACIONES, index=False)
        return df_l

def guardar_liquidaciones(df_l):
    df_l.to_excel(DB_LIQUIDACIONES, index=False)

df = cargar_datos()
df_prestadores = cargar_prestadores()
df_liquidaciones = cargar_liquidaciones()

if os.path.exists("logo.jpg"):
    st.sidebar.image("logo.jpg", use_container_width=True)
elif os.path.exists("logo.png"):
    st.sidebar.image("logo.png", use_container_width=True)
else:
    st.sidebar.markdown("### LLEGANDO")

st.sidebar.markdown("---")
rol = st.sidebar.radio("Selecciona tu Panel:", [
    "Portal de Cliente", 
    "Portal de Prestador", 
    "Portal de Aliado Comercial",
    "Panel Gerencial (Admin)"
])
st.sidebar.markdown("---")
st.sidebar.metric("Tasa BCV Oficial", f"Bs. {st.session_state.tasa_bcv:,.2f}")

if "Portal de Cliente" in rol:
    st.markdown("<h2>LLEGANDO - Solicitud y Estatus</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    tab_solicitar, tab_consultar = st.tabs(["Solicitar Nuevo Servicio", "Consultar Mis Viajes"])
    
    with tab_solicitar:
        with st.form("form_cliente"):
            col1, col2 = st.columns(2)
            with col1:
                nombre = st.text_input("Tu Nombre Completo *")
                telefono = st.text_input("Tu Numero de Telefono *")
                tipo_serv = st.selectbox("Tipo de Servicio *", ["Delivery", "Mototaxi", "Taxi / Carro", "Flete"])
            with col2:
                origen = st.text_input("Direccion de Origen *")
                destino = st.text_input("Direccion de Destino *")
                metodo_pago = st.selectbox("Metodo de Pago Preferido", ["Pago Movil", "Efectivo", "Zelle"])
            
            notas = st.text_area("Detalles adicionales (Opcional)")
            tarifas_base = {"Delivery": 3.0, "Mototaxi": 2.5, "Taxi / Carro": 5.0, "Flete": 15.0}
            tarifa_sugerida = tarifas_base.get(tipo_serv, 3.0)
            st.info(f"Tarifa estimada sugerida para {tipo_serv}: **${tarifa_sugerida:.2f} USD**")
            
            if st.form_submit_button("Enviar Solicitud Ahora"):
                if nombre and telefono and origen and destino:
                    nuevo_id = 1001 if len(df) == 0 else int(df["ID_Servicio"].max()) + 1
                    nuevo_registro = {
                        "ID_Servicio": nuevo_id, 
                        "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Tipo_Servicio": str(tipo_serv), 
                        "Cliente": str(nombre), 
                        "Telefono_Cliente": str(telefono).strip(),
                        "Origen": str(origen), 
                        "Destino": str(destino), 
                        "Conductor": "Por Asignar", 
                        "Monto_USD": float(tarifa_sugerida),
                        "Metodo_Pago": str(metodo_pago), 
                        "Ref_Pago_Movil": "", 
                        "Pago_Verificado": "No",
                        "Estado": "Solicitado", 
                        "Liquidado": "No", 
                        "Valoracion": "Sin valorar", 
                        "Notas": str(notas), 
                        "Aliado_Comercial": "Particular"
                    }
                    df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
                    guardar_datos(df)
                    st.success(f"¡Solicitud enviada con éxito! Tu número de ID es #{nuevo_id}.")
                else:
                    st.error("Completa los campos obligatorios (*).")
                    
    with tab_consultar:
        st.subheader("Mis Servicios Activos e Historial")
        telefono_busqueda = st.text_input("Ingresa tu número de teléfono registrado:").strip()
        
        if telefono_busqueda:
            # Búsqueda flexible ignorando ceros a la izquierda para emparejar con Excel
            tlf_busqueda_limpio = telefono_busqueda.lstrip('0')
            
            def coincide_telefono(val):
                val_str = str(val).strip()
                return val_str == telefono_busqueda or val_str.lstrip('0') == tlf_busqueda_limpio

            resultados = df[df["Telefono_Cliente"].apply(coincide_telefono)]
            
            if len(resultados) > 0:
                resultados = resultados.sort_values(by="ID_Servicio", ascending=False)
                for idx_res, serv in resultados.iterrows():
                    monto_usd = float(serv['Monto_USD'])
                    monto_ves = monto_usd * st.session_state.tasa_bcv
                    estatus_s = str(serv['Estado']).strip()
                    pago_v = str(serv.get('Pago_Verificado', 'No')).strip()
                    
                    with st.expander(f"Servicio #{serv['ID_Servicio']} - {serv['Tipo_Servicio']} [{estatus_s}]", expanded=True):
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"**Cliente:** {serv['Cliente']}")
                            st.markdown(f"**Origen:** {serv['Origen']}")
                            st.markdown(f"**Destino:** {serv['Destino']}")
                        with c2:
                            st.markdown(f"**Estado:** `{estatus_s}`")
                            st.markdown(f"**Tarifa:** **${monto_usd:,.2f} USD** (Bs. {monto_ves:,.2f})")
                        
                        if estatus_s == "Cancelado":
                            st.error("Este servicio fue cancelado.")
                        else:
                            ref_actual = str(serv['Ref_Pago_Movil']) if pd.notna(serv['Ref_Pago_Movil']) else ""
                            if str(serv['Metodo_Pago']).strip().lower() in ["pago movil", "pago móvil"] and monto_usd > 0:
                                st.markdown("---")
                                st.info("Datos para Pago Movil:")
                                st.warning(f"Monto exacto: Bs. {monto_ves:,.2f}")
                                st.write(f"Banco: {DATOS_PAGO_MOVIL['Banco']} | Telefono: {DATOS_PAGO_MOVIL['Telefono']}")
                                st.write(f"Cedula/RIF: {DATOS_PAGO_MOVIL['Cedula_RIF']} | Titular: {DATOS_PAGO_MOVIL['Titular']}")
                                
                                ref_ingresada = st.text_input("Ingresa los dígitos de tu referencia:", value=ref_actual, key=f"ref_{serv['ID_Servicio']}")
                                
                                if st.button(f"Reportar Referencia #{serv['ID_Servicio']}", key=f"btn_ref_{serv['ID_Servicio']}"):
                                    if ref_ingresada:
                                        filtro = df["ID_Servicio"] == int(serv['ID_Servicio'])
                                        df.loc[filtro, "Ref_Pago_Movil"] = str(ref_ingresada)
                                        guardar_datos(df)
                                        st.success("¡Referencia reportada con éxito!")
                                        st.rerun()
                                    else:
                                        st.error("Ingresa al menos los dígitos de la referencia.")
            else:
                st.warning("No se encontraron servicios con ese número exacto.")

elif "Portal de Prestador" in rol:
    st.markdown("<h2>Panel Profesional del Prestador</h2>", unsafe_allow_html=True)
    st.markdown("---")
    nom_prestador = st.text_input("Ingresa tu Nombre exacto autorizado:")
    if nom_prestador:
        prestador_valido = df_prestadores[(df_prestadores["Nombre"].str.lower() == nom_prestador.strip().lower()) & (df_prestadores["Estatus"] == "Autorizado")]
        if len(prestador_valido) == 0:
            st.warning("Tu cuenta no está autorizada o no existe.")
        else:
            st.success("Acceso Autorizado")
            servicios_cond = df[df["Conductor"].astype(str).str.strip().str.lower() == nom_prestador.strip().lower()]
            if len(servicios_cond) > 0:
                for idx, serv in servicios_cond.iterrows():
                    st.write(f"Servicio #{serv['ID_Servicio']} - Cliente: {serv['Cliente']} [{serv['Estado']}]")
            else:
                st.info("No tienes servicios asignados actualmente.")

elif "Portal de Aliado Comercial" in rol:
    st.markdown("<h2>🏬 Portal de Aliados Comerciales</h2>", unsafe_allow_html=True)
    st.markdown("---")
    nombre_aliado = st.text_input("Nombre de su Comercio Aliado:")
    if nombre_aliado:
        st.success(f"Bienvenido, comercio: **{nombre_aliado}**")

elif "Panel Gerencial" in rol:
    st.markdown("<h2>Panel Gerencial LLEGANDO</h2>", unsafe_allow_html=True)
    st.markdown("---")
    clave_ingresada = st.text_input("Contraseña de Administrador:", type="password")
    if clave_ingresada == CLAVE_ADMIN_SECRETA:
        st.success("Acceso Gerencial Autorizado")
        st.dataframe(df, use_container_width=True)
    elif clave_ingresada:
        st.error("Contraseña incorrecta.")

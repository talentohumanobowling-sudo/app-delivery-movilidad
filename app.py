import streamlit as st
import pandas as pd
import os
import requests
from datetime import datetime

st.set_page_config(page_title="App de Movilidad y Servicios", page_icon="🚖", layout="wide")

DB_FILE = "base_datos_servicios.xlsx"

# Función para obtener Tasa BCV automáticamente con respaldo manual
@st.cache_data(ttl=3600)
def obtener_tasa_bcv():
    try:
        response = requests.get("https://ve.dolarapi.com/v1/dolares/oficial", timeout=5)
        if response.status_code == 200:
            return float(response.json().get("promedio", 36.50))
    except Exception:
        pass
    return 36.50

# Configuración de Datos de Pago Móvil y Comisión Base
if "tasa_bcv" not in st.session_state:
    st.session_state.tasa_bcv = obtener_tasa_bcv()

if "porcentaje_prestador" not in st.session_state:
    st.session_state.porcentaje_prestador = 80.0  # El prestador se queda con el 80% y la empresa 20%

DATOS_PAGO_MOVIL = {
    "Banco": "BDV (0102)",
    "Telefono": "0412-3349463",
    "Cedula_RIF": "V-13609424",
    "Titular": "Servicios de Movilidad"
}

def cargar_datos():
    if os.path.exists(DB_FILE):
        df_temp = pd.read_excel(DB_FILE, dtype={"Ref_Pago_Movil": str, "Telefono_Cliente": str})
        return df_temp
    else:
        df_inicial = pd.DataFrame(columns=[
            "ID_Servicio", "Fecha_Hora", "Tipo_Servicio", "Cliente", 
            "Telefono_Cliente", "Origen", "Destino", "Conductor", 
            "Monto_USD", "Metodo_Pago", "Ref_Pago_Movil", "Estado", "Notas"
        ])
        df_inicial.to_excel(DB_FILE, index=False)
        return df_inicial

def guardar_datos(df):
    df["Ref_Pago_Movil"] = df["Ref_Pago_Movil"].astype(str)
    df.to_excel(DB_FILE, index=False)

df = cargar_datos()

# NAVEGACIÓN PRINCIPAL
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3097/3097180.png", width=90)
st.sidebar.title("Menú Principal")
rol = st.sidebar.radio("Selecciona tu Rol:", [
    "👤 Cliente (Solicitar y Consultar)", 
    "🚗 Prestador de Servicios", 
    "⚙️ Administrador"
])

# Tasa BCV visible en la barra lateral
st.sidebar.markdown("---")
st.sidebar.metric("💵 Tasa BCV del Día", f"Bs. {st.session_state.tasa_bcv:,.2f}")

# ==========================================
# 1. VISTA CLIENTE: SOLICITUD Y CONSULTA
# ==========================================
if rol == "👤 Cliente (Solicitar y Consultar)":
    st.markdown("<h2 style='text-align: center; color: #1E88E5;'>📱 Portal de Cliente</h2>", unsafe_allow_html=True)
    
    tab_solicitar, tab_consultar = st.tabs(["📝 Solicitar Nuevo Servicio", "🔍 Consultar Mi Viaje / Pagar"])
    
    with tab_solicitar:
        st.subheader("Solicita tu servicio en línea")
        with st.form("form_cliente"):
            col1, col2 = st.columns(2)
            with col1:
                nombre = st.text_input("Tu Nombre Completo *")
                telefono = st.text_input("Tu Número de Teléfono *")
                tipo_serv = st.selectbox("Tipo de Servicio *", ["Delivery", "Mototaxi", "Taxi / Carro", "Flete"])
            with col2:
                origen = st.text_input("Dirección de Origen / Recogida *")
                destino = st.text_input("Dirección de Destino / Entrega *")
                metodo_pago = st.selectbox("Método de Pago Preferido", ["Pago Móvil", "Efectivo", "Zelle"])
            
            notas = st.text_area("Detalles adicionales (Opcional)")
            
            if st.form_submit_button("🚀 Enviar Solicitud"):
                if nombre and telefono and origen and destino:
                    nuevo_id = 1001 if len(df) == 0 else int(pd.to_numeric(df["ID_Servicio"], errors='coerce').max()) + 1
                    nuevo_registro = {
                        "ID_Servicio": nuevo_id,
                        "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Tipo_Servicio": tipo_serv,
                        "Cliente": nombre,
                        "Telefono_Cliente": str(telefono),
                        "Origen": origen,
                        "Destino": destino,
                        "Conductor": "Por Asignar",
                        "Monto_USD": 0.0,
                        "Metodo_Pago": metodo_pago,
                        "Ref_Pago_Movil": "",
                        "Estado": "Solicitado",
                        "Notas": notas
                    }
                    df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
                    guardar_datos(df)
                    st.success(f"¡Solicitud enviada con éxito! Tu número de ID es **#{nuevo_id}**. Consérvalo para consultar el estado de tu viaje.")
                else:
                    st.error("Por favor completa los campos obligatorios (*).")
                    
    with tab_consultar:
        st.subheader("Consulta el Estado y Conductor de tu Servicio")
        if len(df) == 0:
            st.info("No hay solicitudes registradas.")
        else:
            id_busqueda = st.number_input("Ingresa tu ID de Servicio:", min_value=1001, step=1, value=1001)
            resultado = df[df["ID_Servicio"] == id_busqueda]
            
            if len(resultado) > 0:
                serv = resultado.iloc[0]
                monto_usd = float(serv['Monto_USD'])
                monto_ves = monto_usd * st.session_state.tasa_bcv
                
                st.success(f"### 📍 Servicio #{serv['ID_Servicio']}")
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"**👤 Cliente:** {serv['Cliente']}")
                    st.markdown(f"**🚘 Servicio:** {serv['Tipo_Servicio']}")
                    st.markdown(f"**📍 Origen:** {serv['Origen']}")
                    st.markdown(f"**🏁 Destino:** {serv['Destino']}")
                with c2:
                    st.markdown(f"**🚦 Estado del Viaje:** `{serv['Estado']}`")
                    st.markdown(f"**👨‍✈️ Conductor Asignado:** **{serv['Conductor']}**")
                    st.markdown(f"**💵 Tarifa:** **${monto_usd:,.2f} USD** (Bs. {monto_ves:,.2f})")
                
                if str(serv['Metodo_Pago']).strip().lower() == "pago móvil" and monto_usd > 0:
                    st.markdown("---")
                    st.info("📲 **Datos para Pago Móvil:**")
                    st.warning(f"💱 **Monto exacto a transferir:** **Bs. {monto_ves:,.2f}** *(Tasa BCV: {st.session_state.tasa_bcv})*")
                    col_p1, col_p2 = st.columns(2)
                    with col_p1:
                        st.write(f"🏦 **Banco:** {DATOS_PAGO_MOVIL['Banco']}")
                        st.write(f"📱 **Teléfono:** {DATOS_PAGO_MOVIL['Telefono']}")
                    with col_p2:
                        st.write(f"🪪 **Cédula / RIF:** {DATOS_PAGO_MOVIL['Cedula_RIF']}")
                        st.write(f"👤 **Titular:** {DATOS_PAGO_MOVIL['Titular']}")
                    
                    val_actual = "" if pd.isna(serv['Ref_Pago_Movil']) or str(serv['Ref_Pago_Movil']) == "nan" else str(serv['Ref_Pago_Movil'])
                    ref_ingresada = st.text_input("📝 Ingresa los dígitos de tu referencia bancaria:", value=val_actual)
                    if st.button("Confirmar y Reportar Pago"):
                        idx = df[df["ID_Servicio"] == id_busqueda].index
                        df.at[idx[0], "Ref_Pago_Movil"] = str(ref_ingresada)
                        guardar_datos(df)
                        st.success("¡Referencia guardada!")
                        st.rerun()
            else:
                st.warning("No se encontró ningún servicio con ese ID.")

# ==========================================
# 2. VISTA PRESTADOR DE SERVICIOS
# ==========================================
elif rol == "🚗 Prestador de Servicios":
    st.title("🚗 Panel del Prestador de Servicios")
    st.markdown("---")
    
    nom_prestador = st.text_input("Ingresa tu Nombre para acceder a tus servicios y finanzas:")
    
    if nom_prestador:
        servicios_cond = df[df["Conductor"].astype(str).str.contains(nom_prestador, case=False, na=False)]
        
        tab_viajes, tab_finanzas = st.tabs(["📋 Mis Viajes y Estatus", "💰 Mis Ganancias y Porcentaje"])
        
        with tab_viajes:
            if len(servicios_cond) > 0:
                st.subheader(f"Servicios Asignados a: {nom_prestador}")
                for idx, serv in servicios_cond.iterrows():
                    monto_usd = float(serv['Monto_USD'])
                    monto_ves = monto_

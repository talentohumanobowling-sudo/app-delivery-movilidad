import streamlit as st
import pandas as pd
import os
import requests
from datetime import datetime
import urllib.parse

st.set_page_config(page_title="App de Movilidad y Servicios", page_icon="🚖", layout="wide")

DB_FILE = "base_datos_servicios.xlsx"
DB_PRESTADORES = "prestadores_autorizados.xlsx"

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
        return pd.read_excel(DB_FILE, dtype={"Ref_Pago_Movil": str, "Telefono_Cliente": str})
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

def cargar_prestadores():
    if os.path.exists(DB_PRESTADORES):
        return pd.read_excel(DB_PRESTADORES, dtype=str)
    else:
        df_p = pd.DataFrame(columns=["Nombre", "Telefono", "Vehiculo", "Estatus"])
        df_p.to_excel(DB_PRESTADORES, index=False)
        return df_p

def guardar_prestadores(df_p):
    df_p.to_excel(DB_PRESTADORES, index=False)

df = cargar_datos()
df_prestadores = cargar_prestadores()

st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3097/3097180.png", width=90)
st.sidebar.title("Menú Principal")
rol = st.sidebar.radio("Selecciona tu Rol:", [
    "👤 Cliente (Solicitar y Consultar)", 
    "🚗 Prestador de Servicios", 
    "⚙️ Administrador"
])

st.sidebar.markdown("---")
st.sidebar.metric("💵 Tasa BCV del Día", f"Bs. {st.session_state.tasa_bcv:,.2f}")

# ==========================================
# 1. VISTA CLIENTE
# ==========================================
if rol == "👤 Cliente (Solicitar y Consultar)":
    st.markdown("<h2 style='color: #1E88E5;'>📱 Portal de Cliente</h2>", unsafe_allow_html=True)
    
    tab_solicitar, tab_consultar = st.tabs(["📝 Solicitar Nuevo Servicio", "🔍 Consultar Mi Viaje / Pagar"])
    
    with tab_solicitar:
        with st.form("form_cliente"):
            col1, col2 = st.columns(2)
            with col1:
                nombre = st.text_input("Tu Nombre Completo *")
                telefono = st.text_input("Tu Número de Teléfono (ej. 04141234567) *")
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
                    st.success(f"¡Solicitud enviada con éxito! Tu número de ID es **#{nuevo_id}**.")
                else:
                    st.error("Por favor completa los campos obligatorios (*).")
                    
    with tab_consultar:
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
                    st.markdown(f"**🚦 Estado:** `{serv['Estado']}`")
                    st.markdown(f"**💵 Tarifa:** **${monto_usd:,.2f} USD** (Bs. {monto_ves:,.2f})")
                
                if serv['Estado'] == "Cancelado":
                    st.error("⚠️ Este servicio ha sido cancelado.")
                else:
                    ref_actual = "" if pd.isna(serv['Ref_Pago_Movil']) or str(serv['Ref_Pago_Movil']) == "nan" else str(serv['Ref_Pago_Movil']).strip()
                    
                    if str(serv['Metodo_Pago']).strip().lower() == "pago móvil" and monto_usd > 0:
                        st.markdown("---")
                        st.info("📲 **Datos para Pago Móvil:**")
                        st.warning(f"💱 **Monto exacto:** **Bs. {monto_ves:,.2f}** *(Tasa BCV: {st.session_state.tasa_bcv})*")
                        
                        col_p1, col_p2 = st.columns(2)
                        with col_p1:
                            st.write(f"🏦 **Banco:** {DATOS_PAGO_MOVIL['Banco']}")
                            st.write(f"📱 **Teléfono:** {DATOS_PAGO_MOVIL['Telefono']}")
                        with col_p2:
                            st.write(f"🪪 **Cédula / RIF:** {DATOS_PAGO_MOVIL['Cedula_RIF']}")
                            st.write(f"👤 **Titular:** {DATOS_PAGO_MOVIL['Titular']}")
                        
                        ref_ingresada = st.text_input("📝 Ingresa los dígitos de tu referencia bancaria:", value=ref_actual)
                        if st.button("Confirmar y Reportar Pago"):
                            if ref_ingresada:
                                idx = df[df["ID_Servicio"] == id_busqueda].index
                                df.at[idx[0], "Ref_Pago_Movil"] = str(ref_ingresada)
                                guardar_datos(df)
                                st.success("¡Referencia de pago registrada con éxito!")
                                st.rerun()
                            else:
                                st.error("Ingresa la referencia de pago.")
                    
                    if ref_actual != "" or str(serv['Metodo_Pago']).strip().lower() != "pago móvil":
                        st.markdown("---")
                        st.success("✅ **¡Pago verificado / Registrado con éxito!**")
                        st.markdown(f"### 👨‍✈️ Tu Prestador de Servicio Asignado es: **{serv['Conductor']}**")
                        
                        info_cond = df_prestadores[df_prestadores["Nombre"].str.lower() == str(serv['Conductor']).lower()]
                        if len(info_cond) > 0:
                            tlf_cond = info_cond.iloc[0]["Telefono"].strip().replace(" ", "").replace("+", "")
                            msg_c = f"Hola {serv['Conductor']}, le escribo desde la app respecto a mi servicio #{serv['ID_Servicio']}."
                            url_wa_cond = f"https://wa.me/{tlf_cond}?text={urllib.parse.quote(msg_c)}"
                            st.markdown(f'<a href="{url_wa_cond}" target="_blank"><button style="background-color:#25D366; color:white; border:none; padding:10px 20px; border-radius:5px; cursor:

import streamlit as st
import pandas as pd
import os
import requests
from datetime import datetime
import urllib.parse
import time

st.set_page_config(page_title="LLEGANDO - Red Integral de Movilidad y Logística", page_icon="🚗", layout="wide")

# ==========================================
# ESTILOS CSS AZULES CORPORATIVOS (ACORDE AL LOGO)
# ==========================================
st.markdown("""
    <style>
        .main {
            background-color: #F8FAFC;
        }
        .metric-card {
            background-color: #FFFFFF;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(10, 37, 64, 0.05);
            border-left: 5px solid #1E3A8A;
            margin-bottom: 15px;
        }
        h1, h2, h3 {
            color: #0A2540;
            font-family: 'Helvetica Neue', sans-serif;
        }
        .stButton>button {
            background-color: #1E3A8A;
            color: white;
            border-radius: 6px;
            font-weight: 600;
            border: none;
        }
        .stButton>button:hover {
            background-color: #0F172A;
            color: white;
        }
        .cintillo-alerta {
            background-color: #0A2540;
            color: white;
            padding: 18px 24px;
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(10, 37, 64, 0.3);
            margin-bottom: 20px;
            border-left: 8px solid #3B82F6;
            font-family: 'Helvetica Neue', sans-serif;
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
        if "Pago_Verificado" not in df_temp.columns:
            df_temp["Pago_Verificado"] = "No"
        if "Liquidado" not in df_temp.columns:
            df_temp["Liquidado"] = "No"
        if "Valoracion" not in df_temp.columns:
            df_temp["Valoracion"] = "Sin valorar"
        return df_temp
    else:
        df_inicial = pd.DataFrame(columns=[
            "ID_Servicio", "Fecha_Hora", "Tipo_Servicio", "Cliente", 
            "Telefono_Cliente", "Origen", "Destino", "Conductor", 
            "Monto_USD", "Metodo_Pago", "Ref_Pago_Movil", "Pago_Verificado", "Estado", "Liquidado", "Valoracion", "Notas"
        ])
        df_inicial.to_excel(DB_FILE, index=False)
        return df_inicial

def guardar_datos(df):
    df["Ref_Pago_Movil"] = df["Ref_Pago_Movil"].astype(str)
    df["Telefono_Cliente"] = df["Telefono_Cliente"].astype(str)
    df["Liquidado"] = df["Liquidado"].astype(str)
    df["Valoracion"] = df["Valoracion"].astype(str)
    df.to_excel(DB_FILE, index=False)

def cargar_prestadores():
    if os.path.exists(DB_PRESTADORES):
        df_p_temp = pd.read_excel(DB_PRESTADORES, dtype=str)
        cols_necesarias = ["Nombre", "Telefono", "Vehiculo", "Marca", "Modelo", "Placa", "Foto_Path", "Pago_Afiliacion_5USD", "Pago_Mensualidad_5USD", "Estatus"]
        for col in cols_necesarias:
            if col not in df_p_temp.columns:
                df_p_temp[col] = "Pendiente" if "Pago" in col else ""
        return df_p_temp
    else:
        df_p = pd.DataFrame(columns=["Nombre", "Telefono", "Vehiculo", "Marca", "Modelo", "Placa", "Foto_Path", "Pago_Afiliacion_5USD", "Pago_Mensualidad_5USD", "Estatus"])
        df_p.to_excel(DB_PRESTADORES, index=False)
        return df_p

def guardar_prestadores(df_p):
    df_p.to_excel(DB_PRESTADORES, index=False)

def cargar_liquidaciones():
    if os.path.exists(DB_LIQUIDACIONES):
        return pd.read_excel(DB_LIQUIDACIONES, dtype=str)
    else:
        df_l = pd.DataFrame(columns=["ID_Liquidacion", "Fecha", "Prestador", "Monto_USD", "Referencia_Pago", "Estatus_Confirmacion"])
        df_l.to_excel(DB_LIQUIDACIONES, index=False)
        return df_l

def guardar_liquidaciones(df_l):
    df_l.to_excel(DB_LIQUIDACIONES, index=False)

df = cargar_datos()
df_prestadores = cargar_prestadores()
df_liquidaciones = cargar_liquidaciones()

# ==========================================
# BARRA LATERAL - LOGOTIPO OFICIAL Y NAVEGACIÓN
# ==========================================
if os.path.exists("logo.jpg"):
    st.sidebar.image("logo.jpg", use_container_width=True)
elif os.path.exists("logo.png"):
    st.sidebar.image("logo.png", use_container_width=True)
else:
    st.sidebar.markdown("### 🔵 LLEGANDO")
    st.sidebar.markdown("<p style='font-size: 0.8em; color: #666;'>Sube 'logo.jpg' a tu repositorio para mostrar el logotipo oficial.</p>", unsafe_allow_html=True)

st.sidebar.markdown("---")

rol = st.sidebar.radio("Selecciona tu Panel:", [
    "👤 Portal de Cliente", 
    "🚗 Portal de Prestador", 
    "⚙️ Panel Gerencial (Admin)"
])

st.sidebar.markdown("---")
st.sidebar.metric("💵 Tasa BCV Oficial", f"Bs. {st.session_state.tasa_bcv:,.2f}")

# ==========================================
# 1. VISTA CLIENTE
# ==========================================
if "Cliente" in rol:
    st.markdown("<h2 style='color: #0A2540;'>🔵 LLEGANDO - Solicitud y Estatus</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    tab_solicitar, tab_consultar = st.tabs(["📝 Solicitar Nuevo Servicio", "🔍 Consultar Mis Viajes por Teléfono"])
    
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
            
            if st.form_submit_button("🚀 Enviar Solicitud Ahora"):
                if nombre and telefono and origen and destino:
                    nuevo_id = 1001 if len(df) == 0 else int(pd.to_numeric(df["ID_Servicio"], errors='coerce').max()) + 1
                    nuevo_registro = {
                        "ID_Servicio": nuevo_id,
                        "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Tipo_Servicio": tipo_serv,
                        "Cliente": nombre,
                        "Telefono_Cliente": str(telefono).strip(),
                        "Origen": origen,
                        "Destino": destino,
                        "Conductor": "Por Asignar",
                        "Monto_USD": 0.0,
                        "Metodo_Pago": metodo_pago,
                        "Ref_Pago_Movil": "",
                        "Pago_Verificado": "No",
                        "Estado": "Solicitado",
                        "Liquidado": "No",
                        "Valoracion": "Sin valorar",
                        "Notas": notas
                    }
                    df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
                    guardar_datos(df)
                    st.success(f"¡Solicitud enviada con éxito! Tu número de ID es **#{nuevo_id}**.")
                else:
                    st.error("Por favor completa los campos obligatorios (*).")
                    
    with tab_consultar:
        st.subheader("🔍 Mis Servicios Activos e Historial")
        
        telefono_busqueda = st.text_input("Ingresa tu número de teléfono registrado:")
        
        if telefono_busqueda:
            t_busq_clean = telefono_busqueda.strip()
            resultados = df[df["Telefono_Cliente"].astype(str).str.strip() == t_busq_clean]
            
            if len(resultados) > 0:
                for idx_res, serv in resultados.iterrows():
                    monto_usd = float(serv['Monto_USD'])
                    monto_ves = monto_usd * st.session_state.tasa_bcv
                    estatus_s = str(serv['Estado']).strip()
                    pago_v = str(serv.get('Pago_Verificado', 'No')).strip()
                    
                    if pago_v.lower() in ["sí", "si"] or estatus_s in ["Asignado", "En Camino", "Llegué al Sitio", "En Trayecto", "Completado"]:
                        st.markdown(f'<div class="cintillo-alerta"><h3 style="color: #ffffff; margin: 0 0 5px 0;">🔵 LLEGANDO • ¡Actualización en tu Viaje #{serv["ID_Servicio"]}!</h3><p style="margin: 0; font-size: 1.05em;">Estatus actual: <b>{estatus_s}</b> | Prestador: <b>{serv["Conductor"]}</b> | Pago: <b>{pago_v}</b></p></div>', unsafe_allow_html=True)

                    with st.expander(f"📍 Servicio #{serv['ID_Servicio']} - {serv['Tipo_Servicio']} [{estatus_s}]", expanded=True):
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"**👤 Cliente:** {serv['Cliente']}")
                            st.markdown(f"**📍 Origen:** {serv['Origen']}")
                            st.markdown(f"**🏁 Destino:** {serv['Destino']}")
                        with c2:
                            st.markdown(f"**🚦 Estado:** `{estatus_s}`")
                            st.markdown(f"**💵 Tarifa:** **${monto_usd:,.2f} USD** (Bs. {monto_ves:,.2f})")
                        
                        if estatus_s == "Cancelado":
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
                                
                                ref_ingresada = st.text_input("📝 Ingresa los dígitos de tu referencia bancaria:", value=ref_actual, key=f"ref_{serv['ID_Servicio']}")
                                if st.button(f"Reportar Referencia #{serv['ID_Servicio']}", key=f"btn_ref_{serv['ID_Servicio']}"):
                                    if ref_ingresada:
                                        idx = df[df["ID_Servicio"] == serv['ID_Servicio']].index
                                        df.at[idx[0], "Ref_Pago_Movil"] = str(ref_ingresada)
                                        guardar_datos(df)
                                        st.success("¡Referencia reportada! Notificación enviada al administrador.")
                                        st.rerun()
                                    else:
                                        st.error("Ingresa la referencia de pago.")
                            
                            if pago_v.lower() in ["sí", "si"] or str(serv['Metodo_Pago']).strip().lower() != "pago móvil":
                                st.markdown("---")
                                st.success("✅ **¡Pago Confirmado por la Empresa!**")
                                st.markdown(f"### 👨‍✈️ Tu Prestador Asignado: **{serv['Conductor']}**")
                                
                                info_cond = df_prestadores[df_prestadores["Nombre"].str.lower() == str(serv['Conductor']).lower()]
                                if len(info_cond) > 0:
                                    cond_row = info_cond.iloc[0]
                                    foto_path = str(cond_row.get('Foto_Path', '')).strip()
                                    marca_v = str(cond_row.get('Marca', 'N/A')).strip()
                                    modelo_v = str(cond_row.get('Modelo', 'N/A')).strip()
                                    placa_v = str(cond_row.get('Placa', 'N/A')).strip()
                                    tlf_cond = str(cond_row['Telefono']).strip().replace(" ", "").replace("+", "")
                                    
                                    col_info1, col_info2 = st.columns([1, 2])
                                    with col_info1:
                                        if foto_path and foto_path != "nan" and os.path.exists(foto_path):
                                            st.image(foto_path, width=120, caption="Foto de Perfil")
                                        else:
                                            st.info("📷 [Sin foto registrada]")
                                    with col_info2:
                                        st.markdown(f"🚗 **Vehículo:** {cond_row['Vehiculo']}")
                                        st.markdown(f"🏷️ **Marca / Modelo:** {marca_v} {modelo_v}")
                                        st.markdown(f"🔢 **Placa:** `{placa_v}`")
                                    
                                    msg_c = f"Hola {serv['Conductor']}, le escribo desde LLEGANDO respecto al servicio #{serv['ID_Servicio']}."
                                    url_wa_cond = f"https://wa.me/{tlf_cond}?text={urllib.parse.quote(msg_c)}"
                                    st.link_button("💬 Contactar Conductor por WhatsApp", url_wa_cond)
                                
                                if estatus_s == "Completado":
                                    st.markdown("---")
                                    st.markdown("#### ⭐ Valora la Calidad de tu Servicio")
                                    val_actual = str(serv.get("Valoracion", "Sin valorar"))
                                    
                                    escalas = ["Sin valorar", "Mal servicio", "Regular servicio", "Buen servicio", "Excelente servicio"]
                                    idx_val = escalas.index(val_actual) if val_actual in escalas else 0
                                    
                                    nueva_valoracion = st.selectbox("Selecciona tu valoración:", escalas, index=idx_val, key=f"val_{serv['ID_Servicio']}")
                                    if st.button(f"Guardar Valoración #{serv['ID_Servicio']}", key=f"btn_val_{serv['ID_Servicio']}"):
                                        idx = df[df["ID_Servicio"] == serv['ID_Servicio']].index
                                        df.at[idx[0], "Valoracion"] = nueva_valoracion
                                        guardar_datos(df)
                                        st.success("¡Gracias por valorar nuestro servicio!")
                                        st.rerun()
                            else:
                                st.info("⏳ **Pago en proceso de verificación por el administrador.**")
            else:
                st.warning("No se encontró ningún servicio registrado con ese número de teléfono.")
        else:
            st.info("💡 Ingresa tu número de teléfono arriba para consultar tus viajes activos.")
            
        time.sleep(3)
        st.rerun()

# ==========================================
# 2. VISTA PRESTADOR DE SERVICIOS
# ==========================================
elif "Prestador" in rol:
    st.markdown("<h2 style='color: #0A2540;'>🚗 Panel Profesional del Prestador - LLEGANDO</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    nom_prestador = st.text_input("Ingresa tu Nombre exacto autorizado para acceder:")
    
    if not nom_prestador:
        st.info("💡 Por favor ingresa tu nombre exacto autorizado para ver tus viajes asignados y estatus de suscripción.")
    else:
        prestador_valido = df_prestadores[(df_prestadores["Nombre"].str.lower() == nom_prestador.strip().lower()) & (df_prestadores["Estatus"] == "Autorizado")]
        
        if len(prestador_valido) == 0:
            st.warning("⚠️ Tu cuenta se encuentra suspendida, no existe o no ha sido autorizada por la administración.")
        else:
            st.success("✅ ¡Acceso Autorizado!")
            
            p_row = prestador_valido.iloc[0]
            afiliacion_est = str(p_row.get("Pago_Afiliacion_5USD", "Pendiente"))
            mensualidad_est = str(p_row.get("Pago_Mensualidad_5USD", "Pendiente"))
            
            st.markdown("### 📋 Estatus de Suscripción y Membresía ($5 Afiliación / $5 Mensualidad)")
            col_s1, col_s2 = st.columns(2)
            col_s1.metric("Afiliación ($5)", afiliacion_est)
            col_s2.metric("Mensualidad ($5)", mensualidad_est)
            
            st.markdown("---")

            servicios_cond = df[df["Conductor"].astype(str).str.strip().str.lower() == nom_prestador.strip().lower()]
            viajes_activos_cond = servicios_cond[servicios_cond["Estado"].isin(["Asignado", "En Camino", "Llegué al Sitio", "En Trayecto"])]
            
            if len(viajes_activos_cond) > 0:
                st.markdown(f'<div class="cintillo-alerta"><h3 style="color: #ffffff; margin: 0 0 5px 0;">🚗 LLEGANDO • ¡Nuevo Servicio Asignado!</h3><p style="margin: 0; font-size: 1.05em;">Tienes <b>{len(viajes_activos_cond)}</b> viaje(s) activo(s) esperando por atención.</p></div>', unsafe_allow_html=True)

            tab_viajes, tab_finanzas, tab_liquidaciones = st.tabs(["📋 Mis Viajes Activos", "💰 Mis Ganancias", "🤝 Liquidaciones y Pagos"])
            
            with tab_viajes:
                if len(servicios_cond) > 0:
                    for idx, serv in servicios_cond.iterrows():
                        monto_usd = float(serv['Monto_USD'])
                        monto_ves = monto_usd * st.session_state.tasa_bcv
                        ganancia_usd = monto_usd * (st.session_state.porcentaje_prestador / 100)
                        ganancia_ves = ganancia_usd * st.session_state.tasa_bcv
                        
                        with st.expander(f"Servicio #{serv['ID_Servicio']} - {serv['Origen']} ➡️ {serv['Destino']} [{serv['Estado']}]"):
                            st.write(f"**Cliente:** {serv['Cliente']} | **Teléfono:** {serv['Telefono_Cliente']}")
                            st.write(f"**Monto Cobrar:** ${monto_usd:.2f} USD (Bs. {monto_ves:,.2f})")
                            st.write(f"**Tu Ganancia ({st.session_state.porcentaje_prestador}%):** **${ganancia_usd:.2f} USD** (Bs. {ganancia_ves:,.

import streamlit as st
import pandas as pd
import os
import requests
from datetime import datetime
import urllib.parse
import streamlit.components.v1 as components

st.set_page_config(page_title="LLEGANDO - Movilidad & Delivery", page_icon="🚀", layout="wide")

# ==========================================
# ESTILOS CSS 100% VERDES Y PROFESIONALES
# ==========================================
st.markdown("""
    <style>
        .main {
            background-color: #F4FBF7;
        }
        .metric-card {
            background-color: #FFFFFF;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 128, 0, 0.05);
            border-left: 5px solid #198754;
            margin-bottom: 15px;
        }
        h1, h2, h3 {
            color: #0F5132;
            font-family: 'Helvetica Neue', sans-serif;
        }
        .stButton>button {
            background-color: #198754;
            color: white;
            border-radius: 6px;
            font-weight: 600;
            border: none;
        }
        .stButton>button:hover {
            background-color: #146c43;
            color: white;
        }
    </style>
""", unsafe_allow_html=True)

DB_FILE = "base_datos_servicios.xlsx"
DB_PRESTADORES = "prestadores_autorizados.xlsx"
DB_LIQUIDACIONES = "liquidaciones_prestadores.xlsx"

CLAVE_ADMIN_SECRETA = "Sky1012#"

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

if "ultimo_pago_verificado" not in st.session_state:
    st.session_state.ultimo_pago_verificado = {}

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
        return df_temp
    else:
        df_inicial = pd.DataFrame(columns=[
            "ID_Servicio", "Fecha_Hora", "Tipo_Servicio", "Cliente", 
            "Telefono_Cliente", "Origen", "Destino", "Conductor", 
            "Monto_USD", "Metodo_Pago", "Ref_Pago_Movil", "Pago_Verificado", "Estado", "Notas"
        ])
        df_inicial.to_excel(DB_FILE, index=False)
        return df_inicial

def guardar_datos(df):
    df["Ref_Pago_Movil"] = df["Ref_Pago_Movil"].astype(str)
    df["Telefono_Cliente"] = df["Telefono_Cliente"].astype(str)
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
# BARRA LATERAL - LLEGANDO
# ==========================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/7541/7541900.png", width=80)
st.sidebar.markdown("### 🚀 LLEGANDO")
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
    st.markdown("<h2 style='color: #198754;'>🚀 LLEGANDO - Solicitud y Estatus</h2>", unsafe_allow_html=True)
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
                    id_s = serv['ID_Servicio']
                    pago_verif_actual = str(serv.get('Pago_Verificado', 'No')).strip().lower()
                    estado_actual = serv['Estado']
                    
                    pago_previo = st.session_state.ultimo_pago_verificado.get(id_s, "No")
                    if pago_previo != "sí" and pago_verif_actual in ["sí", "si"] and estado_actual not in ["Cancelado", "Completado"]:
                        audio_confirmacion_cliente = "<script>try { const ac = new (window.AudioContext || window.webkitAudioContext)(); const osc = ac.createOscillator(); const g = ac.createGain(); osc.type='triangle'; osc.frequency.value=600; g.gain.setValueAtTime(0.15, ac.currentTime); osc.connect(g); g.connect(ac.destination); osc.start(); osc.stop(ac.currentTime + 0.4); } catch(e){}</script>"
                        components.html(audio_confirmacion_cliente, height=0)
                        st.toast("🎉 ¡Tu pago ha sido verificado y confirmado por la empresa!", icon="✅")
                    
                    st.session_state.ultimo_pago_verificado[id_s] = pago_verif_actual

                    monto_usd = float(serv['Monto_USD'])
                    monto_ves = monto_usd * st.session_state.tasa_bcv
                    
                    with st.expander(f"📍 Servicio #{serv['ID_Servicio']} - {serv['Tipo_Servicio']} [{serv['Estado']}]", expanded=True):
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"**👤 Cliente:** {serv['Cliente']}")
                            st.markdown(f"**📍 Origen:** {serv['Origen']}")
                            st.markdown(f"**🏁 Destino:** {serv['Destino']}")
                        with c2:
                            st.markdown(f"**🚦 Estado:** `{serv['Estado']}`")
                            st.markdown(f"**💵 Tarifa:** **${monto_usd:,.2f} USD** (Bs. {monto_ves:,.2f})")
                        
                        if serv['Estado'] == "Cancelado":
                            st.error("⚠️ Este servicio ha sido cancelado.")
                        else:
                            ref_actual = "" if pd.isna(serv['Ref_Pago_Movil']) or str(serv['Ref_Pago_Movil']) == "nan" else str(serv['Ref_Pago_Movil']).strip()
                            pago_verif = str(serv.get('Pago_Verificado', 'No')).strip()
                            
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
                                        st.success("¡Referencia reportada! Esperando confirmación bancaria del administrador.")
                                        st.rerun()
                                    else:
                                        st.error("Ingresa la referencia de pago.")
                            
                            if pago_verif.lower() == "sí" or pago_verif.lower() == "si" or str(serv['Metodo_Pago']).strip().lower() != "pago móvil":
                                st.markdown("---")
                                st.success("✅ **¡Pago Confirmado por la Empresa!**")
                                st.markdown(f"### 👨‍✈️ Tu Prestador Asignado es: **{serv['Conductor']}**")
                                
                                info_cond = df_prestadores[df_prestadores["Nombre"].str.lower() == str(serv['Conductor']).lower()]
                                if len(info_cond) > 0:
                                    tlf_cond = info_cond.iloc[0]["Telefono"].strip().replace(" ", "").replace("+", "")
                                    msg_c = f"Hola {serv['Conductor']}, le escribo desde LLEGANDO respecto a mi servicio #{serv['ID_Servicio']}."
                                    url_wa_cond = f"https://wa.me/{tlf_cond}?text={urllib.parse.quote(msg_c)}"
                                    st.link_button("💬 Contactar Conductor por WhatsApp", url_wa_cond)
                            else:
                                st.info("⏳ **Pago en proceso de verificación por el administrador.**")
            else:
                st.warning("No se encontró ningún servicio registrado con ese número de teléfono.")

# ==========================================
# 2. VISTA PRESTADOR DE SERVICIOS
# ==========================================
elif "Prestador" in rol:
    st.markdown("<h2 style='color: #198754;'>🚗 Panel Profesional del Prestador - LLEGANDO</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    nom_prestador = st.text_input("Ingresa tu Nombre exacto autorizado para acceder:")
    
    if nom_prestador:
        prestador_valido = df_prestadores[(df_prestadores["Nombre"].str.lower() == nom_prestador.lower()) & (df_prestadores["Estatus"] == "Autorizado")]
        
        if len(prestador_valido) == 0:
            st.warning("⚠️ Tu nombre no se encuentra registrado o aún no ha sido autorizado.")
        else:
            st.success("✅ ¡Acceso Autorizado!")
            
            # Script de Alarma en Tiempo Real (Polling nativo en JS cada 5s sin recargar la página)
            realtime_prestador_js = f"""
                <script>
                    let ultimoConteo = {len(df[(df["Conductor"].astype(str).str.strip().str.lower() == nom_prestador.strip().lower()) & (df["Estado"].isin(["Asignado", "En Camino", "Llegué al Sitio", "En Trayecto"]))])};
                    
                    setInterval(async () => {{
                        try {{
                            const response = window.location.href;
                            // Polling silencioso de verificación de nuevos registros en la app
                            const res = await fetch(window.location.href);
                            const html = await res.text();
                            // Si la app detecta cambios en las asignaciones activas, suena la alarma y recarga
                            if (html.includes("🚨")) {{
                                const ac = new (window.AudioContext || window.webkitAudioContext)();
                                const osc = ac.createOscillator();
                                const g = ac.createGain();
                                osc.type='square'; osc.frequency.value=750;
                                g.gain.setValueAtTime(0.12, ac.currentTime);
                                osc.connect(g); g.connect(ac.destination);
                                osc.start(); osc.stop(ac.currentTime + 0.6);
                            }}
                        }} catch (e) {{}}
                    }}, 6000);
                </script>
            """
            components.html(realtime_prestador_js, height=0)

            servicios_cond = df[df["Conductor"].astype(str).str.strip().str.lower() == nom_prestador.strip().lower()]
            viajes_activos_cond = servicios_cond[servicios_cond["Estado"].isin(["Asignado", "En Camino", "Llegué al Sitio", "En Trayecto"])]
            
            if len(viajes_activos_cond) > 0:
                st.warning(f"🚨 ¡Tienes **{len(viajes_activos_cond)}** servicio(s) activo(s) asignado(s)!")

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
                            st.write(f"**Tu Ganancia ({st.session_state.porcentaje_prestador}%):** **${ganancia_usd:.2f} USD** (Bs. {ganancia_ves:,.2f})")
                            st.write(f"**Ref. Pago Móvil:** {serv['Ref_Pago_Movil']} | **Pago Verificado:** {serv.get('Pago_Verificado', 'No')}")
                            
                            estados_posibles = ["Asignado", "En Camino", "Llegué al Sitio", "En Trayecto", "Completado", "Cancelado"]
                            idx_est = estados_posibles.index(serv['Estado']) if serv['Estado'] in estados_posibles else 0
                            nuevo_estado = st.selectbox("Actualizar Estatus:", estados_posibles, index=idx_est, key=f"st_{serv['ID_Servicio']}")
                            
                            col_b1, col_b2, col_b3 = st.columns(3)
                            with col_b1:
                                if st.button(f"Guardar Estatus #{serv['ID_Servicio']}", key=f"btn_{serv['ID_Servicio']}"):
                                    df.at[idx, "Estado"] = nuevo_estado
                                    guardar_datos(df)
                                    st.success("¡Estatus actualizado!")
                                    st.rerun()
                            with col_b2:
                                tlf_cliente = str(serv['Telefono_Cliente']).strip().replace(" ", "").replace("+", "")
                                msg = f"Hola {serv['Cliente']}, le escribe su conductor de LLEGANDO {nom_prestador} respecto al servicio #{serv['ID_Servicio']}."
                                url_wa = f"https://wa.me/{tlf_cliente}?text={urllib.parse.quote(msg)}"
                                st.link_button("💬 WhatsApp", url_wa)
                            with col_b3:
                                if st.button(f"❌ Cancelar #{serv['ID_Servicio']}", key=f"btn_canc_p_{serv['ID_Servicio']}"):
                                    df.at[idx, "Estado"] = "Cancelado"
                                    guardar_datos(df)
                                    st.warning("Servicio cancelado.")
                                    st.rerun()
                else:
                    st.info("No tienes servicios asignados actualmente.")
                    
            with tab_finanzas:
                completados = servicios_cond[servicios_cond["Estado"] == "Completado"]
                total_cobrado_usd = pd.to_numeric(completados["Monto_USD"], errors='coerce').sum()
                ganancia_net_usd = total_cobrado_usd * (st.session_state.porcentaje_prestador / 100)
                ganancia_net_ves = ganancia_net_usd * st.session_state.tasa_bcv
                
                c_g1, c_g2, c_g3 = st.columns(3)
                c_g1.metric("Servicios Completados", len(completados))
                c_g2.metric("Ganancia Total USD", f"${ganancia_net_usd:,.2f} USD")
                c_g3.metric("Ganancia Total VES", f"Bs. {ganancia_net_ves:,.2f}")
                
                if len(completados) > 0:
                    st.dataframe(completados[["ID_Servicio", "Fecha_Hora", "Tipo_Servicio", "Monto_USD", "Metodo_Pago"]], use_container_width=True)

            with tab_liquidaciones:
                st.subheader("🤝 Liquidaciones y Pagos Realizados por LLEGANDO")
                liq_prestador = df_liquidaciones[df_liquidaciones["Prestador"].astype(str).str.strip().str.lower() == nom_prestador.strip().lower()]
                
                if len(liq_prestador) > 0:
                    for idx_l, liq in liq_prestador.iterrows():
                        monto_liq_usd = float(liq['Monto_USD'])
                        monto_liq_ves = monto_liq_usd * st.session_state.tasa_bcv
                        estatus_conf = str(liq

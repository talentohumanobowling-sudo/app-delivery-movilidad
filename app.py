import streamlit as st
import pandas as pd
import os
import requests
from datetime import datetime
import urllib.parse
import streamlit.components.v1 as components

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

# Variables de control para evitar que el sonido suene al abrir la app (solo en eventos nuevos)
if "ultimo_conteo_admin" not in st.session_state:
    st.session_state.ultimo_conteo_admin = None
if "ultimo_conteo_prestador" not in st.session_state:
    st.session_state.ultimo_conteo_prestador = None
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
            
            if st.form_submit_button("🚀 Enviar Solicitud"):
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
        telefono_busqueda = st.text_input("Ingresa tu número de teléfono registrado para ver tus servicios:")
        
        if telefono_busqueda:
            t_busq_clean = telefono_busqueda.strip()
            resultados = df[df["Telefono_Cliente"].astype(str).str.strip() == t_busq_clean]
            
            if len(resultados) > 0:
                for idx_res, serv in resultados.iterrows():
                    id_s = serv['ID_Servicio']
                    pago_verif_actual = str(serv.get('Pago_Verificado', 'No')).strip().lower()
                    estado_actual = serv['Estado']
                    
                    # Detectar si el pago acaba de ser aprobado (solo suena en el cambio de "No" a "Sí")
                    pago_previo = st.session_state.ultimo_pago_verificado.get(id_s, "No")
                    if pago_previo != "sí" and pago_verif_actual in ["sí", "si"] and estado_actual not in ["Cancelado", "Completado"]:
                        audio_confirmacion_cliente = """
                            <script>
                                const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
                                function playConfirmSound() {
                                    const osc = audioCtx.createOscillator();
                                    const gain = audioCtx.createGain();
                                    osc.type = 'triangle';
                                    osc.frequency.value = 600; 
                                    gain.gain.setValueAtTime(0.15, audioCtx.currentTime);
                                    osc.connect(gain);
                                    gain.connect(audioCtx.destination);
                                    osc.start();
                                    osc.stop(audioCtx.currentTime + 0.4);
                                }
                                playConfirmSound();
                            </script>
                        """
                        components.html(audio_confirmacion_cliente, height=0)
                        st.toast("🎉 ¡Tu pago ha sido verificado y confirmado por la empresa!", icon="✅")
                    
                    # Actualizar estado previo en memoria
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
                                st.markdown(f"### 👨‍✈️ Tu Prestador de Servicio Asignado es: **{serv['Conductor']}**")
                                
                                info_cond = df_prestadores[df_prestadores["Nombre"].str.lower() == str(serv['Conductor']).lower()]
                                if len(info_cond) > 0:
                                    tlf_cond = info_cond.iloc[0]["Telefono"].strip().replace(" ", "").replace("+", "")
                                    msg_c = f"Hola {serv['Conductor']}, le escribo desde la app respecto al servicio #{serv['ID_Servicio']}."
                                    url_wa_cond = f"https://wa.me/{tlf_cond}?text={urllib.parse.quote(msg_c)}"
                                    st.link_button("💬 Contactar a mi Conductor por WhatsApp", url_wa_cond)
                            else:
                                st.info("⏳ **Pago en proceso de verificación por el administrador.** Una vez confirmado, verás aquí los datos de tu conductor.")
            else:
                st.warning("No se encontró ningún servicio registrado con ese número de teléfono.")

# ==========================================
# 2. VISTA PRESTADOR DE SERVICIOS
# ==========================================
elif rol == "🚗 Prestador de Servicios":
    st.title("🚗 Panel del Prestador de Servicios (Tiempo Real)")
    
    auto_refresh_prestador = """
        <script>
            setTimeout(function(){
                window.location.reload();
            }, 5000);
        </script>
    """
    components.html(auto_refresh_prestador, height=0)
    
    nom_prestador = st.text_input("Ingresa tu Nombre exacto (autorizado por admin) para acceder:")
    
    if nom_prestador:
        prestador_valido = df_prestadores[(df_prestadores["Nombre"].str.lower() == nom_prestador.lower()) & (df_prestadores["Estatus"] == "Autorizado")]
        
        if len(prestador_valido) == 0:
            st.warning("⚠️ Tu nombre no se encuentra registrado o aún no ha sido autorizado por el Administrador.")
        else:
            st.success("✅ ¡Acceso Autorizado!")
            servicios_cond = df[df["Conductor"].astype(str).str.strip().str.lower() == nom_prestador.strip().lower()]
            viajes_activos_cond = servicios_cond[servicios_cond["Estado"].isin(["Asignado", "En Camino", "Llegué al Sitio", "En Trayecto"])]
            
            conteo_actual_prestador = len(viajes_activos_cond)
            
            # Sonido solo si el número de viajes activos aumenta en tiempo real
            if st.session_state.ultimo_conteo_prestador is not None and conteo_actual_prestador > st.session_state.ultimo_conteo_prestador:
                audio_prestador = """
                    <script>
                        const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
                        function playDriverAlert() {
                            const osc = audioCtx.createOscillator();
                            const gain = audioCtx.createGain();
                            osc.type = 'square';
                            osc.frequency.value = 750; 
                            gain.gain.setValueAtTime(0.12, audioCtx.currentTime);
                            osc.connect(gain);
                            gain.connect(audioCtx.destination);
                            osc.start();
                            osc.stop(audioCtx.currentTime + 0.6);
                        }
                        playDriverAlert();
                    </script>
                """
                components.html(audio_prestador, height=0)
            
            st.session_state.ultimo_conteo_prestador = conteo_actual_prestador

            if len(viajes_activos_cond) > 0:
                st.warning(f"🚨 ¡Tienes **{conteo_actual_prestador}** servicio(s) activo(s) asignado(s) pendiente(s)!")

            tab_viajes, tab_finanzas = st.tabs(["📋 Mis Viajes y Estatus", "💰 Mis Ganancias y Porcentaje"])
            
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
                            nuevo_estado = st.selectbox("Actualizar Estatus:", estados_posibles

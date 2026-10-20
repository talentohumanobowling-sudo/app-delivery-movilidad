import streamlit as st
import pandas as pd
import os
import requests
from datetime import datetime
import urllib.parse
import streamlit.components.v1 as components

st.set_page_config(page_title="LLEGANDO - Movilidad & Delivery", page_icon="🚀", layout="wide")

# ==========================================
# CONFIGURACIÓN PWA Y CINTILLO FLOTANTE
# ==========================================
pwa_header = """
    <link rel="manifest" href="data:application/manifest+json;charset=utf-8,{
        'name': 'LLEGANDO - Movilidad & Delivery',
        'short_name': 'LLEGANDO',
        'start_url': '.',
        'display': 'standalone',
        'background_color': '#F4FBF7',
        'theme_color': '#198754',
        'icons': [{
            'src': 'https://cdn-icons-png.flaticon.com/512/7541/7541900.png',
            'sizes': '512x512',
            'type': 'image/png'
        }]
    }">
"""
st.markdown(pwa_header, unsafe_allow_html=True)

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
        
        # Registro automático y solicitud de permisos para notificaciones flotantes en el teléfono
        components.html("""
            <div style="background:#E8F5E9; padding:12px; border-radius:8px; border:1px solid #198754; margin-bottom:15px; font-family:sans-serif;">
                <p style="margin:0 0 8px 0; color:#0F5132; font-weight:600;">📱 Tu Identidad y Cintillo de Alertas:</p>
                <input type="text" id="tel_reg" placeholder="Ingresa tu número (ej: 04141234567)" style="width:65%; padding:6px; border-radius:4px; border:1px solid #ccc;">
                <button onclick="guardarTel()" style="background:#198754; color:white; border:none; padding:7px 10px; border-radius:4px; font-weight:bold; cursor:pointer;">Guardar</button>
                <button onclick="pedirPermiso()" style="background:#0F5132; color:white; border:none; padding:7px 10px; border-radius:4px; font-weight:bold; cursor:pointer; margin-left:5px;">🔔 Activar Cintillo</button>
                <p id="msg_reg" style="margin:5px 0 0 0; color:#146c43; font-size:0.9em; font-weight:bold;"></p>
            </div>
            <script>
                const input = document.getElementById('tel_reg');
                const msg = document.getElementById('msg_reg');
                const guardado = localStorage.getItem('llegando_telefono_cliente');
                if (guardado) {
                    input.value = guardado;
                    msg.innerText = "✅ Teléfono vinculado: " + guardado;
                }
                function guardarTel() {
                    if(input.value.trim() !== "") {
                        localStorage.setItem('llegando_telefono_cliente', input.value.trim());
                        msg.innerText = "✅ ¡Teléfono guardado en tu dispositivo!";
                        setTimeout(() => { window.location.reload(); }, 800);
                    }
                }
                function pedirPermiso() {
                    if ("Notification" in window) {
                        Notification.requestPermission().then(permission => {
                            if (permission === "granted") {
                                alert("¡Cintillo de notificaciones activado con éxito!");
                            } else {
                                alert("Permiso de notificaciones denegado.");
                            }
                        });
                    }
                }
            </script>
        """, height=130)

        telefono_busqueda = st.text_input("Confirma tu número de teléfono registrado:")
        
        if telefono_busqueda:
            if st.button("🔊 Activar Alerta Sonora en Tiempo Real"):
                audio_init_cli = "<script>try { const ac = new (window.AudioContext || window.webkitAudioContext)(); const osc = ac.createOscillator(); const g = ac.createGain(); osc.type='triangle'; osc.frequency.value=600; g.gain.setValueAtTime(0.15, ac.currentTime); osc.connect(g); osc.connect(ac.destination); osc.start(); osc.stop(ac.currentTime + 0.3); } catch(e){}</script>"
                components.html(audio_init_cli, height=0)
                st.toast("🎉 ¡Cintillo de notificaciones activo!", icon="🚀")

            # Script de sondeo que lanza un aviso flotante y cintillo si hay confirmación
            realtime_cliente_js = """
                <script>
                    setInterval(async () => {
                        try {
                            const res = await fetch(window.location.href);
                            const html = await res.text();
                            if (html.includes("Pago Confirmado") || html.includes("✅")) {
                                if (Notification.permission === "granted") {
                                    new Notification("🚀 LLEGANDO", {
                                        body: "¡Tu pago ha sido verificado y confirmado por la empresa!",
                                        icon: "https://cdn-icons-png.flaticon.com/512/7541/7541900.png"
                                    });
                                }
                                const ac = new (window.AudioContext || window.webkitAudioContext)();
                                const osc = ac.createOscillator();
                                const g = ac.createGain();
                                osc.type='triangle'; osc.frequency.value=600;
                                g.gain.setValueAtTime(0.15, ac.currentTime);
                                osc.connect(g); g.connect(ac.destination);
                                osc.start(); osc.stop(ac.currentTime + 0.5);
                            }
                        } catch (e) {}
                    }, 4000);
                </script>
            """
            components.html(realtime_cliente_js, height=0)

            t_busq_clean = telefono_busqueda.strip()
            resultados = df[df["Telefono_Cliente"].astype(str).str.strip() == t_busq_clean]
            
            if len(resultados) > 0:
                for idx_res, serv in resultados.iterrows():
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
                                    msg_c = f"Hola {serv['Conductor']}, le escribo desde LLEGANDO respecto al servicio #{serv['ID_Servicio']}."
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
    
    components.html("""
        <div style="background:#E8F5E9; padding:12px; border-radius:8px; border:1px solid #198754; margin-bottom:15px; font-family:sans-serif;">
            <p style="margin:0 0 8px 0; color:#0F5132; font-weight:600;">🚗 Tu Nombre Autorizado y Cintillo de Alertas:</p>
            <input type="text" id="prest_reg" placeholder="Tu Nombre Exacto" style="width:65%; padding:6px; border-radius:4px; border:1px solid #ccc;">
            <button onclick="guardarPrest()" style="background:#198754; color:white; border:none; padding:7px 10px; border-radius:4px; font-weight:bold; cursor:pointer;">Guardar</button>
            <button onclick="pedirPermiso()" style="background:#0F5132; color:white; border:none; padding:7px 10px; border-radius:4px; font-weight:bold; cursor:pointer; margin-left:5px;">🔔 Activar Cintillo</button>
            <p id="msg_prest" style="margin:5px 0 0 0; color:#146c43; font-size:0.9em; font-weight:bold;"></p>
        </div>
        <script>
            const inputP = document.getElementById('prest_reg');
            const msgP = document.getElementById('msg_prest');
            const guardadoP = localStorage.getItem('llegando_nombre_prestador');
            if (guardadoP) {
                inputP.value = guardadoP;
                msgP.innerText = "✅ Prestador vinculado: " + guardadoP;
            }
            function guardarPrest() {
                if(inputP.value.trim() !== "") {
                    localStorage.setItem('llegando_nombre_prestador', inputP.value.trim());
                    msgP.innerText = "✅ ¡Nombre guardado!";
                    setTimeout(() => { window.location.reload(); }, 800);
                }
            }
            function pedirPermiso() {
                if ("Notification" in window) {
                    Notification.requestPermission().then(permission => {
                        if (permission === "granted") {
                            alert("¡Cintillo de notificaciones activo para nuevos viajes!");
                        }
                    });
                }
            }
        </script>
    """, height=130)

    nom_prestador = st.text_input("Ingresa tu Nombre exacto autorizado para acceder:")
    
    if nom_prestador:
        prestador_valido = df_prestadores[(df_prestadores["Nombre"].str.lower() == nom_prestador.lower()) & (df_prestadores["Estatus"] == "Autorizado")]
        
        if len(prestador_valido) == 0:
            st.warning("⚠️ Tu nombre no se encuentra registrado o aún no ha sido autorizado.")
        else:
            st.success("✅ ¡Acceso Autorizado!")
            
            if st.button("🔊 Activar Alarma de Nuevos Viajes"):
                audio_init_prest = "<script>try { const ac = new (window.AudioContext || window.webkitAudioContext)(); const osc = ac.createOscillator(); const g = ac.createGain(); osc.type='square'; osc.frequency.value=750; g.gain.setValueAtTime(0.12, ac.currentTime); osc.connect(g); osc.connect(ac.destination); osc.start(); osc.stop(ac.currentTime + 0.3); } catch(e){}</script>"
                components.html(audio_init_prest, height=0)
                st.toast("🚀 ¡Cintillo y alarma de prestador activos!", icon="🚗")

            realtime_prestador_js = """
                <script>
                    setInterval(async () => {
                        try {
                            const res = await fetch(window.location.href);
                            const html = await res.text();
                            if (html.includes("Servicio #") || html.includes("Asignado")) {
                                if (Notification.permission === "granted") {
                                    new Notification("🚗 LLEGANDO - Nuevo Viaje", {
                                        body: "¡Se te ha asignado un nuevo servicio en la plataforma!",
                                        icon: "https://cdn-icons-png.flaticon.com/512/7541/7541900.png"
                                    });
                                }
                                const ac = new (window.AudioContext || window.webkitAudioContext)();
                                const osc = ac.createOscillator();
                                const g = ac.createGain();
                                osc.type='square'; osc.frequency.value=750;
                                g.gain.setValueAtTime(0.12, ac.currentTime);
                                osc.connect(g); g.connect(ac.destination);
                                osc.start(); osc.stop(ac.currentTime + 0.6);
                            }
                        } catch (e) {}
                    }, 4000);
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
                        estatus_conf = str(liq['Estatus_Confirmacion']).strip()
                        
                        st.markdown(f"### Liquidación #{liq['ID_Liquidacion']} - Fecha: {liq['Fecha']}")
                        st.write(f"💵 **Monto Pagado:** ${monto_liq_usd:,.2f} USD (Bs. {monto_liq_ves:,.2f})")
                        st.write(f"📝 **Referencia / Detalle:** `{liq['Referencia_Pago']}`")
                        st.write(f"🚦 **Estatus de Confirmación:** `{estatus_conf}`")
                        
                        if estatus_conf.lower() != "confirmado":
                            if st.button(f"✅ Confirmar Recepción #{liq['ID_Liquidacion']}", key=f"conf_liq_{liq['ID_Liquidacion']}"):
                                df_liquidaciones.at[idx_l, "Estatus_Confirmacion"] = "Confirmado"
                                guardar_liquidaciones(df_liquidaciones)
                                st.success("¡Has confirmado la recepción del pago con éxito!")
                                st.rerun()
                        else:
                            st.success("✅ Pago confirmado por ti.")
                        st.markdown("---")
                else:
                    st.info("No tienes liquidaciones registradas todavía.")

# ==========================================
# 3. VISTA ADMINISTRADOR (PANEL GERENCIAL)
# ==========================================
elif "Admin" in rol or "Gerencial" in rol:
    st.markdown("<h2 style='color: #198754;'>⚙️ Panel Gerencial LLEGANDO (Seguro)</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    clave_ingresada = st.text_input("🔑 Ingresa tu Contraseña de Administrador:", type="password")
    
    if clave_ingresada == CLAVE_ADMIN_SECRETA:
        st.success("✅ ¡Acceso Gerencial Autorizado!")
        st.markdown("---")
        
        col_adm_btn1, col_adm_btn2 = st.columns(2)
        with col_adm_btn1:
            if st.button("🔊 Activar Alarma Gerencial"):
                audio_init_adm = "<script>try { const ac = new (window.AudioContext || window.webkitAudioContext)(); const osc = ac.createOscillator(); const g = ac.createGain(); osc.type='sine'; osc.frequency.value=880; g.gain.setValueAtTime(0.1, ac.currentTime); osc.connect(g); osc.connect(ac.destination); osc.start(); osc.stop(ac.currentTime + 0.3); } catch(e){}</script>"
                components.html(audio_init_adm, height=0)
                st.toast("🚀 ¡Alarma gerencial activa!", icon="⚙️")
        with col_adm_btn2:
            if st.button("🔔 Activar Permiso de Cintillo Push"):
                components.html("""
                    <script>
                        if ("Notification" in window) {
                            Notification.requestPermission().then(permission => {
                                if (permission === "granted") {
                                    alert("¡Cintillo de nuevos pedidos activado!");
                                }
                            });
                        }
                    </script>
                """, height=0)
                st.success("¡Solicitud de permiso enviada!")

        realtime_admin_js = """
            <script>
                setInterval(async () => {
                    try {
                        const res = await fetch(window.location.href);
                        const html = await res.text();
                        if (html.includes("Solicitado") || html.includes("Pendiente")) {
                            if (Notification.permission === "granted") {
                                new Notification("⚙️ LLEGANDO - Nuevo Pedido", {
                                    body: "¡Hay una nueva solicitud pendiente de asignación!",
                                    icon: "https://cdn-icons-png.flaticon.com/512/7541/7541900.png"
                                });
                            }
                            const ac = new (window.AudioContext || window.webkitAudioContext)();
                            const osc = ac.createOscillator();
                            const g = ac.createGain();
                            osc.type='sine'; osc.frequency.value=880;
                            g.gain.setValueAtTime(0.1, ac.currentTime);
                            osc.connect(g); g.connect(ac.destination);
                            osc.start(); osc.stop(ac.currentTime + 0.5);
                        }
                    } catch (e) {}
                }, 4000);
            </script>
        """
        components.html(realtime_admin_js, height=0)
        
        pendientes_alarma = df[df["Estado"].isin(["Solicitado", "Pendiente"])]
        conteo_actual_admin = len(pendientes_alarma)

        if conteo_actual_admin > 0:
            st.warning(f"🚨 ¡Atención Gerente! Tienes **{conteo_actual_admin}** solicitud(es) entrante(s) en tiempo real.")

        with st.expander("⚙️ Configuración General del Sistema", expanded=False):
            col_adm1, col_adm2 = st.columns(2)
            with col_adm1:
                st.session_state.tasa_bcv = st.number_input("Tasa BCV Oficial (Bs. / USD)", value=float(st.session_state.tasa_bcv), step=0.10)
            with col_adm2:
                st.session_state.porcentaje_prestador = st.number_input("% Ganancia Prestador", value=float(st.session_state.porcentaje_prestador), step=5.0)
                
        st.markdown("---")
        
        opcion_admin = st.selectbox("Seleccione la sección de administración:", [
            "👥 Autorizar Prestadores", 
            "📥 Asignar y Confirmar Pagos", 
            "💰 Liquidaciones a Prestadores", 
            "📋 Control General y Cancelaciones", 
            "📊 Finanzas", 
            "📁 Excel"
        ])
        
        if opcion_admin == "👥 Autorizar Prestadores":
            st.subheader("👥 Gestión y Autorización de Prestadores de Servicio")
            
            with st.form("form_nuevo_prestador"):
                c_p1, c_p2, c_p3 = st.columns(3)
                with c_p1:
                    nombre_p = st.text_input("Nombre y Apellido *")
                with c_p2:
                    telefono_p = st.text_input("Teléfono *")
                with c_p3:
                    vehiculo_p = st.selectbox("Vehículo / Tipo *", ["Moto", "Carro / Taxi", "Flete / Camión", "Delivery (Moto)", "Delivery (Carro)"])
                    
                if st.form_submit_button("➕ Registrar y Autorizar Prestador"):
                    if nombre_p and telefono_p:
                        nuevo_p = pd.DataFrame([{
                            "Nombre": nombre_p,
                            "Telefono": telefono_p,
                            "Vehiculo": vehiculo_p,
                            "Estatus": "Autorizado"
                        }])
                        df_prestadores = pd.concat([df_prestadores, nuevo_p], ignore_index=True)
                        guardar_prestadores(df_prestadores)
                        st.success(f"¡Prestador {nombre_p} autorizado con éxito!")
                        st.rerun()
                    else:
                        st.error("Completa los campos requeridos.")
                        
            st.markdown("#### Prestadores Registrados en el Sistema:")
            if len(df_prestadores) > 0:
                st.dataframe(df_prestadores, use_container_width=True)
            else:
                st.info("No hay prestadores registrados.")

        elif opcion_admin == "📥 Asignar y Confirmar Pagos":
            st.subheader("📥 Asignación Inteligente y Verificación Bancaria de Pagos")
            pendientes = df[df["Estado"].isin(["Solicitado", "Pendiente", "Asignado"])]
            
            if len(pendientes) > 0:
                prestadores_autorizados = df_prestadores[df_prestadores["Estatus"] == "Autorizado"]
                if len(prestadores_autorizados) == 0:
                    st.warning("⚠️ Debes autorizar al menos un prestador antes de asignar servicios.")
                
                for idx, serv in pendientes.iterrows():
                    tipo_solicitud = str(serv['Tipo_Servicio']).strip().lower()
                    
                    if "moto" in tipo_solicitud:
                        prestadores_filtrados = prestadores_autorizados[prestadores_autorizados["Vehiculo"].str.lower().str.contains("moto")]
                    elif "taxi" in tipo_solicitud or "carro" in tipo_solicitud:
                        prestadores_filtrados = prestadores_autorizados[prestadores_autorizados["Vehiculo"].str.lower().str.contains("carro|taxi")]
                    elif "flete" in tipo_solicitud:
                        prestadores_filtrados = prestadores_autorizados[prestadores_autorizados["Vehiculo"].str.lower().str.contains("flete|camión|camion")]
                    elif "delivery" in tipo_solicitud:
                        prestadores_filtrados = prestadores_autorizados[prestadores_autorizados["Vehiculo"].str.lower().str.contains("delivery|moto|carro")]
                    else:
                        prestadores_filtrados = prestadores_autorizados
                    
                    lista_candidatos = prestadores_filtrados["Nombre"].tolist()
                    if not lista_candidatos:
                        lista_candidatos = prestadores_autorizados["Nombre"].tolist()
                    
                    with st.expander(f"Solicitud #{serv['ID_Servicio']} - {serv['Cliente']} ({serv['Tipo_Servicio']}) | Ref. Pago: {serv['Ref_Pago_Movil']}"):
                        st.write(f"**Origen:** {serv['Origen']} ➡️ **Destino:** {serv['Destino']}")
                        st.write(f"**Teléfono:** {serv['Telefono_Cliente']} | **Método de Pago:** {serv['Metodo_Pago']}")
                        st.write(f"**Referencia Reportada por Cliente:** `{serv['Ref_Pago_Movil']}`")
                        st.write(f"**Estatus de Pago Actual:** `{serv.get('Pago_Verificado', 'No')}`")
                        
                        c_asig1, c_asig2 = st.columns(2)
                        with c_asig1:
                            cond_asignado = st.selectbox(f"Prestadores compatibles para {serv['Tipo_Servicio']}", lista_candidatos, key=f"cond_{serv['ID_Servicio']}")
                        with c_asig2:
                            tarifa_usd = st.number_input(f"Asignar Tarifa USD", min_value=0.0, step=0.5, value=float(serv['Monto_USD']) if pd.notna(serv['Monto_USD']) else 0.0, key=f"tar_{serv['ID_Servicio']}")
                            
                        verif_pago = st.selectbox(f"¿Pago verificado en el banco?", ["No", "Sí"], index=1 if str(serv.get('Pago_Verificado', 'No')).lower() == "sí" else 0, key=f"ver_{serv['ID_Servicio']}")
                            
                        col_act1, col_act2 = st.columns(2)
                        with col_act1:
                            if st.button(f"💾 Guardar Cambios #{serv['ID_Servicio']}", key=f"btn_asig_{serv['ID_Servicio']}"):
                                df.at[idx, "Conductor"] = cond_asignado
                                df.at[idx, "Monto_USD"] = tarifa_usd
                                df.at[idx, "Pago_Verificado"] = verif_pago
                                df.at[idx, "Estado"] = "Asignado"
                                guardar_datos(df)
                                st.success(f"¡Servicio #{serv['ID_Servicio']} actualizado correctamente!")
                                st.rerun()
                        with col_act2:
                            if st.button(f"❌ Cancelar Servicio #{serv['ID_Servicio']}", key=f"btn_canc_{serv['ID_Servicio']}"):
                                df.at[idx, "Estado"] = "Cancelado"
                                guardar_datos(df)
                                st.warning(f"Servicio #{serv['ID_Servicio']} cancelado.")
                                st.rerun()
            else:
                st.info("No hay solicitudes pendientes.")

        elif opcion_admin == "💰 Liquidaciones a Prestadores":
            st.subheader("💰 Control de Liquidación y Pagos por Prestador")
            prestadores_activos = df_prestadores[df_prestadores["Estatus"] == "Autorizado"]["Nombre"].tolist()
            
            if prestadores_activos:
                prestador_sel = st.selectbox("Selecciona al Prestador de Servicio:", prestadores_activos)
                
                servicios_p = df[(df["Conductor"].astype(str).str.strip().str.lower() == prestador_sel.strip().lower()) & (df["Estado"] == "Completado")]
                total_recaudado_p = pd.to_numeric(servicios_p["Monto_USD"], errors='coerce').sum()
                ganancia_prestador_usd = total_recaudado_p * (st.session_state.porcentaje_prestador / 100)
                ganancia_prestador_ves = ganancia_prestador_usd * st.session_state.tasa_bcv
                
                st.markdown(f"### Resumen Financiero para: **{prestador_sel}**")
                c_l1, c_l2, c_l3 = st.columns(3)
                c_l1.metric("Servicios Completados", len(servicios_p))
                c_l2.metric("Total a Liquidar (USD)", f"${ganancia_prestador_usd:,.2f} USD")
                c_l3.metric("Total a Liquidar (VES)", f"Bs. {ganancia_prestador_ves:,.2f}")
                
                with st.form("form_registrar_liquidacion"):
                    st.markdown("#### Registrar Pago de Liquidación al Prestador")
                    monto_liq = st.number_input("Monto a Liquidar USD", value=float(ganancia_prestador_usd), min_value=0.0, step=1.0)
                    ref_liq = st.text_input("Referencia de Pago o Nota de Transferencia / Efectivo *")
                    
                    if st.form_submit_button("📤 Registrar y Enviar Liquidación"):
                        if ref_liq:
                            nuevo_id_liq = 1 if len(df_liquidaciones) == 0 else int(pd.to_numeric(df_liquidaciones["ID_Liquidacion"], errors='coerce').max()) + 1
                            nueva_liq = pd.DataFrame([{
                                "ID_Liquidacion": nuevo_id_liq,
                                "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Prestador": prestador_sel,
                                "Monto_USD": monto_liq,
                                "Referencia_Pago": ref_liq,
                                "Estatus_Confirmacion": "Pendiente de Confirmación"
                            }])
                            df_liquidaciones = pd.concat([df_liquidaciones, nueva_liq], ignore_index=True)
                            guardar_liquidaciones(df_liquidaciones)
                            st.success("¡Liquidación registrada con éxito!")
                            st.rerun()
                        else:
                            st.error("Por favor ingresa la referencia de pago.")
                
                st.markdown("---")
                st.markdown("#### Historial de Liquidaciones Registradas:")
                liq_hist = df_liquidaciones[df_liquidaciones["Prestador"].astype(str).str.strip().str.lower() == prestador_sel.strip().lower()]
                if len(liq_hist) > 0:
                    st.dataframe(liq_hist, use_container_width=True)
                else:
                    st.info("No hay liquidaciones registradas para este prestador.")
            else:
                st.warning("No hay prestadores autorizados disponibles.")

        elif opcion_admin == "📋 Control General y Cancelaciones":
            st.subheader("📋 Control General de Servicios y Cancelaciones")
            if len(df) > 0:
                st.dataframe(df.astype(str), use_container_width=True)
                
                st.markdown("---")
                st.markdown("#### ⚙️ Gestión de Cancelación Directa")
                id_a_cancelar = st.selectbox("Selecciona el ID del Servicio a Cancelar", df["ID_Servicio"].unique())
                if st.button("🚫 Marcar Servicio como Cancelado"):
                    idx = df[df["ID_Servicio"] == id_a_cancelar].index
                    df.at[idx[0], "Estado"] = "Cancelado"
                    guardar_datos(df)
                    st.success(f"¡El servicio #{id_a_cancelar} ha sido cancelado exitosamente!")
                    st.rerun()
            else:
                st.info("No hay servicios registrados.")

        elif opcion_admin == "📊 Finanzas":
            st.subheader("📊 Finanzas Generales de la Empresa")
            completados_tot = df[df["Estado"] == "Completado"]
            if len(completados_tot) > 0:
                total_recaudado = pd.to_numeric(completados_tot["Monto_USD"], errors='coerce').sum()
                empresa_pct = 100.0 - st.session_state.porcentaje_prestador
                ganancia_empresa_usd = total_recaudado * (empresa_pct / 100)
                
                c_f1, c_f2 = st.columns(2)
                c_f1.metric("Recaudado Total USD", f"${total_recaudado:,.2f} USD")
                c_f2.metric(f"Comisión Empresa ({empresa_pct}%)", f"${ganancia_empresa_usd:,.2f} USD")
            else:
                st.info("Aún no hay servicios completados.")

        elif opcion_admin == "📁 Excel":
            st.subheader("📁 Descargar Bases de Datos")
            if os.path.exists(DB_FILE):
                with open(DB_FILE, "rb") as file:
                    st.download_button("⬇️ Descargar Base de Servicios", file, file_name="base_datos_servicios.xlsx")
            if os.path.exists(DB_PRESTADORES):
                with open(DB_PRESTADORES, "rb") as file:
                    st.download_button("⬇️ Descargar Base de Prestadores", file, file_name="prestadores_autorizados.xlsx")
            if os.path.exists(DB_LIQUIDACIONES):
                with open(DB_LIQUIDACIONES, "rb") as file:
                    st.download_button("⬇️ Descargar Base de Liquidaciones", file, file_name="liquidaciones_prestadores.xlsx")
    elif clave_ingresada != "":
        st.error("❌ Contraseña incorrecta. Acceso denegado.")
    else:
        st.info("🔒 Ingresa tu contraseña en el campo superior para desbloquear el panel gerencial.")

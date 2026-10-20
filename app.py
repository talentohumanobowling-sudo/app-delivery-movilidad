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
        
        # Registro automático de identidad por dispositivo
        components.html("""
            <div style="background:#E8F5E9; padding:12px; border-radius:8px; border:1px solid #198754; margin-bottom:15px; font-family:sans-serif;">
                <p style="margin:0 0 8px 0; color:#0F5132; font-weight:600;">📱 Tu Identidad en este Dispositivo:</p>
                <input type="text" id="tel_reg" placeholder="Ingresa tu número (ej: 04141234567)" style="width:70%; padding:6px; border-radius:4px; border:1px solid #ccc;">
                <button onclick="guardarTel()" style="background:#198754; color:white; border:none; padding:7px 12px; border-radius:4px; font-weight:bold; cursor:pointer;">Guardar y Recordar</button>
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
                        msg.innerText = "✅ ¡Teléfono guardado!";
                        setTimeout(() => { window.location.reload(); }, 800);
                    }
                }
            </script>
        """, height=120)

        telefono_busqueda = st.text_input("Confirma tu número de teléfono registrado:")
        
        if telefono_busqueda:
            if st.button("🔊 Activar Mensajería Flotante en Tiempo Real"):
                audio_init_cli = """
                    <script>
                        try {
                            const ac = new (window.AudioContext || window.webkitAudioContext)();
                            const osc = ac.createOscillator();
                            const g = ac.createGain();
                            osc.type = 'triangle';
                            osc.frequency.value = 600;
                            g.gain.setValueAtTime(0.15, ac.currentTime);
                            osc.connect(g);
                            osc.connect(ac.destination);
                            osc.start();
                            osc.stop(ac.currentTime + 0.3);
                        } catch(e) {}
                    </script>
                """
                components.html(audio_init_cli, height=0)
                st.success("¡Mensajería flotante activada!")

            # Cintillo flotante dinámico para el Cliente
            realtime_cliente_js = """
                <div id="cintillo-cliente" style="display:none; position:fixed; top:10px; left:50%; transform:translateX(-50%); z-index:999999; background:#0F5132; color:white; padding:14px 22px; border-radius:10px; box-shadow:0 6px 15px rgba(0,0,0,0.3); font-family:sans-serif; text-align:center; width:90%; max-width:500px; border:2px solid #198754;">
                    <div style="font-weight:bold; font-size:1.1em; margin-bottom:4px;">🚀 LLEGANDO • Actualización de Viaje</div>
                    <div style="font-size:0.95em;">¡Su viaje ha sido actualizado por la administración!</div>
                </div>
                <script>
                    const miTlf = "%s";
                    setInterval(async () => {
                        try {
                            const res = await fetch(window.location.href);
                            const html = await res.text();
                            if (html.includes(miTlf) && (html.includes("Pago Confirmado") || html.includes("Asignado") || html.includes("Completado"))) {
                                const cintillo = document.getElementById('cintillo-cliente');
                                if(cintillo) { cintillo.style.display = 'block'; }
                                const ac = new (window.AudioContext || window.webkitAudioContext)();
                                const osc = ac.createOscillator();
                                const g = ac.createGain();
                                osc.type = 'triangle';
                                osc.frequency.value = 600;
                                g.gain.setValueAtTime(0.15, ac.currentTime);
                                osc.connect(g);
                                g.connect(ac.destination);
                                osc.start();
                                osc.stop(ac.currentTime + 0.5);
                            }
                        } catch (e) {}
                    }, 4000);
                </script>
            """ % telefono_busqueda.strip()
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
                            st.markdown(f"**💵 Tarifa:**

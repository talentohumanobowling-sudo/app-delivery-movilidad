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
DB_ALIADOS = "aliados_comerciales.xlsx"
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

def guardar_datos(df):
    df["Ref_Pago_Movil"] = df["Ref_Pago_Movil"].astype(str)
    df["Telefono_Cliente"] = df["Telefono_Cliente"].astype(str)
    df["Liquidado"] = df["Liquidado"].astype(str)
    df["Valoracion"] = df["Valoracion"].astype(str)
    df["Aliado_Comercial"] = df["Aliado_Comercial"].astype(str)
    df["Capture_Pago"] = df["Capture_Pago"].astype(str)
    df.to_excel(DB_FILE, index=False)

def cargar_prestadores():
    if os.path.exists(DB_PRESTADORES):
        df_p_temp = pd.read_excel(DB_PRESTADORES, dtype=str)
        cols_necesarias = ["Nombre", "Telefono", "Vehiculo", "Marca", "Modelo", "Placa", "Foto_Path", "Pago_Afiliacion_5USD", "Pago_Mensualidad_5USD", "Estatus", "Saldo_Wallet"]
        for col in cols_necesarias:
            if col not in df_p_temp.columns:
                df_p_temp[col] = "0.0" if col == "Saldo_Wallet" else ("Pendiente" if "Pago" in col else "")
        return df_p_temp
    else:
        df_p = pd.DataFrame(columns=["Nombre", "Telefono", "Vehiculo", "Marca", "Modelo", "Placa", "Foto_Path", "Pago_Afiliacion_5USD", "Pago_Mensualidad_5USD", "Estatus", "Saldo_Wallet"])
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

def cargar_aliados():
    if os.path.exists(DB_ALIADOS):
        df_a = pd.DataFrame(columns=["Nombre", "Ubicacion", "Telefono", "Logo_Path"])
        df_a.to_excel(DB_ALIADOS, index=False)
        return df_a
    else:
        df_a_temp = pd.read_excel(DB_ALIADOS, dtype=str)
        cols_a = ["Nombre", "Ubicacion", "Telefono", "Logo_Path"]
        for col in cols_a:
            if col not in df_a_temp.columns:
                df_a_temp[col] = ""
        return df_a_temp

def guardar_aliados(df_a):
    df_a.to_excel(DB_ALIADOS, index=False)

df = cargar_datos()
df_prestadores = cargar_prestadores()
df_liquidaciones = cargar_liquidaciones()
df_aliados = cargar_aliados()

# ==========================================
# BARRA LATERAL
# ==========================================
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

# ==========================================
# 1. VISTA CLIENTE
# ==========================================
if "Portal de Cliente" in rol:
    st.markdown("<h2>LLEGANDO - Solicitud y Estatus</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("""
        <div class="banner-instalacion">
            <h3 style="color: white; margin-top: 0;">Lleva LLEGANDO en la pantalla de tu telefono</h3>
            <p>Instala nuestra aplicacion web:</p>
            <ul>
                <li><b>Android (Chrome):</b> Toca el menu (⋮) y selecciona <b>"Instalar aplicacion"</b> o <b>"Agregar a la pantalla principal"</b>.</li>
                <li><b>iPhone (Safari):</b> Toca el boton <b>Compartir</b> y selecciona <b>"Agregar a inicio"</b>.</li>
            </ul>
        </div>
    """, unsafe_allow_html=True)
    
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
                    nuevo_id = 1001 if len(df) == 0 else int(pd.to_numeric(df["ID_Servicio"], errors='coerce').max()) + 1
                    nuevo_registro = {
                        "ID_Servicio": nuevo_id, "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Tipo_Servicio": tipo_serv, "Cliente": nombre, "Telefono_Cliente": str(telefono).strip(),
                        "Origen": origen, "Destino": destino, "Conductor": "Por Asignar", "Monto_USD": tarifa_sugerida,
                        "Metodo_Pago": metodo_pago, "Ref_Pago_Movil": "", "Pago_Verificado": "No",
                        "Estado": "Solicitado", "Liquidado": "No", "Valoracion": "Sin valorar", "Notas": notas, 
                        "Aliado_Comercial": "Particular", "Capture_Pago": ""
                    }
                    df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
                    guardar_datos(df)
                    st.success(f"¡Solicitud enviada! Tu numero de ID es #{nuevo_id}.")
                else:
                    st.error("Completa los campos obligatorios (*).")
                    
    with tab_consultar:
        st.subheader("Mis Servicios Activos e Historial")
        telefono_busqueda = st.text_input("Ingresa tu numero de telefono registrado:")
        
        if telefono_busqueda:
            resultados = df[df["Telefono_Cliente"].astype(str).str.strip() == telefono_busqueda.strip()]
            if len(resultados) > 0:
                resultados = resultados.sort_values(by="ID_Servicio", ascending=False)
                for idx_res, serv in resultados.iterrows():
                    monto_usd = float(serv['Monto_USD'])
                    monto_ves = monto_usd * st.session_state.tasa_bcv
                    estatus_s = str(serv['Estado']).strip()
                    pago_v = str(serv.get('Pago_Verificado', 'No')).strip()
                    
                    if pago_v.lower() in ["si", "sí"] or estatus_s in ["Asignado", "En Camino", "Llegue al Sitio", "En Trayecto", "Completado"]:
                        st.markdown(f'''
                            <div class="cintillo-alerta">
                                <h3 style="color: #ffffff; margin: 0 0 5px 0;">Actualizacion en tu Viaje #{serv["ID_Servicio"]}</h3>
                                <p style="margin: 0;">Estatus: <b>{estatus_s}</b> | Prestador: <b>{serv["Conductor"]}</b> | Pago: <b>{pago_v}</b></p>
                            </div>
                            <audio autoplay style="display:none;"><source src="https://assets.mixkit.co/active_storage/sfx/2869/2869-preview.mp3" type="audio/mpeg"></audio>
                        ''', unsafe_allow_html=True)

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
                            ref_actual = "" if pd.isna(serv['Ref_Pago_Movil']) or str(serv['Ref_Pago_Movil']) == "nan" else str(serv['Ref_Pago_Movil']).strip()
                            if str(serv['Metodo_Pago']).strip().lower() == "pago movil" and monto_usd > 0:
                                st.markdown("---")
                                st.info("Datos para Pago Movil:")
                                st.warning(f"Monto exacto: Bs. {monto_ves:,.2f}")
                                st.write(f"Banco: {DATOS_PAGO_MOVIL['Banco']} | Telefono: {DATOS_PAGO_MOVIL['Telefono']}")
                                st.write(f"Cedula/RIF: {DATOS_PAGO_MOVIL['Cedula_RIF']} | Titular: {DATOS_PAGO_MOVIL['Titular']}")
                                
                                ref_ingresada = st.text_input("Ingresa los digitos de tu referencia:", value=ref_actual, key=f"ref_{serv['ID_Servicio']}")
                                capture_file = st.file_uploader("Sube el capture o comprobante de pago (JPG/PNG)", type=["jpg", "jpeg", "png"], key=f"cap_{serv['ID_Servicio']}")
                                
                                if st.button(f"Reportar Pago y Referencia #{serv['ID_Servicio']}", key=f"btn_ref_{serv['ID_Servicio']}"):
                                    if ref_ingresada:
                                        path_capture = str(serv.get('Capture_Pago', ''))
                                        if capture_file is not None:
                                            os.makedirs("comprobantes_pago", exist_ok=True)
                                            path_capture = os.path.join("comprobantes_pago", f"serv_{serv['ID_Servicio']}.jpg")
                                            with open(path_capture, "wb") as f:
                                                f.write(capture_file.getbuffer())
                                        
                                        idx = df[df["ID_Servicio"] == serv['ID_Servicio']].index
                                        df.at[idx[0], "Ref_Pago_Movil"] = str(ref_ingresada)
                                        df.at[idx[0], "Capture_Pago"] = path_capture
                                        guardar_datos(df)
                                        st.success("¡Referencia y comprobante reportados con éxito al administrador!")
                                        st.rerun()
                                    else:
                                        st.error("Ingresa al menos la referencia de pago.")
                            
                            c_pago_path = str(serv.get('Capture_Pago', ''))
                            if c_pago_path and c_pago_path != "nan" and os.path.exists(c_pago_path):
                                st.image(c_pago_path, width=200, caption="Comprobante de Pago Subido")

                            if pago_v.lower() in ["si", "sí"] or str(serv['Metodo_Pago']).strip().lower() != "pago movil":
                                st.markdown("---")
                                st.success("Pago Confirmado por la Empresa")
                                st.markdown(f"### Prestador Asignado: **{serv['Conductor']}**")
                                
                                info_cond = df_prestadores[df_prestadores["Nombre"].str.lower() == str(serv['Conductor']).lower()]
                                if len(info_cond) > 0:
                                    cond_row = info_cond.iloc[0]
                                    foto_path = str(cond_row.get('Foto_Path', '')).strip()
                                    tlf_cond = str(cond_row['Telefono']).strip().replace(" ", "").replace("+", "")
                                    
                                    col_info1, col_info2 = st.columns([1, 2])
                                    with col_info1:
                                        if foto_path and foto_path != "nan" and os.path.exists(foto_path):
                                            st.image(foto_path, width=120)
                                    with col_info2:
                                        st.markdown(f"Vehiculo: {cond_row['Vehiculo']} ({cond_row.get('Marca', '')} {cond_row.get('Modelo', '')})")
                                        st.markdown(f"Placa: `{cond_row.get('Placa', '')}`")
                                    
                                    url_waze = f"https://waze.com/ul?q={urllib.parse.quote(serv['Origen'])}"
                                    url_gmaps = f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(serv['Origen'])}"
                                    st.link_button("🗺️ Abrir Ruta en Waze", url_waze)
                                    st.link_button("📍 Abrir Ruta en Google Maps", url_gmaps)
                                    
                                    msg_c = f"Hola {serv['Conductor']}, le escribo desde LLEGANDO respecto al servicio #{serv['ID_Servicio']}."
                                    st.link_button("Contactar Conductor por WhatsApp", f"https://wa.me/{tlf_cond}?text={urllib.parse.quote(msg_c)}")
                                
                                if estatus_s == "Completado":
                                    st.markdown("---")
                                    val_actual = str(serv.get("Valoracion", "Sin valorar"))
                                    escalas = ["Sin valorar", "Mal servicio", "Regular servicio", "Buen servicio", "Excelente servicio"]
                                    idx_val = escalas.index(val_actual) if val_actual in escalas else 0
                                    nueva_valoracion = st.selectbox("Valora tu servicio:", escalas, index=idx_val, key=f"val_{serv['ID_Servicio']}")
                                    if st.button(f"Guardar Valoracion #{serv['ID_Servicio']}", key=f"btn_val_{serv['ID_Servicio']}"):
                                        idx = df[df["ID_Servicio"] == serv['ID_Servicio']].index
                                        df.at[idx[0], "Valoracion"] = nueva_valoracion
                                        guardar_datos(df)
                                        st.success("¡Gracias por valorar!")
                                        st.rerun()
            else:
                st.warning("No se encontraron servicios con ese numero.")
        time.sleep(3)
        st.rerun()

# ==========================================
# 2. VISTA PRESTADOR DE SERVICIOS
# ==========================================
elif "Portal de Prestador" in rol:
    st.markdown("<h2>Panel Profesional del Prestador</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    st.markdown("""
        <div class="banner-instalacion">
            <h3 style="color: white; margin-top: 0;">Instala la App de Prestadores</h3>
            <p>Android (Menú ⋮ ➡️ Instalar) | iPhone (Compartir ➡️ Agregar a inicio)</p>
        </div>
    """, unsafe_allow_html=True)

    nom_prestador = st.text_input("Ingresa tu Nombre exacto autorizado:")
    if nom_prestador:
        prestador_valido = df_prestadores[(df_prestadores["Nombre"].str.lower() == nom_prestador.strip().lower()) & (df_prestadores["Estatus"] == "Autorizado")]
        if len(prestador_valido) == 0:
            st.warning("Tu cuenta no esta autorizada o no existe.")
        else:
            st.success("Acceso Autorizado")
            p_row = prestador_valido.iloc[0]
            
            saldo_wallet = float(p_row.get("Saldo_Wallet", 0.0) if pd.notna(p_row.get("Saldo_Wallet", 0.0)) else 0.0)
            
            st.markdown("### 💼 Billetera Virtual y Membresías")
            col_st1, col_st2, col_st3 = st.columns(3)
            col_st1.metric("Afiliacion ($5)", str(p_row.get("Pago_Afiliacion_5USD", "Pendiente")))
            col_st2.metric("Mensualidad ($5)", str(p_row.get("Pago_Mensualidad_5USD", "Pendiente")))
            col_st3.metric("💳 Saldo Billetera (Wallet)", f"${saldo_wallet:,.2f} USD")
            
            st.markdown("---")

            servicios_cond = df[df["Conductor"].astype(str).str.strip().str.lower() == nom_prestador.strip().lower()]
            viajes_activos_cond = servicios_cond[servicios_cond["Estado"].isin(["Asignado", "En Camino", "Llegue al Sitio", "En Trayecto"])]
            
            if len(viajes_activos_cond) > 0:
                st.markdown(f'''
                    <div class="cintillo-alerta">
                        <h3 style="color: #ffffff; margin: 0 0 5px 0;">¡Nuevo Servicio Asignado!</h3>
                        <p style="margin: 0;">Tienes <b>{len(viajes_activos_cond)}</b> viaje(s) activo(s).</p>
                    </div>
                    <audio autoplay style="display:none;"><source src="https://assets.mixkit.co/active_storage/sfx/2869/2869-preview.mp3" type="audio/mpeg"></audio>
                ''', unsafe_allow_html=True)

            tab_viajes, tab_finanzas, tab_liquidaciones = st.tabs(["Mis Viajes Activos", "Mis Ganancias", "Liquidaciones"])
            
            with tab_viajes:

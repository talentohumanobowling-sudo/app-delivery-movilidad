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
        return df_temp
    else:
        df_inicial = pd.DataFrame(columns=[
            "ID_Servicio", "Fecha_Hora", "Tipo_Servicio", "Cliente", 
            "Telefono_Cliente", "Origen", "Destino", "Conductor", 
            "Monto_USD", "Metodo_Pago", "Ref_Pago_Movil", "Pago_Verificado", "Estado", "Liquidado", "Valoracion", "Notas", "Aliado_Comercial"
        ])
        df_inicial.to_excel(DB_FILE, index=False)
        return df_inicial

def guardar_datos(df):
    df["Ref_Pago_Movil"] = df["Ref_Pago_Movil"].astype(str)
    df["Telefono_Cliente"] = df["Telefono_Cliente"].astype(str)
    df["Liquidado"] = df["Liquidado"].astype(str)
    df["Valoracion"] = df["Valoracion"].astype(str)
    df["Aliado_Comercial"] = df["Aliado_Comercial"].astype(str)
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
        return pd.read_excel(DB_ALIADOS, dtype=str)
    else:
        df_a = pd.DataFrame(columns=["Nombre", "Ubicacion", "Telefono", "Logo_Path"])
        df_a.to_excel(DB_ALIADOS, index=False)
        return df_a

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
                        "Estado": "Solicitado", "Liquidado": "No", "Valoracion": "Sin valorar", "Notas": notas, "Aliado_Comercial": "Particular"
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
                                if st.button(f"Reportar Referencia #{serv['ID_Servicio']}", key=f"btn_ref_{serv['ID_Servicio']}"):
                                    if ref_ingresada:
                                        idx = df[df["ID_Servicio"] == serv['ID_Servicio']].index
                                        df.at[idx[0], "Ref_Pago_Movil"] = str(ref_ingresada)
                                        guardar_datos(df)
                                        st.success("¡Referencia reportada al administrador!")
                                        st.rerun()
                                    else:
                                        st.error("Ingresa la referencia.")
                            
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
                if len(servicios_cond) > 0:
                    for idx, serv in servicios_cond.iterrows():
                        monto_usd = float(serv['Monto_USD'])
                        ganancia_usd = monto_usd * (st.session_state.porcentaje_prestador / 100)
                        aliado_info = f" [Aliado: {serv.get('Aliado_Comercial', 'Particular')}]" if serv.get('Aliado_Comercial', 'Particular') != 'Particular' else ""
                        
                        with st.expander(f"Servicio #{serv['ID_Servicio']}{aliado_info} - [{serv['Estado']}]"):
                            st.write(f"Cliente: {serv['Cliente']} | Telefono: {serv['Telefono_Cliente']}")
                            st.write(f"Origen: {serv['Origen']} ➡️ Destino: {serv['Destino']}")
                            st.write(f"Tu Ganancia: **${ganancia_usd:.2f} USD**")
                            
                            url_waze_c = f"https://waze.com/ul?q={urllib.parse.quote(serv['Origen'])}"
                            url_gmaps_c = f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(serv['Origen'])}"
                            st.link_button("🗺️ Navegar al Origen con Waze", url_waze_c)
                            st.link_button("📍 Navegar al Origen con Google Maps", url_gmaps_c)
                            
                            estados_posibles = ["Asignado", "En Camino", "Llegue al Sitio", "En Trayecto", "Completado", "Cancelado"]
                            idx_est = estados_posibles.index(serv['Estado']) if serv['Estado'] in estados_posibles else 0
                            nuevo_estado = st.selectbox("Actualizar Estatus:", estados_posibles, index=idx_est, key=f"st_{serv['ID_Servicio']}")
                            
                            col_b1, col_b2 = st.columns(2)
                            with col_b1:
                                if st.button(f"Guardar #{serv['ID_Servicio']}", key=f"btn_{serv['ID_Servicio']}"):
                                    df.at[idx, "Estado"] = nuevo_estado
                                    guardar_datos(df)
                                    st.success("Estatus actualizado")
                                    st.rerun()
                            with col_b2:
                                tlf_c = str(serv['Telefono_Cliente']).strip().replace(" ", "").replace("+", "")
                                st.link_button("WhatsApp", f"https://wa.me/{tlf_c}")
                else:
                    st.info("No tienes servicios asignados.")
                    
            with tab_finanzas:
                completados = servicios_cond[servicios_cond["Estado"] == "Completado"]
                tot_usd = pd.to_numeric(completados["Monto_USD"], errors='coerce').sum() * (st.session_state.porcentaje_prestador / 100)
                st.metric("Ganancia Total USD", f"${tot_usd:,.2f}")
                if len(completados) > 0:
                    st.dataframe(completados[["ID_Servicio", "Fecha_Hora", "Tipo_Servicio", "Monto_USD", "Valoracion"]], use_container_width=True)

            with tab_liquidaciones:
                liq_prestador = df_liquidaciones[df_liquidaciones["Prestador"].astype(str).str.strip().str.lower() == nom_prestador.strip().lower()]
                if len(liq_prestador) > 0:
                    for idx_l, liq in liq_prestador.iterrows():
                        st.write(f"Liquidacion #{liq['ID_Liquidacion']} - ${float(liq['Monto_USD']):,.2f} USD | Ref: {liq['Referencia_Pago']} | Estatus: {liq['Estatus_Confirmacion']}")
                        if str(liq['Estatus_Confirmacion']).lower() != "confirmado":
                            if st.button(f"Confirmar Recepcion #{liq['ID_Liquidacion']}", key=f"conf_liq_{liq['ID_Liquidacion']}"):
                                df_liquidaciones.at[idx_l, "Estatus_Confirmacion"] = "Confirmado"
                                guardar_liquidaciones(df_liquidaciones)
                                st.success("¡Confirmado!")
                                st.rerun()
                else:
                    st.info("Sin liquidaciones registradas.")
        time.sleep(3)
        st.rerun()

# ==========================================
# 3. VISTA ALIADO COMERCIAL (COMERCIO / RESTAURANTE)
# ==========================================
elif "Portal de Aliado Comercial" in rol:
    st.markdown("<h2>🏬 Portal de Aliados Comerciales - LLEGANDO Delivery</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    # Mostrar comercios registrados autorizados si existen
    nombres_aliados_reg = df_aliados["Nombre"].tolist() if len(df_aliados) > 0 else []
    
    if len(nombres_aliados_reg) > 0:
        nombre_aliado = st.selectbox("Seleccione su Comercio / Negocio Aliado:", nombres_aliados_reg)
    else:
        nombre_aliado = st.text_input("Nombre de su Negocio / Comercio Aliado:")
    
    if nombre_aliado:
        # Buscar info del aliado
        info_aliado = df_aliados[df_aliados["Nombre"].str.lower() == nombre_aliado.strip().lower()]
        ubicacion_comercio = ""
        if len(info_aliado) > 0:
            comercio_row = info_aliado.iloc[0]
            ubicacion_comercio = str(comercio_row.get("Ubicacion", ""))
            logo_comercio = str(comercio_row.get("Logo_Path", ""))
            if logo_comercio and logo_comercio != "nan" and os.path.exists(logo_comercio):
                st.image(logo_comercio, width=100)
            st.write(f"📍 **Ubicación:** {ubicacion_comercio} | 📞 **Teléfono:** {comercio_row.get('Telefono', '')}")
        
        st.success(f"✅ Bienvenido, comercio: **{nombre_aliado}**")
        
        tab_pedido_aliado, tab_historial_aliado = st.tabs(["📦 Solicitar Delivery Inmediato", "📋 Mis Deliveries Registrados"])
        
        with tab_pedido_aliado:
            with st.form("form_aliado_comercial"):
                st.markdown("#### Datos de la Orden y Entrega")
                c_al1, c_al2 = st.columns(2)
                with c_al1:
                    cliente_final = st.text_input("Nombre del Cliente Final *")
                    telefono_final = st.text_input("Teléfono del Cliente Final *")
                    monto_pedido = st.number_input("Monto del Delivery USD ($)", min_value=1.0, value=3.0, step=0.5)
                with c_al2:
                    direccion_origen_comercio = st.text_input("Dirección de su Local / Comercio *", value=ubicacion_comercio)
                    direccion_destino_cliente = st.text_input("Dirección de Entrega (Cliente) *")
                    metodo_pago_comercio = st.selectbox("Método de Pago del Pedido", ["Pago Móvil", "Efectivo", "Zelle", "Pagado al Comercio"])
                
                detalles_orden = st.text_area("Descripción del pedido o notas especiales (Ej. Comida caliente, frágil, etc.)")
                
                if st.form_submit_button("🚀 Solicitar Motorizado / Delivery"):
                    if cliente_final and telefono_final and direccion_destino_cliente:
                        nuevo_id = 1001 if len(df) == 0 else int(pd.to_numeric(df["ID_Servicio"], errors='coerce').max()) + 1
                        nuevo_registro_aliado = {
                            "ID_Servicio": nuevo_id,
                            "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Tipo_Servicio": "Delivery",
                            "Cliente": cliente_final,
                            "Telefono_Cliente": str(telefono_final).strip(),
                            "Origen": direccion_origen_comercio if direccion_origen_comercio else f"Local de {nombre_aliado}",
                            "Destino": direccion_destino_cliente,
                            "Conductor": "Por Asignar",
                            "Monto_USD": monto_pedido,
                            "Metodo_Pago": metodo_pago_comercio,
                            "Ref_Pago_Movil": "",
                            "Pago_Verificado": "Sí",
                            "Estado": "Solicitado",
                            "Liquidado": "No",
                            "Valoracion": "Sin valorar",
                            "Notas": detalles_orden,
                            "Aliado_Comercial": nombre_aliado
                        }
                        df = pd.concat([df, pd.DataFrame([nuevo_registro_aliado])], ignore_index=True)
                        guardar_datos(df)
                        st.success(f"¡Delivery solicitado con éxito! ID de orden asignado: **#{nuevo_id}**")
                    else:
                        st.error("Por favor completa los campos obligatorios (*).")
                        
        with tab_historial_aliado:
            st.subheader(f"Historial de Deliveries solicitados por {nombre_aliado}")
            mis_pedidos = df[df["Aliado_Comercial"].astype(str).str.strip().str.lower() == nombre_aliado.strip().lower()]
            
            if len(mis_pedidos) > 0:
                for idx_p, ped in mis_pedidos.iterrows():
                    st.markdown(f"""
                        <div class="card-aliado">
                            <b>Orden #{ped['ID_Servicio']}</b> | Cliente: {ped['Cliente']} | Destino: {ped['Destino']} <br>
                            <b>Estatus:</b> <code>{ped['Estado']}</code> | <b>Motorizado Asignado:</b> {ped['Conductor']} | <b>Tarifa:</b> ${float(ped['Monto_USD']):.2f} USD
                        </div>
                    """, unsafe_allow_html=True)
            else:
                st.info("No hay órdenes de delivery registradas con este nombre de comercio.")
    else:
        st.info("💡 Ingresa o selecciona el nombre de tu comercio o restaurante para acceder al panel de pedidos.")

# ==========================================
# 4. VISTA ADMINISTRADOR
# ==========================================
elif "Panel Gerencial" in rol:
    st.markdown("<h2>Panel Gerencial LLEGANDO</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    clave_ingresada = st.text_input("Contrasena de Administrador (admin123):", type="password")
    if clave_ingresada == CLAVE_ADMIN_SECRETA:
        st.success("Acceso Gerencial Autorizado")
        st.markdown("---")
        
        pendientes_alarma = df[df["Estado"].isin(["Solicitado", "Pendiente"])]
        if len(pendientes_alarma) > 0:
            st.markdown(f'''
                <div class="cintillo-alerta">
                    <h3 style="color: #ffffff; margin: 0 0 5px 0;">¡Nuevos Pedidos Entrantes!</h3>
                    <p style="margin: 0;">Hay <b>{len(pendientes_alarma)}</b> solicitud(es) esperando asignacion.</p>
                </div>
                <audio autoplay style="display:none;"><source src="https://assets.mixkit.co/active_storage/sfx/2869/2869-preview.mp3" type="audio/mpeg"></audio>
            ''', unsafe_allow_html=True)

        with st.expander("Configuracion General"):
            st.session_state.tasa_bcv = st.number_input("Tasa BCV", value=float(st.session_state.tasa_bcv))
            st.session_state.porcentaje_prestador = st.number_input("% Ganancia Prestador", value=float(st.session_state.porcentaje_prestador))

        opcion_admin = st.selectbox("Seccion:", ["Autorizar Prestadores", "Gestionar Aliados Comerciales", "Asignar y Confirmar Pagos", "Liquidaciones a Prestadores", "Control General", "Finanzas", "Excel"])
        
        if opcion_admin == "Autorizar Prestadores":
            with st.form("nuevo_p"):
                nombre_p = st.text_input("Nombre *")
                telefono_p = st.text_input("Telefono *")
                vehiculo_p = st.selectbox("Vehiculo *", ["Moto", "Carro / Taxi", "Flete", "Delivery"])
                placa_p = st.text_input("Placa *")
                pago_afi = st.selectbox("Afiliacion ($5)", ["Pendiente", "Pagado"])
                pago_men = st.selectbox("Mensualidad ($5)", ["Pendiente", "Pagado"])
                saldo_ini = st.number_input("Saldo Inicial Billetera Wallet ($)", value=0.0)
                foto_file = st.file_uploader("Foto", type=["jpg", "png"])
                
                if st.form_submit_button("Registrar Prestador"):
                    if nombre_p and telefono_p and placa_p:
                        path_g = ""
                        if foto_file:
                            os.makedirs("fotos_prestadores", exist_ok=True)
                            path_g = os.path.join("fotos_prestadores", f"{nombre_p.replace(' ', '_')}.jpg")
                            with open(path_g, "wb") as f: f.write(foto_file.getbuffer())
                        
                        nuevo_p = pd.DataFrame([{
                            "Nombre": nombre_p, "Telefono": telefono_p, "Vehiculo": vehiculo_p,
                            "Marca": "", "Modelo": "", "Placa": placa_p, "Foto_Path": path_g,
                            "Pago_Afiliacion_5USD": pago_afi, "Pago_Mensualidad_5USD": pago_men, 
                            "Estatus": "Autorizado", "Saldo_Wallet": str(saldo_ini)
                        }])
                        df_prestadores = pd.concat([df_prestadores, nuevo_p], ignore_index=True)
                        guardar_prestadores(df_prestadores)
                        st.success("¡Prestador registrado con saldo en wallet!")
                        st.rerun()
                    else:
                        st.error("Campos requeridos vacios.")
            
            if len(df_prestadores) > 0:
                st.markdown("#### Prestadores Registrados y Saldo en Wallet:")
                for idx_p, p_item in df_prestadores.iterrows():
                    with st.expander(f"{p_item['Nombre']} [Wallet: ${p_item.get('Saldo_Wallet', '0.0')}]"):
                        nuevo_saldo_edit = st.number_input(f"Actualizar Saldo Wallet para {p_item['Nombre']}", value=float(p_item.get('Saldo_Wallet', 0.0) if pd.notna(p_item.get('Saldo_Wallet', 0.0)) else 0.0), key=f"edit_wallet_{idx_p}")
                        if st.button("Guardar Nuevo Saldo", key=f"btn_w_{idx_p}"):
                            df_prestadores.at[idx_p, "Saldo_Wallet"] = str(nuevo_saldo_edit)
                            guardar_prestadores(df_prestadores)
                            st.success("¡Saldo de billetera actualizado!")
                            st.rerun()
                        if st.button("Eliminar Prestador", key=f"del_{idx_p}"):
                            df_prestadores = df_prestadores.drop(idx_p).reset_index(drop=True)
                            guardar_prestadores(df_prestadores)
                            st.rerun()

        elif opcion_admin == "Gestionar Aliados Comerciales":
            st.subheader("🏪 Registro y Control de Aliados Comerciales")
            
            with st.form("form_nuevo_aliado"):
                nombre_aliado_n = st.text_input("Nombre del Comercio / Negocio *")
                ubicacion_aliado_n = st.text_input("Ubicación / Dirección del Local *")
                telefono_aliado_n = st.text_input("Teléfono de Contacto *")
                logo_aliado_file = st.file_uploader("Logotipo del Comercio (JPG/PNG)", type=["jpg", "jpeg", "png"])
                
                if st.form_submit_button("➕ Registrar Aliado Comercial"):
                    if nombre_aliado_n and ubicacion_aliado_n and telefono_aliado_n:
                        logo_path_g = ""
                        if logo_aliado_file is not None:
                            os.makedirs("fotos_aliados", exist_ok=True)
                            logo_path_g = os.path.join("fotos_aliados", f"{nombre_aliado_n.strip().replace(' ', '_')}.jpg")
                            with open(logo_path_g, "wb") as f:
                                f.write(logo_aliado_file.getbuffer())
                        
                        nuevo_aliado_df = pd.DataFrame([{
                            "Nombre": nombre_aliado_n,
                            "Ubicacion": ubicacion_aliado_n,
                            "Telefono": telefono_aliado_n,
                            "Logo_Path": logo_path_g
                        }])
                        df_aliados = pd.concat([df_aliados, nuevo_aliado_df], ignore_index=True)
                        guardar_aliados(df_aliados)
                        st.success(f"¡Aliado comercial '{nombre_aliado_n}' registrado con éxito!")
                        st.rerun()
                    else:
                        st.error("Completa los campos obligatorios (*).")
            
            st.markdown("---")
            st.markdown("#### 📋 Comercios Aliados Registrados:")
            if len(df_aliados) > 0:
                for idx_a, aliado_item in df_aliados.iterrows():
                    l_path = str(aliado_item.get("Logo_Path", ""))
                    with st.expander(f"🏬 {aliado_item['Nombre']} - 📍 {aliado_item['Ubicacion']}"):
                        if l_path and l_path != "nan" and os.path.exists(l_path):
                            st.image(l_path, width=80)
                        st.write(f"📞 **Teléfono:** {aliado_item['Telefono']}")
                        st.write(f"📍 **Ubicación:** {aliado_item['Ubicacion']}")
                        if st.button(f"🗑️ Eliminar Aliado", key=f"del_aliado_{idx_a}"):
                            df_aliados = df_aliados.drop(idx_a).reset_index(drop=True)
                            guardar_aliados(df_aliados)
                            st.success("Aliado eliminado.")
                            st.rerun()
            else:
                st.info("No hay aliados comerciales registrados todavía.")

        elif opcion_admin == "Asignar y Confirmar Pagos":
            pendientes = df[df["Estado"].isin(["Solicitado", "Pendiente", "Asignado"])]
            if len(pendientes) > 0:
                p_aut = df_prestadores[df_prestadores["Estatus"] == "Autorizado"]["Nombre"].tolist()
                for idx, serv in pendientes.iterrows():
                    aliado_txt = f" [Aliado: {serv.get('Aliado_Comercial', 'Particular')}]" if serv.get('Aliado_Comercial', 'Particular') != 'Particular' else ""
                    with st.expander(f"#{serv['ID_Servicio']}{aliado_txt} - {serv['Cliente']} ({serv['Tipo_Servicio']})"):
                        st.write(f"Ref: {serv['Ref_Pago_Movil']} | Origen: {serv['Origen']} ➡️ Destino: {serv['Destino']}")
                        c_asig = st.selectbox("Conductor", p_aut, key=f"c_{serv['ID_Servicio']}")
                        tarifa = st.number_input("Tarifa USD", value=float(serv['Monto_USD']), key=f"t_{serv['ID_Servicio']}")
                        v_pago = st.selectbox("Pago verificado", ["No", "Sí"], index=1 if str(serv.get('Pago_Verificado'))=="Sí" else 0, key=f"vp_{serv['ID_Servicio']}")
                        
                        if st.button(f"Guardar #{serv['ID_Servicio']}", key=f"sb_{serv['ID_Servicio']}"):
                            df.at[idx, "Conductor"] = c_asig
                            df.at[idx, "Monto_USD"] = tarifa
                            df.at[idx, "Pago_Verificado"] = v_pago
                            df.at[idx, "Estado"] = "Asignado"
                            guardar_datos(df)
                            st.success("Actualizado")
                            st.rerun()
            else:
                st.info("No hay solicitudes pendientes.")

        elif opcion_admin == "Liquidaciones a Prestadores":
            p_act = df_prestadores[df_prestadores["Estatus"] == "Autorizado"]["Nombre"].tolist()
            if p_act:
                p_sel = st.selectbox("Prestador:", p_act)
                pend_liq = df[(df["Conductor"].astype(str).str.strip().str.lower() == p_sel.strip().lower()) & (df["Estado"] == "Completado") & (df["Liquidado"] != "Sí")]
                tot_liq = pd.to_numeric(pend_liq["Monto_USD"], errors='coerce').sum() * (st.session_state.porcentaje_prestador / 100)
                
                st.metric("Total a Liquidar USD", f"${tot_liq:,.2f}")
                if len(pend_liq) > 0:
                    with st.form("form_liq"):
                        m_liq = st.number_input("Monto USD", value=float(tot_liq))
                        ref_l = st.text_input("Referencia de Pago *")
                        if st.form_submit_button("Registrar Liquidacion"):
                            if ref_l:
                                n_liq_id = 1 if len(df_liquidaciones) == 0 else int(pd.to_numeric(df_liquidaciones["ID_Liquidacion"], errors='coerce').max()) + 1
                                n_l = pd.DataFrame([{
                                    "ID_Liquidacion": n_liq_id, "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "Prestador": p_sel, "Monto_USD": m_liq, "Referencia_Pago": ref_l, "Estatus_Confirmacion": "Pendiente"
                                }])
                                df_liquidaciones = pd.concat([df_liquidaciones, n_l], ignore_index=True)
                                guardar_liquidaciones(df_liquidaciones)
                                for i_s in pend_liq.index: df.at[i_s, "Liquidado"] = "Sí"
                                guardar_datos(df)
                                st.success("Liquidacion registrada")
                                st.rerun()
            else:
                st.warning("No hay prestadores autorizados.")

        elif opcion_admin == "Control General":
            st.dataframe(df.astype(str), use_container_width=True)

        elif opcion_admin == "Finanzas":
            comp = df[df["Estado"] == "Completado"]
            tot = pd.to_numeric(comp["Monto_USD"], errors='coerce').sum()
            st.metric("Total Recaudado USD", f"${tot:,.2f}")

        elif opcion_admin == "Excel":
            if os.path.exists(DB_FILE):
                with open(DB_FILE, "rb") as f: st.download_button("Descargar Servicios", f, file_name="servicios.xlsx")
                
        time.sleep(3)
        st.rerun()
    elif clave_ingresada != "":
        st.error("Contraseña incorrecta.")
    else:
        st.info("Ingresa la contraseña `admin123`.")

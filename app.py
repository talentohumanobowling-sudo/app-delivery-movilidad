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

def reproducir_alerta_sonora():
    audio_html = """
        <audio autoplay>
          <source src="https://assets.mixkit.co/active_storage/sfx/2869/2869-preview.mp3" type="audio/mpeg">
        </audio>
    """
    st.markdown(audio_html, unsafe_allow_html=True)

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

# ==========================================
# 1. VISTA CLIENTE
# ==========================================
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
                    
                    reproducir_alerta_sonora()
                    st.toast(f"¡Solicitud #{nuevo_id} enviada con éxito!", icon="🚨")
                    st.success(f"¡Solicitud enviada! Tu número de ID es #{nuevo_id}.")
                else:
                    st.error("Completa los campos obligatorios (*).")
                    
    with tab_consultar:
        st.subheader("Mis Servicios Activos e Historial")
        telefono_busqueda = st.text_input("Ingresa tu número de teléfono registrado:").strip()
        
        if telefono_busqueda:
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
                                        
                                        reproducir_alerta_sonora()
                                        st.toast("¡Referencia reportada exitosamente!", icon="🔔")
                                        st.success("¡Referencia reportada con éxito!")
                                        st.rerun()
                                    else:
                                        st.error("Ingresa al menos los dígitos de la referencia.")
            else:
                st.warning("No se encontraron servicios con ese número exacto.")

# ==========================================
# 2. VISTA PRESTADOR DE SERVICIOS
# ==========================================
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
            p_row = prestador_valido.iloc[0]
            saldo_wallet = float(p_row.get("Saldo_Wallet", 0.0) if pd.notna(p_row.get("Saldo_Wallet", 0.0)) else 0.0)
            
            st.markdown("### 💼 Billetera Virtual y Membresías")
            col_st1, col_st2, col_st3 = st.columns(3)
            col_st1.metric("Afiliacion ($5)", str(p_row.get("Pago_Afiliacion_5USD", "Pendiente")))
            col_st2.metric("Mensualidad ($5)", str(p_row.get("Pago_Mensualidad_5USD", "Pendiente")))
            col_st3.metric("💳 Saldo Billetera (Wallet)", f"${saldo_wallet:,.2f} USD")
            st.markdown("---")

            servicios_cond = df[df["Conductor"].astype(str).str.strip().str.lower() == nom_prestador.strip().lower()]
            tab_viajes, tab_finanzas, tab_liquidaciones = st.tabs(["Mis Viajes Activos", "Mis Ganancias", "Liquidaciones"])
            
            with tab_viajes:
                if len(servicios_cond) > 0:
                    for idx, serv in servicios_cond.iterrows():
                        monto_usd = float(serv['Monto_USD'])
                        ganancia_usd = monto_usd * (st.session_state.porcentaje_prestador / 100)
                        with st.expander(f"Servicio #{serv['ID_Servicio']} - [{serv['Estado']}]"):
                            st.write(f"Cliente: {serv['Cliente']} | Origen: {serv['Origen']} ➡️ Destino: {serv['Destino']}")
                            st.write(f"Tu Ganancia: **${ganancia_usd:.2f} USD**")
                            
                            estados_posibles = ["Asignado", "En Camino", "Llegue al Sitio", "En Trayecto", "Completado", "Cancelado"]
                            idx_est = estados_posibles.index(serv['Estado']) if serv['Estado'] in estados_posibles else 0
                            nuevo_estado = st.selectbox("Actualizar Estatus:", estados_posibles, index=idx_est, key=f"st_{serv['ID_Servicio']}")
                            
                            if st.button(f"Guardar Estatus #{serv['ID_Servicio']}", key=f"btn_st_{serv['ID_Servicio']}"):
                                filtro_c = df["ID_Servicio"] == int(serv['ID_Servicio'])
                                df.loc[filtro_c, "Estado"] = str(nuevo_estado)
                                guardar_datos(df)
                                st.success("Estatus actualizado")
                                st.rerun()
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
                    st.dataframe(liq_prestador, use_container_width=True)
                else:
                    st.info("Sin liquidaciones registradas.")

# ==========================================
# 3. VISTA ALIADO COMERCIAL
# ==========================================
elif "Portal de Aliado Comercial" in rol:
    st.markdown("<h2>🏬 Portal de Aliados Comerciales</h2>", unsafe_allow_html=True)
    st.markdown("---")
    nombre_aliado = st.text_input("Nombre de su Comercio Aliado:")
    if nombre_aliado:
        st.success(f"Bienvenido, comercio: **{nombre_aliado}**")
        with st.form("form_aliado"):
            c_al1, c_al2 = st.columns(2)
            with c_al1:
                cliente_final = st.text_input("Cliente Final *")
                telefono_final = st.text_input("Teléfono Cliente Final *")
                monto_pedido = st.number_input("Monto Delivery USD ($)", min_value=1.0, value=3.0)
            with c_al2:
                origen_comercio = st.text_input("Dirección de su Local *")
                destino_cliente = st.text_input("Dirección de Entrega *")
                metodo_pago_comercio = st.selectbox("Método de Pago", ["Pago Móvil", "Efectivo", "Zelle"])
            
            if st.form_submit_button("🚀 Solicitar Delivery"):
                if cliente_final and telefono_final and destino_cliente:
                    nuevo_id = 1001 if len(df) == 0 else int(df["ID_Servicio"].max()) + 1
                    nuevo_registro_aliado = {
                        "ID_Servicio": nuevo_id, "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Tipo_Servicio": "Delivery", "Cliente": str(cliente_final), "Telefono_Cliente": str(telefono_final).strip(),
                        "Origen": str(origen_comercio) if origen_comercio else f"Local de {nombre_aliado}",
                        "Destino": str(destino_cliente), "Conductor": "Por Asignar", "Monto_USD": float(monto_pedido),
                        "Metodo_Pago": str(metodo_pago_comercio), "Ref_Pago_Movil": "", "Pago_Verificado": "Sí",
                        "Estado": "Solicitado", "Liquidado": "No", "Valoracion": "Sin valorar", "Notas": "",
                        "Aliado_Comercial": str(nombre_aliado)
                    }
                    df = pd.concat([df, pd.DataFrame([nuevo_registro_aliado])], ignore_index=True)
                    guardar_datos(df)
                    reproducir_alerta_sonora()
                    st.toast(f"¡Delivery #{nuevo_id} solicitado con éxito!", icon="🚀")
                    st.success(f"¡Delivery solicitado! ID #{nuevo_id}")
                else:
                    st.error("Completa los campos obligatorios.")

# ==========================================
# 4. VISTA ADMINISTRADOR (CON TODAS LAS FUNCIONES RESTAURADAS)
# ==========================================
elif "Panel Gerencial" in rol:
    st.markdown("<h2>Panel Gerencial LLEGANDO</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    clave_ingresada = st.text_input("Contraseña de Administrador (admin123):", type="password")
    if clave_ingresada == CLAVE_ADMIN_SECRETA:
        st.success("Acceso Gerencial Autorizado")
        st.markdown("---")
        
        with st.expander("⚙️ Configuración General del Sistema"):
            st.session_state.tasa_bcv = st.number_input("Tasa BCV", value=float(st.session_state.tasa_bcv))
            st.session_state.porcentaje_prestador = st.number_input("% Ganancia Prestador", value=float(st.session_state.porcentaje_prestador))

        opcion_admin = st.selectbox("Seleccione Sección Gerencial:", [
            "Asignar y Confirmar Pagos", 
            "Autorizar Prestadores", 
            "Liquidaciones a Prestadores", 
            "Control General de Servicios", 
            "Finanzas", 
            "Descargar Excel"
        ])
        st.markdown("---")
        
        if opcion_admin == "Asignar y Confirmar Pagos":
            st.subheader("Gestión de Solicitudes, Pagos y Asignación de Conductores")
            if len(df) > 0:
                for idx_a, row_a in df.iterrows():
                    with st.expander(f"ID #{row_a['ID_Servicio']} - {row_a['Cliente']} [{row_a['Estado']}]"):
                        st.write(f"Tipo: {row_a['Tipo_Servicio']} | Origen: {row_a['Origen']} ➡️ Destino: {row_a['Destino']}")
                        st.write(f"Método de Pago: {row_a['Metodo_Pago']} | Ref: `{row_a['Ref_Pago_Movil']}` | Verificado: **{row_a['Pago_Verificado']}**")
                        
                        lista_conductores = ["Por Asignar"] + list(df_prestadores[df_prestadores["Estatus"] == "Autorizado"]["Nombre"])
                        cond_actual = row_a['Conductor'] if row_a['Conductor'] in lista_conductores else "Por Asignar"
                        idx_c = lista_conductores.index(cond_actual)
                        
                        nuevo_cond = st.selectbox("Asignar Conductor:", lista_conductores, index=idx_c, key=f"cond_{row_a['ID_Servicio']}")
                        nuevo_verif = st.selectbox("Pago Verificado:", ["No", "Sí"], index=1 if row_a['Pago_Verificado'] in ["Sí", "si", "Si"] else 0, key=f"verif_{row_a['ID_Servicio']}")
                        estados_lista = ["Solicitado", "Asignado", "En Camino", "Llegue al Sitio", "En Trayecto", "Completado", "Cancelado"]
                        idx_e = estados_lista.index(row_a['Estado']) if row_a['Estado'] in estados_lista else 0
                        nuevo_est = st.selectbox("Estado del Servicio:", estados_lista, index=idx_e, key=f"est_admin_{row_a['ID_Servicio']}")
                        
                        if st.button(f"Guardar Cambios Servicio #{row_a['ID_Servicio']}", key=f"save_admin_{row_a['ID_Servicio']}"):
                            filtro_adm = df["ID_Servicio"] == int(row_a['ID_Servicio'])
                            df.loc[filtro_adm, "Conductor"] = str(nuevo_cond)
                            df.loc[filtro_adm, "Pago_Verificado"] = str(nuevo_verif)
                            df.loc[filtro_adm, "Estado"] = str(nuevo_est)
                            guardar_datos(df)
                            st.success("¡Actualizado correctamente!")
                            st.rerun()
            else:
                st.info("No hay servicios registrados.")

        elif opcion_admin == "Autorizar Prestadores":
            st.subheader("Registro y Autorización de Nuevos Conductores / Prestadores")
            with st.form("nuevo_p"):
                nombre_p = st.text_input("Nombre Completo *")
                telefono_p = st.text_input("Teléfono *")
                vehiculo_p = st.selectbox("Tipo de Vehículo *", ["Moto", "Carro / Taxi", "Flete", "Delivery"])
                placa_p = st.text_input("Placa *")
                pago_afi = st.selectbox("Pago Afiliación ($5)", ["Pendiente", "Pagado"])
                pago_men = st.selectbox("Pago Mensualidad ($5)", ["Pendiente", "Pagado"])
                saldo_ini = st.number_input("Saldo Inicial Billetera Wallet ($)", value=0.0)
                
                if st.form_submit_button("Autorizar Prestador"):
                    if nombre_p and telefono_p and placa_p:
                        nuevo_p = pd.DataFrame([{
                            "Nombre": str(nombre_p), "Telefono": str(telefono_p), "Vehiculo": str(vehiculo_p),
                            "Marca": "", "Modelo": "", "Placa": str(placa_p), "Foto_Path": "",
                            "Pago_Afiliacion_5USD": str(pago_afi), "Pago_Mensualidad_5USD": str(pago_men), 
                            "Estatus": "Autorizado", "Saldo_Wallet": str(saldo_ini)
                        }])
                        df_prestadores = pd.concat([df_prestadores, nuevo_p], ignore_index=True)
                        guardar_prestadores(df_prestadores)
                        st.success("¡Prestador autorizado con éxito!")
                        st.rerun()
                    else:
                        st.error("Completa los campos obligatorios.")
            
            if len(df_prestadores) > 0:
                st.markdown("#### Prestadores Autorizados Actuales:")
                st.dataframe(df_prestadores, use_container_width=True)

        elif opcion_admin == "Liquidaciones a Prestadores":
            st.subheader("Generar Liquidación a Prestador")
            lista_cond_aut = list(df_prestadores[df_prestadores["Estatus"] == "Autorizado"]["Nombre"])
            if len(lista_cond_aut) > 0:
                with st.form("form_liq"):
                    prestador_sel = st.selectbox("Seleccionar Prestador", lista_cond_aut)
                    monto_liq = st.number_input("Monto a Liquidar en USD ($)", min_value=1.0, value=10.0)
                    ref_pago_liq = st.text_input("Referencia de Pago Bancario realizado al Prestador *")
                    
                    if st.form_submit_button("Registrar y Enviar Liquidación"):
                        if ref_pago_liq:
                            nuevo_id_l = 1 if len(df_liquidaciones) == 0 else int(len(df_liquidaciones)) + 1
                            nueva_l = pd.DataFrame([{
                                "ID_Liquidacion": str(nuevo_id_l),
                                "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                "Prestador": str(prestador_sel),
                                "Monto_USD": float(monto_liq),
                                "Referencia_Pago": str(ref_pago_liq),
                                "Estatus_Confirmacion": "Pendiente de Confirmar"
                            }])
                            df_liquidaciones = pd.concat([df_liquidaciones, nueva_l], ignore_index=True)
                            guardar_liquidaciones(df_liquidaciones)
                            st.success("¡Liquidación registrada con éxito!")
                            st.rerun()
                        else:
                            st.error("Ingresa la referencia de pago.")
                
                st.markdown("#### Historial de Liquidaciones:")
                st.dataframe(df_liquidaciones, use_container_width=True)
            else:
                st.warning("No hay prestadores autorizados disponibles.")

        elif opcion_admin == "Control General de Servicios":
            st.subheader("Base de Datos Completa de Servicios")
            st.dataframe(df, use_container_width=True)

        elif opcion_admin == "Finanzas":
            st.subheader("Resumen Financiero")
            total_recaudado = pd.to_numeric(df[df["Pago_Verificado"].str.lower().isin(["sí", "si"])]["Monto_USD"], errors='coerce').sum()
            st.metric("Total Recaudado (USD Verificado)", f"${total_recaudado:,.2f}")
            st.dataframe(df[["ID_Servicio", "Fecha_Hora", "Cliente", "Monto_USD", "Metodo_Pago", "Pago_Verificado"]], use_container_width=True)

        elif opcion_admin == "Descargar Excel":
            st.subheader("Exportar Datos")
            st.download_button("📥 Descargar Base de Datos Completa (Excel)", data=open(DB_FILE, "rb").read(), file_name="llegando_db.xlsx")

    else:
        if clave_ingresada:
            st.error("Contraseña incorrecta.")

import streamlit as st
import pandas as pd
import os
import requests
from datetime import datetime

st.set_page_config(page_title="App de Movilidad y Servicios", page_icon="🚖", layout="wide")

DB_FILE = "base_datos_servicios.xlsx"

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

df = cargar_datos()

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
                    st.markdown(f"**👨‍✈️ Conductor:** **{serv['Conductor']}**")
                    st.markdown(f"**💵 Tarifa:** **${monto_usd:,.2f} USD** (Bs. {monto_ves:,.2f})")
                
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
                    
                    val_actual = "" if pd.isna(serv['Ref_Pago_Movil']) or str(serv['Ref_Pago_Movil']) == "nan" else str(serv['Ref_Pago_Movil'])
                    ref_ingresada = st.text_input("📝 Ingresa los dígitos de tu referencia:", value=val_actual)
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
                for idx, serv in servicios_cond.iterrows():
                    monto_usd = float(serv['Monto_USD'])
                    monto_ves = monto_usd * st.session_state.tasa_bcv
                    ganancia_usd = monto_usd * (st.session_state.porcentaje_prestador / 100)
                    ganancia_ves = ganancia_usd * st.session_state.tasa_bcv
                    
                    with st.expander(f"Servicio #{serv['ID_Servicio']} - {serv['Origen']} ➡️ {serv['Destino']} [{serv['Estado']}]"):
                        st.write(f"**Cliente:** {serv['Cliente']} | **Teléfono:** {serv['Telefono_Cliente']}")
                        st.write(f"**Monto Cobrar:** ${monto_usd:.2f} USD (Bs. {monto_ves:,.2f})")
                        st.write(f"**Tu Ganancia ({st.session_state.porcentaje_prestador}%):** **${ganancia_usd:.2f} USD** (Bs. {ganancia_ves:,.2f})")
                        st.write(f"**Ref. Pago Móvil:** {serv['Ref_Pago_Movil']}")
                        
                        estados_posibles = ["Asignado", "En Camino", "Llegué al Sitio", "En Trayecto", "Completado", "Cancelado"]
                        idx_est = estados_posibles.index(serv['Estado']) if serv['Estado'] in estados_posibles else 0
                        nuevo_estado = st.selectbox("Actualizar Estatus:", estados_posibles, index=idx_est, key=f"st_{serv['ID_Servicio']}")
                        
                        if st.button(f"Guardar Estatus #{serv['ID_Servicio']}", key=f"btn_{serv['ID_Servicio']}"):
                            df.at[idx, "Estado"] = nuevo_estado
                            guardar_datos(df)
                            st.success("¡Estatus actualizado!")
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

# ==========================================
# 3. VISTA ADMINISTRADOR
# ==========================================
elif rol == "⚙️ Administrador":
    st.title("⚙️ Panel de Control Gerencial")
    
    with st.expander("⚙️ Configuración General del Sistema", expanded=True):
        col_adm1, col_adm2 = st.columns(2)
        with col_adm1:
            st.session_state.tasa_bcv = st.number_input("Tasa BCV (Bs. / USD)", value=float(st.session_state.tasa_bcv), step=0.10)
        with col_adm2:
            st.session_state.porcentaje_prestador = st.number_input("% Ganancia Prestador", value=float(st.session_state.porcentaje_prestador), step=5.0)
            
    st.markdown("---")
    
    opcion_admin = st.selectbox("Seleccione una sección de administración:", ["📥 Solicitudes Entrantes / Asignación", "📋 Control General", "📊 Finanzas", "📁 Excel"])
    
    if opcion_admin == "📥 Solicitudes Entrantes / Asignación":
        st.subheader("📥 Solicitudes por Asignar Conductor y Tarifa")
        pendientes = df[df["Estado"].isin(["Solicitado", "Pendiente"])]
        
        if len(pendientes) > 0:
            for idx, serv in pendientes.iterrows():
                with st.expander(f"Solicitud #{serv['ID_Servicio']} - {serv['Cliente']} ({serv['Tipo_Servicio']})"):
                    st.write(f"**Origen:** {serv['Origen']} ➡️ **Destino:** {serv['Destino']}")
                    st.write(f"**Teléfono:** {serv['Telefono_Cliente']} | **Pago:** {serv['Metodo_Pago']}")
                    
                    c_asig1, c_asig2 = st.columns(2)
                    with c_asig1:
                        cond_asignado = st.text_input(f"Asignar Conductor", key=f"cond_{serv['ID_Servicio']}")
                    with c_asig2:
                        tarifa_usd = st.number_input(f"Asignar Tarifa USD", min_value=0.0, step=0.5, key=f"tar_{serv['ID_Servicio']}")
                        
                    if st.button(f"Confirmar Asignación #{serv['ID_Servicio']}", key=f"btn_asig_{serv['ID_Servicio']}"):
                        df.at[idx, "Conductor"] = cond_asignado if cond_asignado else "Por Asignar"
                        df.at[idx, "Monto_USD"] = tarifa_usd
                        df.at[idx, "Estado"] = "Asignado"
                        guardar_datos(df)
                        st.success(f"¡Servicio #{serv['ID_Servicio']} asignado con éxito!")
                        st.rerun()
        else:
            st.info("No hay solicitudes pendientes por asignar en este momento.")

    elif opcion_admin == "📋 Control General":
        st.subheader("📋 Todos los Servicios Registrados")
        st.dataframe(df.astype(str), use_container_width=True)

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
            st.info("Aún no hay servicios completados para calcular finanzas.")

    elif opcion_admin == "📁 Excel":
        st.subheader("📁 Descargar Base de Datos")
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as file:
                st.download_button("⬇️ Descargar Archivo Excel", file, file_name="base_datos_servicios.xlsx")

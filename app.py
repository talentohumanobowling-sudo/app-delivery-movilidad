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

if "Portal de Cliente" in rol:
    st.markdown("<h2>LLEGANDO - Solicitud y Estatus</h2>", unsafe_allow_html=True)
    st.markdown("---")
    tab_solicitar, tab_consultar = st.tabs(["Solicitar Nuevo Servicio", "Consultar Mis Viajes"])
    
    with tab_solicitar:
        with st.form("form_cliente"):
            col1, col2 = st.columns(2)
            with col1:
                nombre = st.text_input("Tu Nombre Completo *")
                telefono = st.text_input("Tu Número de Teléfono *")
                tipo_serv = st.selectbox("Tipo de Servicio *", ["Delivery", "Mototaxi", "Taxi / Carro", "Flete"])
            with col2:
                origen = st.text_input("Dirección de Origen *")
                destino = st.text_input("Dirección de Destino *")
                metodo_pago = st.selectbox("Método de Pago Preferido", ["Pago Móvil", "Efectivo", "Zelle"])
            
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
                    st.success(f"¡Solicitud enviada! Tu número de ID es #{nuevo_id}.")
                else:
                    st.error("Completa los campos obligatorios (*).")
                    
    with tab_consultar:
        st.subheader("Mis Servicios Activos e Historial")
        telefono_busqueda = st.text_input("Ingresa tu número de teléfono registrado:")
        if telefono_busqueda:
            resultados = df[df["Telefono_Cliente"].astype(str).str.strip() == telefono_busqueda.strip()]
            if len(resultados) > 0:
                resultados = resultados.sort_values(by="ID_Servicio", ascending=False)
                for idx_res, serv in resultados.iterrows():
                    monto_usd = float(serv['Monto_USD'])
                    monto_ves = monto_usd * st.session_state.tasa_bcv
                    estatus_s = str(serv['Estado']).strip()
                    with st.expander(f"Servicio #{serv['ID_Servicio']} - {serv['Tipo_Servicio']} [{estatus_s}]", expanded=True):
                        st.markdown(f"**Origen:** {serv['Origen']} | **Destino:** {serv['Destino']}")
                        st.markdown(f"**Tarifa:** **${monto_usd:,.2f} USD** (Bs. {monto_ves:,.2f})")
            else:
                st.warning("No se encontraron servicios con ese número.")

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
            st.metric("💳 Saldo Billetera", f"${saldo_wallet:,.2f} USD")
            servicios_cond = df[df["Conductor"].astype(str).str.strip().str.lower() == nom_prestador.strip().lower()]
            tab_viajes, tab_finanzas = st.tabs(["Mis Viajes Activos", "Mis Ganancias"])
            with tab_viajes:
                st.subheader("Servicios Asignados en Curso")
                viajes_activos_cond = servicios_cond[servicios_cond["Estado"].isin(["Asignado", "En Camino", "Llegue al Sitio", "En Trayecto"])]
                if len(viajes_activos_cond) == 0:
                    st.info("No tienes servicios activos.")
                else:
                    st.dataframe(viajes_activos_cond[["ID_Servicio", "Tipo_Servicio", "Origen", "Destino", "Monto_USD", "Estado"]])
            with tab_finanzas:
                st.subheader("Resumen de Ganancias")
                servicios_completados = servicios_cond[servicios_cond["Estado"] == "Completado"]
                total_ganado = servicios_completados["Monto_USD"].astype(float).sum() * (st.session_state.porcentaje_prestador / 100.0)
                st.metric("Total Acumulado", f"${total_ganado:,.2f} USD")

elif "Portal de Aliado Comercial" in rol:
    st.markdown("<h2>Portal de Aliados Comerciales</h2>", unsafe_allow_html=True)
    st.markdown("---")
    nombre_aliado = st.text_input("Ingresa el Nombre de tu Comercio / Aliado:")
    if nombre_aliado:
        st.success(f"Bienvenido, comercio: {nombre_aliado}")
        with st.form("form_pedido_aliado"):
            st.subheader("Solicitar Servicio de Delivery para tu Local")
            c_cliente = st.text_input("Nombre del Cliente Final")
            c_tlf = st.text_input("Teléfono del Cliente")
            c_destino = st.text_input("Dirección de Entrega del Pedido")
            c_monto = st.number_input("Monto del Pedido ($ USD)", min_value=1.0, value=10.0)
            if st.form_submit_button("Crear Delivery Express"):
                nuevo_id = 1001 if len(df) == 0 else int(pd.to_numeric(df["ID_Servicio"], errors='coerce').max()) + 1
                nuevo_reg = {
                    "ID_Servicio": nuevo_id, "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Tipo_Servicio": "Delivery", "Cliente": c_cliente, "Telefono_Cliente": str(c_tlf).strip(),
                    "Origen": nombre_aliado, "Destino": c_destino, "Conductor": "Por Asignar", "Monto_USD": c_monto,
                    "Metodo_Pago": "Efectivo / POS", "Ref_Pago_Movil": "", "Pago_Verificado": "Sí",
                    "Estado": "Solicitado", "Liquidado": "No", "Valoracion": "Sin valorar", "Notas": f"Aliado: {nombre_aliado}", 
                    "Aliado_Comercial": nombre_aliado, "Capture_Pago": ""
                }
                df = pd.concat([df, pd.DataFrame([nuevo_reg])], ignore_index=True)
                guardar_datos(df)
                st.success(f"¡Delivery creado con éxito! ID #{nuevo_id}")

elif "Panel Gerencial (Admin)" in rol:
    st.markdown("<h2>Panel Gerencial y de Control Admin</h2>", unsafe_allow_html=True)
    st.markdown("---")
    clave = st.text_input("Contraseña de Administrador:", type="password")
    if clave == CLAVE_ADMIN_SECRETA:
        st.success("Acceso Gerencial Concedido")
        tab_serv, tab_prest, tab_conf = st.tabs(["Control de Servicios", "Prestadores Autorizados", "Configuración"])
        with tab_serv:
            st.subheader("Gestión de Todos los Servicios")
            st.dataframe(df)
        with tab_prest:
            st.subheader("Control de Prestadores")
            st.dataframe(df_prestadores)
        with tab_conf:
            st.subheader("Ajustes del Sistema")
            nueva_tasa = st.number_input("Actualizar Tasa BCV Manual (Bs.)", value=float(st.session_state.tasa_bcv))
            if st.button("Guardar Tasa"):
                st.session_state.tasa_bcv = nueva_tasa
                st.success(f"Tasa actualizada a Bs. {nueva_tasa:,.2f}")
    elif clave:
        st.error("Contraseña incorrecta.")

import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(page_title="Plataforma de Movilidad y Delivery", page_icon="🚖", layout="wide")

DB_FILE = "base_datos_servicios.xlsx"

# Datos de Pago Móvil de la Empresa
DATOS_PAGO_MOVIL = {
    "Banco": "BDV (0102)",
    "Telefono": "0412-3349463",
    "Cedula_RIF": "V-13609424",
    "Titular": "Servicios de Movilidad"
}

def cargar_datos():
    if os.path.exists(DB_FILE):
        return pd.read_excel(DB_FILE)
    else:
        df_inicial = pd.DataFrame(columns=[
            "ID_Servicio", "Fecha_Hora", "Tipo_Servicio", "Cliente", 
            "Telefono_Cliente", "Origen", "Destino", "Conductor", 
            "Monto_USD", "Metodo_Pago", "Ref_Pago_Movil", "Estado", "Notas"
        ])
        df_inicial.to_excel(DB_FILE, index=False)
        return df_inicial

def guardar_datos(df):
    df.to_excel(DB_FILE, index=False)

df = cargar_datos()

# MENU SUPERIOR / SELECCIÓN DE ROL
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3097/3097180.png", width=100)
st.sidebar.title("Bienvenido")
rol = st.sidebar.radio("¿Cómo deseas ingresar?", [
    "👤 Soy Cliente", 
    "🚗 Soy Conductor / Repartidor", 
    "⚙️ Administrador"
])

# ==========================================
# 1. VISTA EXCLUSIVA PARA EL CLIENTE
# ==========================================
if rol == "👤 Soy Cliente":
    st.markdown("<h2 style='text-align: center; color: #1E88E5;'>📱 Portal de Atención al Cliente</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Consulta el estado de tu servicio, asignación de unidad y realiza tu pago.</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    if len(df) == 0:
        st.info("No hay servicios activos registrados.")
    else:
        id_busqueda = st.number_input("👉 Ingresa tu Número de Servicio (ID):", min_value=1001, step=1, value=1001)
        resultado = df[df["ID_Servicio"] == id_busqueda]
        
        if len(resultado) > 0:
            serv = resultado.iloc[0]
            
            # Tarjeta resumen del cliente
            st.success(f"### 📍 Servicio #{serv['ID_Servicio']}")
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f"**👤 Cliente:** {serv['Cliente']}")
                st.markdown(f"**🚘 Servicio:** {serv['Tipo_Servicio']}")
                st.markdown(f"**📍 Salida:** {serv['Origen']}")
                st.markdown(f"**🏁 Llegada:** {serv['Destino']}")
            with col2:
                st.markdown(f"**🚦 Estatus:** `{serv['Estado']}`")
                st.markdown(f"**👨‍✈️ Conductor:** **{serv['Conductor']}**")
                st.markdown(f"**💵 Total a Pagar:** **${serv['Monto_USD']:,.2f} USD**")
            
            st.markdown("---")
            
            # Sección de Pago Móvil para el Cliente
            if serv['Metodo_Pago'] == "Pago Móvil":
                st.info("📲 **Datos para realizar tu Pago Móvil:**")
                col_p1, col_p2 = st.columns(2)
                with col_p1:
                    st.write(f"🏦 **Banco:** {DATOS_PAGO_MOVIL['Banco']}")
                    st.write(f"📱 **Teléfono:** {DATOS_PAGO_MOVIL['Telefono']}")
                with col_p2:
                    st.write(f"🪪 **Cédula / RIF:** {DATOS_PAGO_MOVIL['Cedula_RIF']}")
                    st.write(f"👤 **Titular:** {DATOS_PAGO_MOVIL['Titular']}")
                
                ref_ingresada = st.text_input("📝 Ingresa los dígitos de tu referencia bancaria:", value=str(serv['Ref_Pago_Movil']) if pd.notna(serv['Ref_Pago_Movil']) else "")
                if st.button("Confirmar y Reportar Pago"):
                    idx = df[df["ID_Servicio"] == id_busqueda].index
                    df.loc[idx, "Ref_Pago_Movil"] = ref_ingresada
                    guardar_datos(df)
                    st.success("¡Referencia registrada correctamente! Gracias por tu pago.")
                    st.rerun()
        else:
            st.warning("No encontramos ningún servicio registrado con ese número de ID.")

# ==========================================
# 2. VISTA EXCLUSIVA PARA CONDUCTORES
# ==========================================
elif rol == "🚗 Soy Conductor / Repartidor":
    st.title("🚗 Panel de Operaciones - Conductores")
    st.markdown("---")
    
    nom_cond = st.text_input("Ingresa tu Nombre / Apellido para filtrar tus servicios:")
    
    if nom_cond:
        servicios_cond = df[df["Conductor"].astype(str).str.contains(nom_cond, case=False, na=False)]
        
        if len(servicios_cond) > 0:
            st.subheader(f"Servicios Asignados a: {nom_cond}")
            for idx, serv in servicios_cond.iterrows():
                with st.expander(f"Servicio #{serv['ID_Servicio']} - {serv['Origen']} ➡️ {serv['Destino']} [{serv['Estado']}]"):
                    st.write(f"**Cliente:** {serv['Cliente']} ({serv['Telefono_Cliente']})")
                    st.write(f"**Monto Cobrar:** ${serv['Monto_USD']} USD - ({serv['Metodo_Pago']})")
                    st.write(f"**Notas:** {serv['Notas']}")
                    
                    nuevo_estado = st.selectbox(f"Actualizar Estado (#{serv['ID_Servicio']})", ["Pendiente", "En Camino", "Completado", "Cancelado"], key=f"est_{serv['ID_Servicio']}")
                    if st.button(f"Guardar Estatus #{serv['ID_Servicio']}", key=f"btn_{serv['ID_Servicio']}"):
                        df.loc[idx, "Estado"] = nuevo_estado
                        guardar_datos(df)
                        st.success("¡Estado actualizado!")
                        st.rerun()
        else:
            st.info("No tienes servicios asignados bajo ese nombre actualmente.")

# ==========================================
# 3. VISTA EXCLUSIVA ADMINISTRADOR
# ==========================================
elif rol == "⚙️ Administrador":
    st.title("⚙️ Panel de Control Gerencial")
    
    opcion_admin = st.selectbox("Gestión Admin", ["➕ Registrar Solicitud", "📋 Control General", "📊 Finanzas", "📁 Archivo Excel"])
    
    if opcion_admin == "➕ Registrar Solicitud":
        st.subheader("📝 Registrar Nuevo Viaje / Servicio")
        with st.form("form_admin"):
            c1, c2 = st.columns(2)
            with c1:
                tipo = st.selectbox("Tipo de Servicio", ["Delivery", "Mototaxi", "Taxi / Carro", "Flete"])
                cliente = st.text_input("Cliente *")
                telefono = st.text_input("Teléfono Cliente")
                origen = st.text_input("Origen *")
                destino = st.text_input("Destino *")
            with c2:
                conductor = st.text_input("Conductor Asignado")
                monto = st.number_input("Tarifa USD *", min_value=0.0, step=0.5)
                metodo = st.selectbox("Método de Pago", ["Pago Móvil", "Efectivo", "Zelle", "Transferencia"])
                ref_pm = st.text_input("Ref. Pago Móvil")
                estado = st.selectbox("Estado", ["Pendiente", "En Camino", "Completado", "Cancelado"])
                notas = st.text_area("Observaciones")
                
            if st.form_submit_button("Guardar en Sistema"):
                if cliente and origen and destino:
                    nuevo_id = 1001 if len(df) == 0 else int(df["ID_Servicio"].max()) + 1
                    nuevo_reg = {
                        "ID_Servicio": nuevo_id, "Fecha_Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Tipo_Servicio": tipo, "Cliente": cliente, "Telefono_Cliente": telefono,
                        "Origen": origen, "Destino": destino, "Conductor": conductor if conductor else "Por Asignar",
                        "Monto_USD": monto, "Metodo_Pago": metodo, "Ref_Pago_Movil": ref_pm, "Estado": estado, "Notas": notas
                    }
                    df = pd.concat([df, pd.DataFrame([nuevo_reg])], ignore_index=True)
                    guardar_datos(df)
                    st.success(f"¡Servicio registrado con el ID **#{nuevo_id}**!")

    elif opcion_admin == "📋 Control General":
        st.subheader("📋 Todos los Servicios")
        st.dataframe(df, use_container_width=True)

    elif opcion_admin == "📊 Finanzas":
        st.subheader("📊 Métricas de Ingresos")
        if len(df) > 0:
            st.metric("Total Recaudado", f"${df['Monto_USD'].sum():,.2f} USD")
            st.bar_chart(df.groupby("Tipo_Servicio")["Monto_USD"].sum())

    elif opcion_admin == "📁 Archivo Excel":
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as file:
                st.download_button("⬇️ Descargar Base de Datos Excel", file, file_name="base_datos_servicios.xlsx")

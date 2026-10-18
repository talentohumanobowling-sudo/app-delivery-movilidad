import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(page_title="App de Servicios: Delivery, Mototaxi y Taxi", page_icon="🚖", layout="wide")

DB_FILE = "base_datos_servicios.xlsx"

# Función para cargar datos desde Excel
def cargar_datos():
    if os.path.exists(DB_FILE):
        return pd.read_excel(DB_FILE)
    else:
        df_inicial = pd.DataFrame(columns=[
            "ID_Servicio", "Fecha_Hora", "Tipo_Servicio", "Cliente", 
            "Telefono_Cliente", "Origen", "Destino", "Conductor", 
            "Monto_USD", "Metodo_Pago", "Estado", "Notas"
        ])
        df_inicial.to_excel(DB_FILE, index=False)
        return df_inicial

# Función para guardar datos en Excel
def guardar_datos(df):
    df.to_excel(DB_FILE, index=False)

df = cargar_datos()

# Encabezado
st.title("🚖 Plataforma de Gestión de Servicios")
st.caption("Sistema Multiusuario Sin Límites | Conexión Automática con Excel")
st.markdown("---")

# Menú lateral
st.sidebar.title("Navegación")
opcion = st.sidebar.radio("Ir a:", [
    "➕ Registrar Solicitud", 
    "📋 Control de Viajes y Envíos", 
    "📊 Panel Financiero", 
    "📥 Exportar / Cargar Excel"
])

# 1. REGISTRAR SOLICITUD
if opcion == "➕ Registrar Solicitud":
    st.subheader("📝 Solicitar / Asignar Nuevo Servicio")
    
    with st.form("form_registro", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            tipo = st.selectbox("Tipo de Servicio *", ["Delivery (Encomiendas)", "Mototaxi", "Taxi / Carro", "Flete / Carga"])
            cliente = st.text_input("Nombre del Cliente *")
            telefono = st.text_input("Teléfono del Cliente")
            origen = st.text_input("Punto de Origen / Recogida *")
            destino = st.text_input("Punto de Destino / Entrega *")
            
        with col2:
            conductor = st.text_input("Conductor / Repartidor Asignado")
            monto = st.number_input("Monto / Tarifa ($ USD) *", min_value=0.0, step=0.5, format="%.2f")
            metodo_pago = st.selectbox("Método de Pago", ["Efectivo", "Pago Móvil", "Zelle", "Transferencia", "Otro"])
            estado = st.selectbox("Estado Inicial", ["Pendiente", "En Camino", "Completado", "Cancelado"])
            notas = st.text_area("Detalles / Observaciones")
            
        btn_guardar = st.form_submit_button("🚀 Registrar Servicio")
        
        if btn_guardar:
            if not cliente or not origen or not destino:
                st.error("Por favor completa los campos obligatorios (*).")
            else:
                nuevo_id = 1001 if len(df) == 0 else int(df["ID_Servicio"].max()) + 1
                fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                nuevo_registro = {
                    "ID_Servicio": nuevo_id,
                    "Fecha_Hora": fecha_actual,
                    "Tipo_Servicio": tipo,
                    "Cliente": cliente,
                    "Telefono_Cliente": telefono,
                    "Origen": origen,
                    "Destino": destino,
                    "Conductor": conductor if conductor else "Por Asignar",
                    "Monto_USD": monto,
                    "Metodo_Pago": metodo_pago,
                    "Estado": estado,
                    "Notas": notas
                }
                
                df = pd.concat([df, pd.DataFrame([nuevo_registro])], ignore_index=True)
                guardar_datos(df)
                st.success(f"¡Servicio #{nuevo_id} registrado y guardado exitosamente en el Excel!")

# 2. CONTROL DE VIAJES Y ENVÍOS
elif opcion == "📋 Control de Viajes y Envíos":
    st.subheader("📋 Estado Actual de la Operación")
    
    if len(df) == 0:
        st.info("Aún no hay registros en el sistema.")
    else:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filtro_tipo = st.multiselect("Filtrar por Servicio", df["Tipo_Servicio"].unique(), default=df["Tipo_Servicio"].unique())
        with col_f2:
            filtro_estado = st.multiselect("Filtrar por Estado", df["Estado"].unique(), default=df["Estado"].unique())
            
        df_filtrado = df[(df["Tipo_Servicio"].isin(filtro_tipo)) & (df["Estado"].isin(filtro_estado))]
        st.dataframe(df_filtrado, use_container_width=True)
        
        st.markdown("---")
        st.subheader("🔄 Actualizar Estado / Conductor de un Servicio")
        col_u1, col_u2, col_u3 = st.columns(3)
        
        with col_u1:
            id_act = st.selectbox("Seleccionar ID de Servicio", df["ID_Servicio"].unique())
        with col_u2:
            nuevo_est = st.selectbox("Nuevo Estado", ["Pendiente", "En Camino", "Completado", "Cancelado"])
        with col_u3:
            nuevo_cond = st.text_input("Asignar / Cambiar Conductor")
            
        if st.button("Actualizar Registro en Excel"):
            idx = df[df["ID_Servicio"] == id_act].index
            if len(idx) > 0:
                df.loc[idx, "Estado"] = nuevo_est
                if nuevo_cond:
                    df.loc[idx, "Conductor"] = nuevo_cond
                guardar_datos(df)
                st.success(f"¡Servicio #{id_act} actualizado en Excel!")
                st.rerun()

# 3. PANEL FINANCIERO
elif opcion == "📊 Panel Financiero":
    st.subheader("📊 Indicadores de Rendimiento e Ingresos")
    
    if len(df) == 0:
        st.info("Sin datos acumulados.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Servicios Totales", len(df))
        c2.metric("Ingresos Totales", f"${df['Monto_USD'].sum():,.2f} USD")
        c3.metric("Completados", len(df[df["Estado"] == "Completado"]))
        c4.metric("En Curso / Pendientes", len(df[df["Estado"].isin(["Pendiente", "En Camino"])]))
        
        st.markdown("---")
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.markdown("##### Ingresos por Tipo de Servicio")
            st.bar_chart(df.groupby("Tipo_Servicio")["Monto_USD"].sum())
        with col_g2:
            st.markdown("##### Métodos de Pago")
            st.bar_chart(df["Metodo_Pago"].value_counts())

# 4. EXPORTAR / CARGAR EXCEL
elif opcion == "📥 Exportar / Cargar Excel":
    st.subheader("📁 Gestión de Archivo Excel")
    
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        st.markdown("### Descargar Base de Datos")
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as file:
                st.download_button(
                    label="⬇️ Descargar Excel (base_datos_servicios.xlsx)",
                    data=file,
                    file_name="base_datos_servicios.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    with col_e2:
        st.markdown("### Cargar o Reemplazar Excel")
        subir_excel = st.file_uploader("Subir archivo Excel existente", type=["xlsx", "xls"])
        if subir_excel is not None:
            df_nuevo = pd.read_excel(subir_excel)
            guardar_datos(df_nuevo)
            st.success("¡Base de datos actualizada correctamente!")

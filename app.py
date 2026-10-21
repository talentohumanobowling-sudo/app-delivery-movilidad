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

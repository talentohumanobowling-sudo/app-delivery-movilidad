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

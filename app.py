# JC Enhancement Programation
# Luis Eduardo Espejel Hernández

import streamlit as st
import pandas as pd
from datetime import timedelta, time
import io

st.title("KMX Asignación de JC en Ventanas")
st.write("Sube el archivo de Jit Calls para procesar las órdenes y asignar ventanas.")

# Cargar archivo de órdenes
file_ordenes = st.file_uploader("Sube archivo de ordenes.xlsx", type="xlsx")

if file_ordenes:
    ordenes_df = pd.read_excel(file_ordenes)
    ventanas_df = pd.read_excel("JC Delivery Schedule.xlsx", engine="openpyxl")

    # Limpiar nombres de columnas
    ordenes_df.columns = ordenes_df.columns.str.strip()
    ventanas_df.columns = ventanas_df.columns.str.strip()

    # Función para convertir hora a timedelta
    def convertir_hora(valor):
        if isinstance(valor, (float, int)):
            return timedelta(days=valor)
        elif isinstance(valor, str):
            try:
                h, m, s = map(int, valor.split(":"))
                return timedelta(hours=h, minutes=m, seconds=s)
            except:
                return pd.NaT
        elif isinstance(valor, time):
            return timedelta(hours=valor.hour, minutes=valor.minute, seconds=valor.second)
        else:
            return pd.NaT

    # Convertir hora y fecha
    ordenes_df["HoraExcel"] = ordenes_df["Require Time"].apply(convertir_hora)
    ordenes_df["Require Date"] = pd.to_datetime(ordenes_df["Require Date"], errors="coerce")
    ordenes_df["FechaHoraEntrega"] = ordenes_df["Require Date"] + ordenes_df["HoraExcel"]

    # Diccionario de días en español
    dias_es = {
        "Monday": "LUNES",
        "Tuesday": "MARTES",
        "Wednesday": "MIÉRCOLES",
        "Thursday": "JUEVES",
        "Friday": "VIERNES",
        "Saturday": "SÁBADO",
        "Sunday": "DOMINGO"
    }

    # Crear DíaSemana solo cuando hay fecha válida
    ordenes_df["DiaSemana"] = ordenes_df["FechaHoraEntrega"].dt.day_name().map(dias_es).fillna("SIN FECHA")

    # Función para convertir hora a decimal
    def hora_a_decimal(hora):
        if isinstance(hora, str):
            try:
                h, m, s = map(int, hora.split(":"))
                return h / 24 + m / 1440 + s / 86400
            except:
                return None
        elif isinstance(hora, time):
            return hora.hour / 24 + hora.minute / 1440 + hora.second / 86400
        else:
            return None

    ventanas_df["InicioDecimal"] = ventanas_df["Init Time"].apply(hora_a_decimal)
    ventanas_df["FinDecimal"] = ventanas_df["Fin Time"].apply(hora_a_decimal)

    dias_semana = {
        0: "LUNES", 1: "MARTES", 2: "MIÉRCOLES", 3: "JUEVES",
        4: "VIERNES", 5: "SÁBADO", 6: "DOMINGO"
    }

    ahora = pd.Timestamp.now()

    # Función para asignar ventana
    def asignar_ventana(fila):
        proveedor = str(fila["Vendor"]).strip().upper()
        dia_actual = str(fila["DiaSemana"]).strip().upper()
        hora = fila["FechaHoraEntrega"]

        if pd.isna(hora):
            return "Sin fecha", None

        # Si la fecha y hora ya pasó, marcar directamente como fuera de ventana
        if hora < ahora:
            return "Fuera de ventana", dia_actual

        hora_decimal = hora.hour / 24 + hora.minute / 1440 + hora.second / 86400

        # Ventanas del mismo día
        ventanas_dia = ventanas_df[(ventanas_df["Vendor"].astype(str).str.upper() == proveedor) &
                                   (ventanas_df["Day"].astype(str).str.upper() == dia_actual)]

        for _, v in ventanas_dia.iterrows():
            ini = v["InicioDecimal"]
            fin = v["FinDecimal"]
            if ini is None or fin is None:
                continue
            if ini < fin:  # Ventana normal
                if ini <= hora_decimal <= fin:
                    return v["Ventana"], dia_actual
            else:  # Ventana cruza de día
                if hora_decimal >= ini:
                    return v["Ventana"], dia_actual

        # Ventanas del día anterior
        dia_anterior = dias_semana[(fila["FechaHoraEntrega"].weekday() - 1) % 7]
        ventanas_previas = ventanas_df[(ventanas_df["Vendor"].astype(str).str.upper() == proveedor) &
                                       (ventanas_df["Day"].astype(str).str.upper() == dia_anterior)]

        for _, v in ventanas_previas.iterrows():
            ini = v["InicioDecimal"]
            fin = v["FinDecimal"]
            if ini is None or fin is None:
                continue
            if ini > fin and hora_decimal <= fin:
                return v["Ventana"], dia_anterior

        return "Fuera de ventana", None

    # Aplicar función
    ordenes_df[["VentanaAsignada", "DiaVentana"]] = ordenes_df.apply(lambda fila: pd.Series(asignar_ventana(fila)), axis=1)

    # Ordenar mostrando primero las órdenes fuera de ventana
    resultado = ordenes_df[[
        "JIT Call No", "Vendor", "Material", "JIT Call Quantity", "FechaHoraEntrega",
        "VentanaAsignada"
    ]].sort_values(by="VentanaAsignada", ascending=False)

    st.write("✅ Procesamiento completo.")
    st.write("Enviar cantidades en las siguientes ventanas urgentemente para evitar cortos.")

    # Cálculo resumen
    total_ordenes = len(resultado)
    ordenes_fuera = (resultado["VentanaAsignada"] == "Fuera de ventana").sum()
    eficiencia = 100 * (total_ordenes - ordenes_fuera) / total_ordenes if total_ordenes > 0 else 0

    st.write(f"**Total de órdenes:** {total_ordenes}")
    st.write(f"**❌ Órdenes fuera de ventana:** {ordenes_fuera}")
    st.write(f"**% Eficiencia:** {eficiencia:.2f}%")

    # Resaltar filas fuera de ventana
    def resaltar_fuera(val):
        color = "background-color: #ffcccc" if val == "Fuera de ventana" else ""
        return color

    st.dataframe(
        resultado.style.applymap(resaltar_fuera, subset=["VentanaAsignada"])
    )

    # Crear tabla pivote
    resultado['JIT Call Quantity'] = pd.to_numeric(resultado['JIT Call Quantity'], errors='coerce').fillna(0)

    ventanas = ['Fuera de ventana'] + sorted([v for v in resultado['VentanaAsignada'].unique() if v != 'Fuera de ventana'])

    tabla_pivote = pd.pivot_table(
        resultado,
        index='Material',
        columns='VentanaAsignada',
        values='JIT Call Quantity',
        aggfunc='sum',
        fill_value=0
    )

    # Reordenar columnas para que 'Fuera de ventana' quede primero
    tabla_pivote = tabla_pivote.reindex(columns=ventanas, fill_value=0)

    st.write("### Tabla pivote: Suma por número de parte y Ventana Asignada")
    st.dataframe(tabla_pivote)

    # Botón para descargar Excel con ambas hojas
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        resultado.to_excel(writer, index=False, sheet_name='Órdenes')
        tabla_pivote.to_excel(writer, sheet_name='Tabla Pivote')
    output.seek(0)

    st.download_button(
        label="Descargar resultado completo en Excel",
        data=output,
        file_name="Asignaciones_Ventanas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

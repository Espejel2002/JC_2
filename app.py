# JC Enhancement Programation
# Luis Eduardo Espejel Hernández

import streamlit as st
import pandas as pd
from datetime import timedelta, time
import io

st.title("Asignación de Jit Calls en Ventanas por proveedor")
st.write("Sube los archivos para procesar las órdenes y asignar ventanas.")

file_ordenes = st.file_uploader("Sube archivo de ordenes.xlsx", type="xlsx")

if file_ordenes:
    ordenes_df = pd.read_excel(file_ordenes)
    ventanas_df = pd.read_excel("JC Delivery Schedule.xlsx", engine="openpyxl")

    ordenes_df.columns = ordenes_df.columns.str.strip()
    ventanas_df.columns = ventanas_df.columns.str.strip()

    # Función para convertir hora
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

    ordenes_df["HoraExcel"] = ordenes_df["Require Time"].apply(convertir_hora)
    ordenes_df["FechaHoraEntrega"] = pd.to_datetime(ordenes_df["Require Date"], errors="coerce") + ordenes_df["HoraExcel"]
    ordenes_df["DiaSemana"] = ordenes_df["FechaHoraEntrega"].dt.day_name(locale='es_ES').str.upper()

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

    def asignar_ventana(fila):
        proveedor = str(fila["Vendor"]).strip().upper()
        dia_actual = str(fila["DiaSemana"]).strip().upper()
        hora = fila["FechaHoraEntrega"]
        if pd.isna(hora):
            return "Sin fecha", None

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

    ordenes_df[["VentanaAsignada", "DiaVentana"]] = ordenes_df.apply(lambda fila: pd.Series(asignar_ventana(fila)), axis=1)

    resultado = ordenes_df[["JIT Call No", "Vendor", "FechaHoraEntrega", "DiaSemana", "VentanaAsignada", "DiaVentana"]]

    st.write("✅ Procesamiento completo. Aquí están las primeras filas:")
    st.dataframe(resultado.head(20))

    # ✅ Botón para descargar Excel
    output = io.BytesIO()
    resultado.to_excel(output, index=False)
    output.seek(0)

    st.download_button(
        label="Descargar resultado en Excel",
        data=output,
        file_name="Asignaciones_Ventanas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
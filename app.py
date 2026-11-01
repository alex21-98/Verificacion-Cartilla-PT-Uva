import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from io import BytesIO

# Configuración de la página web
st.set_page_config(
    page_title="Validador de Calidad - Uvas | Agroindustria",
    page_icon="🍇",
    layout="wide"
)

st.title("🍇 Sistema de Validación y Auditoría de Calidad - Uvas")
st.markdown("""
Esta herramienta permite a los asistentes de planta cargar el reporte de Excel, auditar la coherencia lógica de las **pudriciones y bayas rajadas** por rangos de racimos, y **verificar de forma visual la foto de evidencia (S3)** para corregir la información de manera rápida y segura.
""")

# Componente para subir el archivo Excel
uploaded_file = st.file_uploader("Sube tu archivo Excel de evaluaciones (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    @st.cache_data
    def load_data(file):
        df = pd.read_excel(file, sheet_name=0)
        return df

    df_original = load_data(uploaded_file)
    df = df_original.copy()
    
    # Barra lateral de control
    st.sidebar.header("Opciones de Visualización")
    filtro_estado = st.sidebar.radio(
        "Mostrar registros:",
        ["Todos los registros", "Solo registros con errores (Para corregir)", "Solo registros correctos (OK)"]
    )

    # Búsqueda inteligente de columnas de pudrición y rajaduras
    def find_col(keyword_list):
        for col in df.columns:
            if all(k in col.upper() for k in keyword_list):
                return col
        return None

    col_r13 = find_col(['RAC', '1-3'])
    col_p13 = find_col(['PUDRICIÓN', '1-3']) or find_col(['PUDRICION', '1-3'])
    col_br13 = find_col(['RAJADURAS', '1-3']) or find_col(['RAJAD', '1-3'])

    col_r45 = find_col(['RAC', '4-5'])
    col_p45 = find_col(['PUDRICIÓN', '4-5']) or find_col(['PUDRICION', '4-5'])
    col_br45 = find_col(['RAJADURAS', '4-5']) or find_col(['RAJAD', '4-5'])

    col_r6 = find_col(['RAC', '6'])
    col_p6 = find_col(['PUDRICIÓN', '6']) or find_col(['PUDRICION', '6'])
    col_br6 = find_col(['RAJADURAS', '6']) or find_col(['RAJAD', '6'])

    def safe_float(val):
        try:
            return float(val) if pd.notna(val) else 0.0
        except:
            return 0.0

    # Lógica estricta de auditoría
    auditoria_resultados = []
    detalle_errores = []

    for idx, row in df.iterrows():
        errores_fila = []
        
        # 1. Rango [1-3]
        if col_r13 and col_p13 and col_br13:
            r13 = safe_float(row.get(col_r13, 0))
            p13 = safe_float(row.get(col_p13, 0))
            br13 = safe_float(row.get(col_br13, 0))
            tot13 = p13 + br13
            if r13 > 0 and not (1 * r13 <= tot13 <= 3 * r13):
                errores_fila.append(f"[Rango 1-3]: {int(r13)} racimos exigen de {int(1*r13)} a {int(3*r13)} bayas, pero hay {int(tot13)}")
            elif r13 == 0 and tot13 > 0:
                errores_fila.append(f"[Rango 1-3]: Hay {int(tot13)} bayas pero 0 racimos indicados")

        # 2. Rango [4-5]
        if col_r45 and col_p45 and col_br45:
            r45 = safe_float(row.get(col_r45, 0))
            p45 = safe_float(row.get(col_p45, 0))
            br45 = safe_float(row.get(col_br45, 0))
            tot45 = p45 + br45
            if r45 > 0 and not (4 * r45 <= tot45 <= 5 * r45):
                errores_fila.append(f"[Rango 4-5]: {int(r45)} racimos exigen de {int(4*r45)} a {int(5*r45)} bayas, pero hay {int(tot45)}")
            elif r45 == 0 and tot45 > 0:
                errores_fila.append(f"[Rango 4-5]: Hay {int(tot45)} bayas pero 0 racimos indicados")

        # 3. Rango [6+]
        if col_r6 and col_p6 and col_br6:
            r6 = safe_float(row.get(col_r6, 0))
            p6 = safe_float(row.get(col_p6, 0))
            br6 = safe_float(row.get(col_br6, 0))
            tot6 = p6 + br6
            if r6 > 0 and tot6 < 6 * r6:
                errores_fila.append(f"[Rango 6+]: {int(r6)} racimos exigen mínimo {int(6*r6)} bayas, pero hay {int(tot6)}")
            elif r6 == 0 and tot6 > 0:
                errores_fila.append(f"[Rango 6+]: Hay {int(tot6)} bayas pero 0 racimos indicados")

        if errores_fila:
            msg = " | ".join(errores_fila)
            auditoria_resultados.append("ERROR: " + msg)
            detalle_errores.append({
                "Fila Excel": idx + 2,
                "ID Registro": row.get('Id', idx),
                "Motivo del Error": msg
            })
        else:
            auditoria_resultados.append("OK")

    df['AUDITORIA_ESTADO'] = auditoria_resultados
    df_errores_log = pd.DataFrame(detalle_errores)

    # Panel de Métricas
    st.subheader("📊 Resumen de Validación")
    total_filas = len(df)
    total_errores = len(df_errores_log)
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Registros Evaluados", total_filas)
    col2.metric("Registros Conformes (OK)", total_filas - total_errores)
    col3.metric("Registros con Errores Lógicos", total_errores, delta_color="inverse")

    # Filtrar según opción de la barra lateral
    if filtro_estado == "Solo registros con errores (Para corregir)":
        df_display = df[df['AUDITORIA_ESTADO'] != 'OK']
    elif filtro_estado == "Solo registros correctos (OK)":
        df_display = df[df['AUDITORIA_ESTADO'] == 'OK']
    else:
        df_display = df

    st.markdown("---")
    
    # Sección Principal dividida en dos columnas: Tabla y Visor de Fotos
    col_izq, col_der = st.columns([1.3, 0.7])

    with col_izq:
        st.subheader("📋 Detalle de Registros")
        st.dataframe(df_display, height=450)

        # Botón para descargar Excel con reporte de errores
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Base Evaluada', index=False)
            if not df_errores_log.empty:
                df_errores_log.to_excel(writer, sheet_name='Reporte de Errores', index=False)
        processed_data = output.getvalue()

        st.download_button(
            label="📥 Descargar Excel con Reporte de Errores",
            data=processed_data,
            file_name="Reporte_Validacion_Calidad.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_der:
        st.subheader("🖼️ Visor de Evidencia Visual (S3)")
        st.markdown("Selecciona una **Fila de Excel** para corroborar los datos directamente con la foto tomada en campo:")

        # Selector de número de fila de Excel
        fila_seleccionada = st.selectbox("Elige la Fila de Excel a revisar:", df.index + 2)
        
        # Extraer hipervínculos del archivo original con openpyxl
        wb_orig = openpyxl.load_workbook(uploaded_file)
        ws_orig = wb_orig.active
        headers_orig = [cell.value for cell in ws_orig[1]]
        
        foto_cols = [c for c in df.columns if 'FOTO' in c.upper() or 'TRAZABILIDAD' in c.upper()]
        
        row_evaluada = df.loc[fila_seleccionada - 2]
        estado_fila = row_evaluada['AUDITORIA_ESTADO']
        
        if estado_fila == "OK":
            st.success(f"Fila {fila_seleccionada}: Estado **OK** ✅")
        else:
            st.error(f"Fila {fila_seleccionada}: **{estado_fila}** ⚠️")

        if foto_cols:
            for col_name in foto_cols:
                if col_name in headers_orig:
                    c_idx = headers_orig.index(col_name) + 1
                    cell_obj = ws_orig.cell(row=fila_seleccionada, column=c_idx)
                    url = cell_obj.hyperlink.target if cell_obj.hyperlink else (cell_obj.value if isinstance(cell_obj.value, str) and cell_obj.value.startswith('http') else None)
                    
                    if url:
                        st.markdown(f"**{col_name}**")
                        st.markdown(f"[🔗 Abrir imagen completa en pestaña]({url})", unsafe_allow_html=True)
                        try:
                            st.image(url, caption=f"Evidencia Fila {fila_seleccionada} - {col_name}", use_container_width=True)
                        except Exception:
                            st.warning("No se pudo previsualizar la miniatura debido a restricciones de red, pero puedes hacer clic en el enlace superior.")
                    else:
                        st.info(f"No hay enlace de foto en {col_name} para esta fila.")
        else:
            st.info("No se detectaron columnas de fotos de evidencia en este archivo.")

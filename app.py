import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Validador y Auditor de Calidad - Uvas", page_icon="🍇", layout="wide")

st.title("🍇 Sistema Web de Auditoría y Control de Calidad - Uvas")
st.markdown("Sube tu archivo Excel para auditar la coherencia exacta de **pudrición y bayas rajadas** por rangos de racimos y verificar las fotos de evidencia.")

uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    @st.cache_data
    def load_data(file):
        df = pd.read_excel(file, sheet_name=0)
        return df

    df = load_data(uploaded_file).copy()
    
    st.sidebar.header("Panel de Control")
    show_only_errors = st.sidebar.checkbox("Mostrar solo registros con errores lógicos", value=False)
    
    # Búsqueda inteligente de columnas
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

    auditoria_resultados = []
    detalle_errores = []

    for idx, row in df.iterrows():
        errores_fila = []
        
        # 1. Validación [1-3]
        if col_r13 and col_p13 and col_br13:
            r13 = safe_float(row.get(col_r13, 0))
            p13 = safe_float(row.get(col_p13, 0))
            br13 = safe_float(row.get(col_br13, 0))
            tot13 = p13 + br13
            if r13 > 0 and not (1 * r13 <= tot13 <= 3 * r13):
                errores_fila.append(f"[1-3]: {int(r13)} racimos exigen entre {int(1*r13)} y {int(3*r13)} bayas, pero hay {int(tot13)}.")
            elif r13 == 0 and tot13 > 0:
                errores_fila.append(f"[1-3]: Hay {int(tot13)} bayas pero 0 racimos indicados.")

        # 2. Validación [4-5]
        if col_r45 and col_p45 and col_br45:
            r45 = safe_float(row.get(col_r45, 0))
            p45 = safe_float(row.get(col_p45, 0))
            br45 = safe_float(row.get(col_br45, 0))
            tot45 = p45 + br45
            if r45 > 0 and not (4 * r45 <= tot45 <= 5 * r45):
                errores_fila.append(f"[4-5]: {int(r45)} racimos exigen entre {int(4*r45)} y {int(5*r45)} bayas, pero hay {int(tot45)}.")
            elif r45 == 0 and tot45 > 0:
                errores_fila.append(f"[4-5]: Hay {int(tot45)} bayas pero 0 racimos indicados.")

        # 3. Validación [6+]
        if col_r6 and col_p6 and col_br6:
            r6 = safe_float(row.get(col_r6, 0))
            p6 = safe_float(row.get(col_p6, 0))
            br6 = safe_float(row.get(col_br6, 0))
            tot6 = p6 + br6
            if r6 > 0 and tot6 < 6 * r6:
                errores_fila.append(f"[6+]: {int(r6)} racimos exigen mínimo {int(6*r6)} bayas, pero hay {int(tot6)}.")
            elif r6 == 0 and tot6 > 0:
                errores_fila.append(f"[6+]: Hay {int(tot6)} bayas pero 0 racimos indicados.")

        if errores_fila:
            msg = " | ".join(errores_fila)
            auditoria_resultados.append("ERROR LÓGICO: " + msg)
            detalle_errores.append({
                "Fila Excel": idx + 2,
                "ID Registro": row.get('Id', idx),
                "Errores Encontrados": msg
            })
        else:
            auditoria_resultados.append("OK")

    df['AUDITORIA_ESTADO'] = auditoria_resultados
    df_errores_log = pd.DataFrame(detalle_errores)

    # Métricas de resumen
    st.subheader("📊 Resumen General de Auditoría")
    total_filas = len(df)
    total_errores = len(df_errores_log)
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Registros", total_filas)
    c2.metric("Conformes (OK)", total_filas - total_errores)
    c3.metric("Con Errores Lógicos", total_errores, delta_color="inverse")

    if show_only_errors:
        df_display = df[df['AUDITORIA_ESTADO'] != 'OK']
    else:
        df_display = df

    st.subheader("📋 Base de Datos Evaluada")
    st.dataframe(df_display)

    if not df_errores_log.empty:
        st.subheader("⚠️ Reporte de Errores (Segunda Hoja del Excel)")
        st.dataframe(df_errores_log)

    # Generar Excel con dos hojas
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Base Evaluada', index=False)
        if not df_errores_log.empty:
            df_errores_log.to_excel(writer, sheet_name='Reporte de Errores', index=False)
    processed_data = output.getvalue()

    st.download_button(
        label="📥 Descargar Excel con Reporte de Errores",
        data=processed_data,
        file_name="Reporte_Auditoria_Calidad_Uvas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.markdown("---")
    st.markdown("### 🖼️ Visor de Fotos y Enlaces de Campo")
    st.markdown("Selecciona la **Fila de Excel** para abrir la foto y contrastarla de inmediato con los datos numéricos:")

    foto_cols = [c for c in df.columns if 'FOTO' in c.upper() or 'TRAZABILIDAD' in c.upper()]
    if foto_cols:
        selected_row_idx = st.selectbox("Elige la Fila Excel a inspeccionar:", df.index + 2)
        
        wb_orig = openpyxl.load_workbook(uploaded_file)
        ws_orig = wb_orig.active
        headers_orig = [cell.value for cell in ws_orig[1]]
        
        row_eval = df.loc[selected_row_idx - 2]
        st.info(f"**Estado de la Fila {selected_row_idx}:** {row_eval['AUDITORIA_ESTADO']}")

        cols_f = st.columns(len(foto_cols))
        for i, col_name in enumerate(foto_cols):
            with cols_f[i]:
                st.markdown(f"**{col_name}**")
                if col_name in headers_orig:
                    c_idx = headers_orig.index(col_name) + 1
                    cell_obj = ws_orig.cell(row=selected_row_idx, column=c_idx)
                    url = cell_obj.hyperlink.target if cell_obj.hyperlink else (cell_obj.value if isinstance(cell_obj.value, str) and cell_obj.value.startswith('http') else None)
                    
                    if url:
                        st.markdown(f"[🔗 Abrir Imagen en Nueva Pestaña]({url})", unsafe_allow_html=True)
                        st.image(url, width=280, caption=f"Fila {selected_row_idx}")
                    else:
                        st.info("Sin enlace disponible en esta celda.")
    else:
        st.info("No se detectaron columnas de fotos.")
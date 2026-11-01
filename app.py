import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from io import BytesIO

st.set_page_config(
    page_title="Validador y Auditor de Calidad - Uvas",
    page_icon="🍇",
    layout="wide"
)

st.title("🍇 Sistema Avanzado de Auditoría y Control de Calidad - Uvas")
st.markdown("""
Esta aplicación web procesa tu reporte de Excel aplicando las reglas de negocio de **pudrición y bayas rajadas**, valida que la suma de racimos por rangos no supere el total de la muestra, extrae los hipervínculos reales de las fotos en S3 y genera un reporte detallado en múltiples hojas.
""")

uploaded_file = st.file_uploader("Sube tu archivo Excel de evaluaciones (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    @st.cache_data
    def load_data(file):
        df = pd.read_excel(file, sheet_name=0)
        return df

    df = load_data(uploaded_file).copy()
    
    # Cargar también el libro con openpyxl para extraer los hipervínculos reales de las fotos
    wb_orig = openpyxl.load_workbook(uploaded_file)
    ws_orig = wb_orig.active
    headers_orig = [cell.value for cell in ws_orig[1]]

    st.sidebar.header("Opciones de Control")
    filtro_estado = st.sidebar.radio(
        "Filtrar registros en tabla:",
        ["Todos los registros", "Solo registros con errores lógicos", "Solo registros conformes (OK)"]
    )

    # Búsqueda inteligente de columnas clave
    def find_col(keyword_list):
        for col in df.columns:
            if all(k in col.upper() for k in keyword_list):
                return col
        return None

    col_r_total = find_col(['N° RACIMOS', 'CLAMSHELL']) or find_col(['RACIMOS'])
    
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

    # Auditoría estricta basada en tus reglas
    auditoria_resultados = []
    detalle_errores = []

    for idx, row in df.iterrows():
        errores_fila = []
        
        r13 = safe_float(row.get(col_r13, 0)) if col_r13 else 0.0
        p13 = safe_float(row.get(col_p13, 0)) if col_p13 else 0.0
        br13 = safe_float(row.get(col_br13, 0)) if col_br13 else 0.0
        tot13 = p13 + br13

        r45 = safe_float(row.get(col_r45, 0)) if col_r45 else 0.0
        p45 = safe_float(row.get(col_p45, 0)) if col_p45 else 0.0
        br45 = safe_float(row.get(col_br45, 0)) if col_br45 else 0.0
        tot45 = p45 + br45

        r6 = safe_float(row.get(col_r6, 0)) if col_r6 else 0.0
        p6 = safe_float(row.get(col_p6, 0)) if col_p6 else 0.0
        br6 = safe_float(row.get(col_br6, 0)) if col_br6 else 0.0
        tot6 = p6 + br6

        # 1. Regla de suma de racimos entre rangos (No puede exceder el total de racimos evaluados si existe la columna)
        suma_racimos_afectados = r13 + r45 + r6
        if col_r_total:
            total_muestra = safe_float(row.get(col_r_total, 0))
            if total_muestra > 0 and suma_racimos_afectados > total_muestra:
                errores_fila.append(f"[Exceso de Racimos]: La suma de racimos afectados ({int(suma_racimos_afectados)}) supera el total evaluado en la muestra ({int(total_muestra)}).")

        # 2. Validación Rango [1-3]
        if r13 > 0:
            if not (1 * r13 <= tot13 <= 3 * r13):
                errores_fila.append(f"[Rango 1-3]: {int(r13)} racimos exigen de {int(1*r13)} a {int(3*r13)} bayas, pero suman {int(tot13)}.")
        elif r13 == 0 and tot13 > 0:
            errores_fila.append(f"[Rango 1-3]: Hay {int(tot13)} bayas registradas pero 0 racimos indicados.")

        # 3. Validación Rango [4-5]
        if r45 > 0:
            if not (4 * r45 <= tot45 <= 5 * r45):
                errores_fila.append(f"[Rango 4-5]: {int(r45)} racimos exigen de {int(4*r45)} a {int(5*r45)} bayas, pero suman {int(tot45)}.")
        elif r45 == 0 and tot45 > 0:
            errores_fila.append(f"[Rango 4-5]: Hay {int(tot45)} bayas registradas pero 0 racimos indicados.")

        # 4. Validación Rango [6+] (Mínimo N * 6 bayas)
        if r6 > 0:
            if tot6 < 6 * r6:
                errores_fila.append(f"[Rango 6+]: {int(r6)} racimos exigen mínimo {int(6*r6)} bayas, pero suman {int(tot6)}.")
        elif r6 == 0 and tot6 > 0:
            errores_fila.append(f"[Rango 6+]: Hay {int(tot6)} bayas registradas pero 0 racimos indicados.")

        if errores_fila:
            msg = " | ".join(errores_fila)
            auditoria_resultados.append("ERROR LÓGICO: " + msg)
            detalle_errores.append({
                "Fila Excel": idx + 2,
                "ID Registro": row.get('Id', idx),
                "Motivo del Error": msg
            })
        else:
            auditoria_resultados.append("OK")

    df['AUDITORIA_ESTADO'] = auditoria_resultados
    df_errores_log = pd.DataFrame(detalle_errores)

    # Métricas superiores
    st.subheader("📊 Resumen General de Auditoría")
    total_filas = len(df)
    total_errores = len(df_errores_log)
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Registros Evaluados", total_filas)
    c2.metric("Registros Conformes (OK)", total_filas - total_errores)
    c3.metric("Registros con Errores Lógicos", total_errores, delta_color="inverse")

    if filtro_estado == "Solo registros con errores lógicos":
        df_display = df[df['AUDITORIA_ESTADO'] != 'OK']
    elif filtro_estado == "Solo registros conformes (OK)":
        df_display = df[df['AUDITORIA_ESTADO'] == 'OK']
    else:
        df_display = df

    st.markdown("---")
    
    col_izq, col_der = st.columns([1.3, 0.7])

    with col_izq:
        st.subheader("📋 Tabla de Registros Evaluados")
        st.dataframe(df_display, height=450)

        # Descarga de Excel con múltiples hojas
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Base Evaluada', index=False)
            if not df_errores_log.empty:
                df_errores_log.to_excel(writer, sheet_name='Reporte de Errores', index=False)
        processed_data = output.getvalue()

        st.download_button(
            label="📥 Descargar Excel con Reporte de Errores y Auditoría",
            data=processed_data,
            file_name="Reporte_Auditoria_Calidad_Uvas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_der:
        st.subheader("🖼️ Visor y Verificador de Fotos (S3)")
        st.markdown("Selecciona la **Fila de Excel** para extraer el enlace real de la foto y corroborar visualmente el conteo:")

        fila_seleccionada = st.selectbox("Elige la Fila de Excel a revisar:", df.index + 2)
        
        row_evaluada = df.loc[fila_seleccionada - 2]
        estado_fila = row_evaluada['AUDITORIA_ESTADO']
        
        if estado_fila == "OK":
            st.success(f"Fila {fila_seleccionada}: Estado **OK** ✅")
        else:
            st.error(f"Fila {fila_seleccionada}: **{estado_fila}** ⚠️")

        # Buscar columnas de fotos o evidencia
        foto_cols = [c for c in df.columns if 'FOTO' in c.upper() or 'TRAZABILIDAD' in c.upper() or 'EVIDENCIA' in c.upper()]
        
        if foto_cols:
            for col_name in foto_cols:
                if col_name in headers_orig:
                    c_idx = headers_orig.index(col_name) + 1
                    cell_obj = ws_orig.cell(row=fila_seleccionada, column=c_idx)
                    
                    # Extraer hipervínculo real de openpyxl o texto URL
                    url = cell_obj.hyperlink.target if cell_obj.hyperlink else (cell_obj.value if isinstance(cell_obj.value, str) and cell_obj.value.startswith('http') else None)
                    
                    if url:
                        st.markdown(f"**{col_name}**")
                        st.markdown(f"[🔗 Abrir Fotografía en el Navegador]({url})", unsafe_allow_html=True)
                        try:
                            st.image(url, caption=f"Evidencia Fila {fila_seleccionada}", use_container_width=True)
                        except Exception:
                            st.warning("No se pudo previsualizar la imagen directamente por seguridad de red, pero haz clic en el enlace de arriba para verla.")
                    else:
                        st.info(f"Sin enlace de foto en {col_name} para esta fila.")
        else:
            st.info("No se detectaron columnas de fotos en este archivo.")

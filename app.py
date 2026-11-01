import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

st.set_page_config(page_title="Auditoría de Calidad - Uvas", page_icon="🍇", layout="wide")

st.title("🍇 Sistema de Auditoría y Control de Calidad - Uvas")
st.markdown("Valida rangos de pudrición/rajaduras, límite de muestra total y verifica fotos de evidencia (S3).")

uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0).copy()
    
    wb_orig = openpyxl.load_workbook(uploaded_file, data_only=False)
    ws_orig = wb_orig.active
    headers_orig = [str(cell.value).strip() if cell.value else "" for cell in ws_orig[1]]

    st.sidebar.header("Opciones")
    filtro = st.sidebar.radio("Mostrar en tabla:", ["Todos los registros", "Solo Errores", "Solo Correctos (OK)"])

    # Función estricta para encontrar columnas sin confundirse
    def find_col_exact(keywords, exclude=None):
        if exclude is None: exclude = []
        for col in df.columns:
            col_up = col.upper()
            if all(k in col_up for k in keywords) and all(e not in col_up for e in exclude):
                return col
        return None

    # Aislar la columna del TOTAL de racimos de la muestra (excluyendo las de defectos)
    col_muestra_total = find_col_exact(['RAC'], exclude=['PUDRIC', 'RAJAD', '1-3', '4-5', '6', '%', 'DEFECTO']) or \
                        find_col_exact(['CLAM'], exclude=['PUDRIC', 'RAJAD', '1-3', '4-5', '6', '%', 'DEFECTO'])

    # Búsqueda EXACTA de columnas por rango
    # [1-3]
    col_r13 = find_col_exact(['RAC./CLAMSHELL', '(1-3']) or find_col_exact(['RAC', '1-3', 'PUDRIC'], exclude=['ENCONTRADAS'])
    col_p13 = find_col_exact(['PUDRICIÓN', 'ENCONTRADAS', '1-3']) or find_col_exact(['PUDRICION', 'ENCONTRADAS', '1-3'])
    col_br13 = find_col_exact(['RAJADURAS', 'ENCONTRADAS', '1-3'])

    # [4-5]
    col_r45 = find_col_exact(['RAC./CLAMSHELL', '(4-5']) or find_col_exact(['RAC', '4-5', 'PUDRIC'], exclude=['ENCONTRADAS'])
    col_p45 = find_col_exact(['PUDRICIÓN', 'ENCONTRADAS', '4-5']) or find_col_exact(['PUDRICION', 'ENCONTRADAS', '4-5'])
    col_br45 = find_col_exact(['RAJADURAS', 'ENCONTRADAS', '4-5'])

    # [6+]
    col_r6 = find_col_exact(['RAC./CLAMSHELL', '(6']) or find_col_exact(['RAC', '6', 'PUDRIC'], exclude=['ENCONTRADAS'])
    col_p6 = find_col_exact(['PUDRICIÓN', 'ENCONTRADAS', '6']) or find_col_exact(['PUDRICION', 'ENCONTRADAS', '6'])
    col_br6 = find_col_exact(['RAJADURAS', 'ENCONTRADAS', '6'])

    def safe_float(val):
        try: return float(val) if pd.notna(val) else 0.0
        except: return 0.0

    resultados = []
    log_errores = []

    for idx, row in df.iterrows():
        errores = []
        
        r13 = safe_float(row.get(col_r13, 0)); p13 = safe_float(row.get(col_p13, 0)); br13 = safe_float(row.get(col_br13, 0))
        r45 = safe_float(row.get(col_r45, 0)); p45 = safe_float(row.get(col_p45, 0)); br45 = safe_float(row.get(col_br45, 0))
        r6  = safe_float(row.get(col_r6, 0));  p6  = safe_float(row.get(col_p6, 0));  br6  = safe_float(row.get(col_br6, 0))
        
        t13 = p13 + br13
        t45 = p45 + br45
        t6  = p6 + br6

        # REGLA ORO: Límite de Muestra
        suma_racimos_afectados = r13 + r45 + r6
        if col_muestra_total:
            total_muestra = safe_float(row.get(col_muestra_total, 0))
            if total_muestra > 0 and suma_racimos_afectados > total_muestra:
                errores.append(f"[LÍMITE]: Suma de racimos afectados ({int(suma_racimos_afectados)}) supera la muestra total ({int(total_muestra)})")

        # REGLA [1-3]
        if r13 > 0:
            if not (1 * r13 <= t13 <= 3 * r13):
                errores.append(f"[1-3]: {int(r13)} rac exigen entre {int(1*r13)} y {int(3*r13)} bayas, hay {int(t13)}")
        elif t13 > 0:
            errores.append(f"[1-3]: Hay {int(t13)} bayas pero 0 racimos indicados")

        # REGLA [4-5]
        if r45 > 0:
            if not (4 * r45 <= t45 <= 5 * r45):
                errores.append(f"[4-5]: {int(r45)} rac exigen entre {int(4*r45)} y {int(5*r45)} bayas, hay {int(t45)}")
        elif t45 > 0:
            errores.append(f"[4-5]: Hay {int(t45)} bayas pero 0 racimos indicados")

        # REGLA [6+]
        if r6 > 0:
            if t6 < 6 * r6:
                errores.append(f"[6+]: {int(r6)} rac exigen mínimo {int(6*r6)} bayas, hay {int(t6)}")
        elif t6 > 0:
            errores.append(f"[6+]: Hay {int(t6)} bayas pero 0 racimos indicados")

        if errores:
            msg = " | ".join(errores)
            resultados.append("ERROR: " + msg)
            log_errores.append({"Fila Excel": idx + 2, "ID Registro": row.get('Id', f"Fila {idx+2}"), "Error": msg})
        else:
            resultados.append("OK")

    df['ESTADO_AUDITORIA'] = resultados
    df_errores = pd.DataFrame(log_errores)

    # UI STREAMLIT
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Registros", len(df))
    c2.metric("Correctos (OK)", len(df) - len(df_errores))
    c3.metric("Con Errores", len(df_errores), delta_color="inverse")

    if filtro == "Solo Errores": df_mostrar = df[df['ESTADO_AUDITORIA'] != 'OK']
    elif filtro == "Solo Correctos (OK)": df_mostrar = df[df['ESTADO_AUDITORIA'] == 'OK']
    else: df_mostrar = df

    st.markdown("---")
    col_izq, col_der = st.columns([1.5, 1])

    with col_izq:
        st.subheader("📋 Base Evaluada")
        st.dataframe(df_mostrar, height=500)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Base Evaluada', index=False)
            if not df_errores.empty:
                df_errores.to_excel(writer, sheet_name='Reporte de Errores', index=False)
        
        st.download_button(label="📥 Descargar Excel de Auditoría", data=output.getvalue(), file_name="Auditoria_Uvas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with col_der:
        st.subheader("🖼️ Verificación Visual")
        fila_sel = st.selectbox("Elige la fila Excel a validar:", df.index + 2)
        estado = df.loc[fila_sel - 2, 'ESTADO_AUDITORIA']
        if estado == "OK": st.success(f"Fila {fila_sel}: OK ✅")
        else: st.error(f"Fila {fila_sel}: {estado} ⚠️")

        foto_cols = [c for c in df.columns if 'EVIDENCIA' in c.upper() or 'FOTO' in c.upper()]
        hay_foto = False
        for c_name in foto_cols:
            if c_name in headers_orig:
                c_idx = headers_orig.index(c_name) + 1
                cell = ws_orig.cell(row=fila_sel, column=c_idx)
                
                url = None
                if cell.hyperlink and cell.hyperlink.target: url = cell.hyperlink.target
                elif isinstance(cell.value, str) and "http" in cell.value: url = cell.value

                if url:
                    hay_foto = True
                    st.markdown(f"**{c_name}**")
                    st.markdown(f"[🔗 Ver Imagen]({url})", unsafe_allow_html=True)
                    try: st.image(url, use_container_width=True)
                    except: st.warning("Haz clic en el enlace para abrir la evidencia.")
        
        if not hay_foto: st.info("Sin evidencia fotográfica.")

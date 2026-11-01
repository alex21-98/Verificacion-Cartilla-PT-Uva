import streamlit as st
import pandas as pd
import openpyxl
import re
from io import BytesIO

# Configuración de página
st.set_page_config(page_title="Auditoría de Calidad - Uvas", page_icon="🍇", layout="wide")

# Estilos CSS corporativos
st.markdown("""
    <style>
        .titulo-principal { color: #002060; font-weight: 700; margin-bottom: 0px;}
        .subtitulo { color: #4A4A4A; font-size: 1.1rem; margin-bottom: 20px;}
        div[data-testid="stMetricValue"] { color: #002060; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="titulo-principal">🍇 Sistema de Auditoría y Control de Calidad - Uvas</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitulo">Valida rangos lógicos, límite exacto de la muestra y revisa evidencias fotográficas interactivas.</p>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0).copy()
    
    wb_orig = openpyxl.load_workbook(uploaded_file, data_only=False)
    ws_orig = wb_orig.active
    headers_orig = [str(cell.value).strip() if cell.value else "" for cell in ws_orig[1]]

    # --- BARRA LATERAL (CONFIGURACIÓN Y FILTROS) ---
    st.sidebar.markdown("### ⚙️ Configuración de Reglas")
    
    # Detección automática de la columna de muestra exacta aportada
    def normalizar_nombre(c):
        return str(c).upper().replace(" ", "-").replace("_", "-")

    posibles_totales = [c for c in df.columns if 'DATOS-DE-EVALUACION-N-RACIMOS-CLAMSHELL' in normalizar_nombre(c) or 'DATOS-DE-EVALUACION' in normalizar_nombre(c) or 'MUESTRA' in normalizar_nombre(c)]
    indice_por_defecto = df.columns.get_loc(posibles_totales[0]) if posibles_totales else 0
    
    # Selector EXPLÍCITO para la regla de total de racimos evaluados
    col_muestra_total = st.sidebar.selectbox(
        "Columna 'Total de Racimos Evaluados':", 
        options=df.columns, 
        index=indice_por_defecto,
        help="El sistema usará esta columna para validar que la suma de racimos afectados no supere la muestra total evaluada."
    )

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔍 Filtro de Tabla")
    filtro = st.sidebar.radio("Mostrar:", ["Todos los registros", "Solo Errores", "Solo Correctos (OK)"])

    def find_col_exact(keywords, exclude=None):
        if exclude is None: exclude = []
        for col in df.columns:
            col_up = col.upper()
            if all(k in col_up for k in keywords) and all(e not in col_up for e in exclude):
                return col
        return None

    # Identificación estricta de columnas de pudrición/rajaduras
    col_r13 = find_col_exact(['RAC./CLAMSHELL', '(1-3']) or find_col_exact(['RAC', '1-3', 'PUDRIC'], exclude=['ENCONTRADAS'])
    col_p13 = find_col_exact(['PUDRICIÓN', 'ENCONTRADAS', '1-3']) or find_col_exact(['PUDRICION', 'ENCONTRADAS', '1-3'])
    col_br13 = find_col_exact(['RAJADURAS', 'ENCONTRADAS', '1-3'])

    col_r45 = find_col_exact(['RAC./CLAMSHELL', '(4-5']) or find_col_exact(['RAC', '4-5', 'PUDRIC'], exclude=['ENCONTRADAS'])
    col_p45 = find_col_exact(['PUDRICIÓN', 'ENCONTRADAS', '4-5']) or find_col_exact(['PUDRICION', 'ENCONTRADAS', '4-5'])
    col_br45 = find_col_exact(['RAJADURAS', 'ENCONTRADAS', '4-5'])

    col_r6 = find_col_exact(['RAC./CLAMSHELL', '(6']) or find_col_exact(['RAC', '6', 'PUDRIC'], exclude=['ENCONTRADAS'])
    col_p6 = find_col_exact(['PUDRICIÓN', 'ENCONTRADAS', '6']) or find_col_exact(['PUDRICION', 'ENCONTRADAS', '6'])
    col_br6 = find_col_exact(['RAJADURAS', 'ENCONTRADAS', '6'])

    def safe_float(val):
        try: return float(val) if pd.notna(val) else 0.0
        except: return 0.0

    resultados = []
    log_errores = []

    # Bucle Lógico y Matemático
    for idx, row in df.iterrows():
        errores = []
        r13 = safe_float(row.get(col_r13, 0)); p13 = safe_float(row.get(col_p13, 0)); br13 = safe_float(row.get(col_br13, 0))
        r45 = safe_float(row.get(col_r45, 0)); p45 = safe_float(row.get(col_p45, 0)); br45 = safe_float(row.get(col_br45, 0))
        r6  = safe_float(row.get(col_r6, 0));  p6  = safe_float(row.get(col_p6, 0));  br6  = safe_float(row.get(col_br6, 0))
        
        t13 = p13 + br13; t45 = p45 + br45; t6 = p6 + br6
        
        # 1. REGLA ESTRICTA DE LÍMITE TOTAL CON LA COLUMNA ASIGNADA
        suma_racimos_afectados = r13 + r45 + r6
        tot_evaluados = safe_float(row.get(col_muestra_total, 0))
        if tot_evaluados > 0 and suma_racimos_afectados > tot_evaluados:
            errores.append(f"[EXCESO TOTAL]: Hay {int(suma_racimos_afectados)} racimos afectados en total, pero la muestra evaluada es solo de {int(tot_evaluados)}")

        # 2. Reglas de consistencia por rangos
        if r13 > 0:
            if not (1 * r13 <= t13 <= 3 * r13): errores.append(f"[1-3]: {int(r13)} racimos exigen {int(1*r13)} a {int(3*r13)} bayas, hay {int(t13)}")
        elif t13 > 0: errores.append(f"[1-3]: Hay {int(t13)} bayas pero 0 racimos indicados")

        if r45 > 0:
            if not (4 * r45 <= t45 <= 5 * r45): errores.append(f"[4-5]: {int(r45)} racimos exigen {int(4*r45)} a {int(5*r45)} bayas, hay {int(t45)}")
        elif t45 > 0: errores.append(f"[4-5]: Hay {int(t45)} bayas pero 0 racimos indicados")

        if r6 > 0:
            if t6 < 6 * r6: errores.append(f"[6+]: {int(r6)} racimos exigen mín {int(6*r6)} bayas, hay {int(t6)}")
        elif t6 > 0: errores.append(f"[6+]: Hay {int(t6)} bayas pero 0 racimos indicados")

        if errores:
            msg = " | ".join(errores)
            resultados.append("ERROR: " + msg)
            log_errores.append({"Fila Excel": idx + 2, "ID Registro": row.get('Id', f"Fila {idx+2}"), "Error": msg})
        else:
            resultados.append("OK")

    df['ESTADO_AUDITORIA'] = resultados
    columnas_reordenadas = ['ESTADO_AUDITORIA'] + [col for col in df.columns if col != 'ESTADO_AUDITORIA']
    df = df[columnas_reordenadas]
    df_errores = pd.DataFrame(log_errores)

    # Panel de Métricas
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Registros", len(df))
    c2.metric("Correctos (OK)", len(df) - len(df_errores))
    c3.metric("Con Errores", len(df_errores), delta_color="inverse")
    st.markdown("---")

    if filtro == "Solo Errores": df_mostrar = df[df['ESTADO_AUDITORIA'] != 'OK']
    elif filtro == "Solo Correctos (OK)": df_mostrar = df[df['ESTADO_AUDITORIA'] == 'OK']
    else: df_mostrar = df

    # Layout de la Interfaz
    col_izq, col_der = st.columns([1.8, 1.2])

    with col_izq:
        st.markdown("**📋 Tabla de Registros (Selecciona una fila para ver su evidencia)**")
        evento = st.dataframe(
            df_mostrar, 
            height=600, 
            use_container_width=True,
            selection_mode="single-row",
            on_select="rerun"
        )

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Base Evaluada', index=False)
            if not df_errores.empty:
                df_errores.to_excel(writer, sheet_name='Reporte Errores', index=False)
        st.download_button("📥 Descargar Excel del Reporte", data=output.getvalue(), file_name="Auditoria_Uvas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with col_der:
        st.markdown("**🖼️ Evidencia Fotográfica**")
        with st.container(height=600, border=True):
            filas_seleccionadas = evento.selection.rows
            
            if filas_seleccionadas:
                indice_real = df_mostrar.index[filas_seleccionadas[0]]
                fila_sel = indice_real + 2
                
                estado = df.loc[indice_real, 'ESTADO_AUDITORIA']
                if estado == "OK": 
                    st.success(f"Fila Excel {fila_sel}: OK ✅")
                else: 
                    st.error(f"Fila Excel {fila_sel}: {estado} ⚠️")

                foto_cols = [c for c in df.columns if 'EVIDENCIA EXTRA' in c.upper()]
                hay_foto = False
                
                for c_name in foto_cols:
                    if c_name in headers_orig:
                        c_idx = headers_orig.index(c_name) + 1
                        cell = ws_orig.cell(row=fila_sel, column=c_idx)
                        
                        url = None
                        if cell.hyperlink and cell.hyperlink.target:
                            url = cell.hyperlink.target
                        elif cell.value:
                            val_str = str(cell.value)
                            if "http" in val_str:
                                match = re.search(r'(https?://[^\s"]+)', val_str)
                                if match: url = match.group(1).rstrip('")')

                        if url:
                            hay_foto = True
                            st.markdown(f"[🔗 Abrir {c_name} en nueva pestaña]({url})", unsafe_allow_html=True)
                            try: st.image(url, use_container_width=True)
                            except Exception: st.warning("La imagen no se puede incrustar, ábrela desde el enlace superior.")
                
                if not hay_foto: 
                    st.info("Esta fila no contiene imágenes de evidencia extra.")
            else:
                st.info("👆 Selecciona una fila en la tabla de la izquierda para visualizar aquí la evidencia fotográfica.")

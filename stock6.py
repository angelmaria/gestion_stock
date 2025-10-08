# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Análisis Stock Farmacia", layout="wide")

# Estilos personalizados
st.markdown("""
<style>
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 3em;
        font-weight: 500;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Análisis de Stock Farmacéutico")
st.markdown("---")

# Mapeo de prefijos a familias funcionales completas
FAMILIAS_MAP = {
    'ADELG': 'ADELGAZANTES',
    'ANTICEL': 'ANTICELULITICOS',
    'AROMA': 'AROMATERAPIA',
    'DEPORTE': 'DEPORTE',
    'DERMO': 'DERMO',
    'DIETSOE': 'DIET SOE',
    'DIET': 'DIETETICA',
    'EFECSOE': 'EFEC SOE',
    'EFEC': 'EFECTOS',
    'EFP': 'EFP',
    'ESPEC': 'ESPECIALIDAD',
    'ESPECSR': 'ESPECIALIDAD',
    'FITO': 'FITOTERAPIA',
    'HIGBUC': 'HIG.BUCAL',
    'HIGCAP': 'HIG.CAPILAR',
    'HIGCORP': 'HIG.CORPORAL',
    'HOMEO': 'HOMEOPATIA',
    'INFAN': 'INFANTIL',
    'INFANSOE': 'INFANTIL SOE',
    'INSEC': 'INSECTOS',
    'NASOI': 'NARIZ OIDOS',
    'OPTIC': 'OPTICA',
    'ORTO': 'ORTOPEDIA',
    'ORTOSOE': 'ORTOPEDIA SOE',
    'PIEMAN': 'PIES/MANOS',
    'GINEC': 'SALUD GINECOLOGICA',
    'SEX': 'SALUD SEXUAL',
    'SOL': 'SOLARES',
    'VET': 'VETERINARIA',
    'VACUNAS': 'VACUNAS',
    'FORMULAS': 'FORMULAS',
    'ENVASE': 'ENVASE CLINICO'
}

# Funciones auxiliares
def formato_euros(valor):
    """Formatea un número con punto para miles y coma para decimales"""
    return f"{valor:,.2f}€".replace(",", "X").replace(".", ",").replace("X", ".")

def formato_numero(valor):
    """Formatea un número con punto para miles y coma para decimales"""
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def calcular_indice_rotacion(ventas_anuales, stock_actual):
    """Calcula el índice de rotación anual"""
    if stock_actual == 0 or pd.isna(stock_actual):
        return 0
    return round(ventas_anuales / stock_actual, 2)

def extraer_familia(categoria_str):
    """Extrae la familia funcional desde el prefijo de la categoría"""
    if pd.isna(categoria_str):
        return 'SIN CLASIFICAR'
    
    prefijo = str(categoria_str).split('-')[0].strip()
    return FAMILIAS_MAP.get(prefijo, 'OTROS')

@st.cache_data
def procesar_excel(uploaded_file, dias_abierto, stock_min_dias, stock_max_dias, dias_cobertura_optimo, margen_seguridad):
    """Procesa el archivo Excel y calcula todos los valores"""
    df = pd.read_excel(uploaded_file)
    
    # Detectar columna TOTAL
    col_total = None
    for col in df.columns:
        if 'total' in str(col).lower() and 'ventas' not in str(col).lower():
            col_total = col
            break
    
    # Si no hay TOTAL, buscar columnas mensuales
    if col_total:
        df['Total_Ventas'] = pd.to_numeric(df[col_total], errors='coerce').fillna(0)
    else:
        columnas_ventas = []
        meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 
                'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
        
        for col in df.columns:
            col_lower = str(col).lower()
            if 'ventas' in col_lower or any(f'_{mes[:2]}' in col_lower or f'{mes}' in col_lower for mes in meses):
                columnas_ventas.append(col)
        
        if columnas_ventas:
            df['Total_Ventas'] = df[columnas_ventas].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)
        else:
            df['Total_Ventas'] = 0
    
    # Calcular ventas diarias
    df['Vtas_Dia'] = df['Total_Ventas'] / dias_abierto
    
    # Categorizar productos (A, B, C, D, E)
    def categorizar_producto(ventas_anuales):
        if ventas_anuales > 260:
            return 'A'
        elif ventas_anuales >= 52:
            return 'B'
        elif ventas_anuales >= 12:
            return 'C'
        elif ventas_anuales >= 1:
            return 'D'
        else:
            return 'E'
    
    df['Categoria'] = df['Total_Ventas'].apply(categorizar_producto)
    
    # CORRECCIÓN: Calcular stocks CORRECTAMENTE aplicando margen a TODAS las categorías
    def calcular_stocks(row):
        cat = row['Categoria']
        vtas_dia = row['Vtas_Dia']
        
        if cat == 'A' or cat == 'B':
            stock_ideal = vtas_dia * dias_cobertura_optimo
            stock_min = vtas_dia * stock_min_dias
            stock_limite = stock_ideal * (1 + margen_seguridad)
        elif cat == 'C':
            stock_ideal = 1
            stock_min = 1
            stock_limite = stock_ideal * (1 + margen_seguridad)  # CORRECCIÓN: Aplicar margen
        elif cat == 'D':
            stock_ideal = 1
            stock_min = 0
            stock_limite = stock_ideal * (1 + margen_seguridad)  # CORRECCIÓN: Aplicar margen
        else:  # E
            stock_ideal = 0
            stock_min = 0
            stock_limite = 0
        
        return pd.Series({
            'Stock_Min_Calc': round(stock_min, 1),
            'Stock_Ideal': round(stock_ideal, 1),
            'Stock_Limite': round(stock_limite, 1)
        })
    
    df[['Stock_Min_Calc', 'Stock_Ideal', 'Stock_Limite']] = df.apply(calcular_stocks, axis=1)
    
    # Buscar y limpiar columnas
    col_stock_actual = None
    col_pvp = None
    col_cn = None
    col_descripcion = None
    col_categoria_funcional = None
    
    for col in df.columns:
        col_lower = str(col).lower()
        if 'stock' in col_lower and ('actual' in col_lower or col_lower == 'stockactual'):
            col_stock_actual = col
        elif col_lower == 'pvp':
            col_pvp = col
        elif col_lower == 'cn' or 'idarti' in col_lower or col_lower == 'codigo':
            if col_cn is None:
                col_cn = col
        elif 'descripcion' in col_lower or 'descripción' in col_lower:
            col_descripcion = col
        elif 'categoria' in col_lower and 'funcional' in col_lower:
            col_categoria_funcional = col
        elif (col_lower == 'categoria' or col_lower == 'categoría') and '-' not in str(df[col].iloc[0] if len(df) > 0 else ''):
            if col_categoria_funcional is None:
                col_categoria_funcional = col
    
    if col_categoria_funcional is None:
        for col in df.columns:
            col_lower = str(col).lower()
            if col_lower == 'categoria' or col_lower == 'categoría':
                col_categoria_funcional = col
                break
    
    # Limpiar PVP
    if col_pvp:
        df[col_pvp] = df[col_pvp].astype(str).str.replace('€', '').str.replace(',', '.').str.strip()
        df[col_pvp] = pd.to_numeric(df[col_pvp], errors='coerce').fillna(0)
    
    # Calcular valores
    if col_stock_actual and col_pvp:
        df[col_stock_actual] = pd.to_numeric(df[col_stock_actual], errors='coerce').fillna(0)
        
        df['Valor_Stock_Actual'] = df[col_stock_actual] * df[col_pvp]
        df['Valor_Stock_Ideal'] = df['Stock_Ideal'] * df[col_pvp]
        df['Valor_Stock_Limite'] = df['Stock_Limite'] * df[col_pvp]
        
        # CORRECCIÓN: Stock SOBRANTE cuando Stock Actual > Stock IDEAL
        df['Stock_Sobrante_Uds'] = np.where(
            df[col_stock_actual] > df['Stock_Ideal'],
            df[col_stock_actual] - df['Stock_Ideal'],
            0
        )
        df['Stock_Sobrante'] = df['Stock_Sobrante_Uds'] * df[col_pvp]
        
        # CORRECCIÓN: Stock FALTANTE cuando Stock Actual < Stock IDEAL
        df['Stock_Faltante_Uds'] = np.where(
            df[col_stock_actual] < df['Stock_Ideal'],
            df['Stock_Ideal'] - df[col_stock_actual],
            0
        )
        df['Stock_Faltante'] = df['Stock_Faltante_Uds'] * df[col_pvp]
        
        df['Reposicion'] = df['Stock_Ideal'] - df[col_stock_actual]
        
        df['Indice_Rotacion'] = df.apply(
            lambda row: calcular_indice_rotacion(row['Total_Ventas'], row[col_stock_actual]),
            axis=1
        )
        
        df['Valor_Ventas'] = df['Total_Ventas'] * df[col_pvp]
    
    # CORRECCIÓN: Procesar familias funcionales desde prefijos
    if col_categoria_funcional:
        df['Familia'] = df[col_categoria_funcional].apply(extraer_familia)
        df['Subfamilia'] = df[col_categoria_funcional]
    
    return df, col_stock_actual, col_pvp, col_cn, col_descripcion, col_categoria_funcional

# Sidebar
st.sidebar.header("⚙️ Configuración")

# Botón resetear
if st.sidebar.button("🔄 Resetear Filtros y Estados", use_container_width=True):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.rerun()

dias_abierto = st.sidebar.number_input(
    "Días abierto al año", 
    min_value=250, 
    max_value=365, 
    value=300, 
    step=1,
    help="Número de días que la farmacia permanece abierta anualmente"
)

st.sidebar.markdown("### Parámetros Stock (A y B)")
col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    stock_min_dias = st.number_input(
        "Mín (días)", 
        min_value=5, 
        max_value=20, 
        value=10, 
        step=1,
        help="Stock mínimo en días de cobertura"
    )
with col_sb2:
    stock_max_dias = st.number_input(
        "Máx (días)", 
        min_value=15, 
        max_value=40, 
        value=20,  # CORRECCIÓN: Cambiado de 30 a 20
        step=1,
        help="Stock máximo en días de cobertura"
    )

dias_cobertura_optimo = st.sidebar.slider(
    "Días cobertura ideal (A y B)", 
    min_value=10, 
    max_value=30, 
    value=15, 
    step=1,
    help="Días de cobertura objetivo para productos de alta rotación"
)

margen_seguridad = st.sidebar.slider(
    "Margen de seguridad (%)", 
    min_value=0.0, 
    max_value=0.30, 
    value=0.0,  # CORRECCIÓN: Cambiado de 0.15 a 0.0
    step=0.05, 
    help="Margen adicional sobre el stock ideal para establecer el límite. Con 0%, Stock Ideal = Stock Límite"
)

# Info sobre preparación del Excel
with st.sidebar.expander("📋 Preparación del Excel"):
    st.markdown("""
    **El sistema detecta automáticamente las familias** desde la columna de categorías con formato:
    
    ```
    ADELG-BARRITAS SUELTAS
    DERMO-ACNE
    FITO-DIGESTION
    ```
    
    **Familias reconocidas:**
    - ADELGAZANTES, ANTICELULITICOS, AROMATERAPIA
    - DEPORTE, DERMO, DIETETICA, EFECTOS
    - EFP, ESPECIALIDAD, FITOTERAPIA
    - HIG.BUCAL, HIG.CAPILAR, HIG.CORPORAL
    - HOMEOPATIA, INFANTIL, INSECTOS
    - NARIZ OIDOS, OPTICA, ORTOPEDIA
    - PIES/MANOS, SALUD GINECOLOGICA, SALUD SEXUAL
    - SOLARES, VETERINARIA, VACUNAS, FORMULAS
    - ENVASE CLINICO
    
    Si tu Excel no tiene este formato, considera añadir una columna "Familia" con los nombres completos.
    """)

# Upload Excel
uploaded_file = st.file_uploader("📁 Cargar archivo Excel con datos de ventas", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        df, col_stock_actual, col_pvp, col_cn, col_descripcion, col_categoria_funcional = procesar_excel(
            uploaded_file, dias_abierto, stock_min_dias, stock_max_dias, dias_cobertura_optimo, margen_seguridad
        )
        
        st.success(f"✅ Archivo procesado correctamente: {len(df):,} productos analizados".replace(",", "."))
        
        # Resumen ejecutivo
        with st.expander("📊 Resumen Ejecutivo", expanded=True):
            col_res1, col_res2, col_res3, col_res4 = st.columns(4)
            
            total_inversion = df['Valor_Stock_Actual'].sum()
            total_ideal = df['Valor_Stock_Ideal'].sum()
            total_sobrante = df['Stock_Sobrante'].sum()
            total_faltante = df['Stock_Faltante'].sum()
            
            with col_res1:
                st.metric("Inversión Total en Stock", formato_euros(total_inversion))
            with col_res2:
                delta_ideal = total_inversion - total_ideal
                st.metric(
                    "Inversión Ideal Objetivo", 
                    formato_euros(total_ideal),
                    delta=formato_euros(delta_ideal) if delta_ideal != 0 else None
                )
            with col_res3:
                pct_sobrante = (total_sobrante / total_inversion * 100) if total_inversion > 0 else 0
                st.metric(
                    "💰 Exceso de Stock", 
                    formato_euros(total_sobrante),
                    delta=f"{pct_sobrante:.1f}% del total"
                )
            with col_res4:
                pct_faltante = (total_faltante / total_ideal * 100) if total_ideal > 0 else 0
                st.metric(
                    "⚠️ Déficit de Stock", 
                    formato_euros(total_faltante),
                    delta=f"{pct_faltante:.1f}% del ideal"
                )
            
            # Alertas inteligentes
            if total_sobrante > total_faltante:
                st.warning(f"⚠️ **Exceso de inventario detectado**: El exceso supera al déficit en {formato_euros(total_sobrante - total_faltante)}. Se recomienda revisar productos con baja rotación.")
            elif total_faltante > total_sobrante:
                st.info(f"📈 **Oportunidad de optimización**: Déficit de {formato_euros(total_faltante - total_sobrante)}. Considere aumentar stock en productos de alta demanda.")
        
        st.markdown("---")
        
        # ========== GRÁFICO 1: DISTRIBUCIÓN POR CATEGORÍAS DE ROTACIÓN ==========
        st.subheader("📈 Clasificación por Velocidad de Rotación")
        st.caption("Distribución de productos según su frecuencia de venta anual (A: >260 uds, B: 52-260 uds, C: 12-51 uds, D: 1-11 uds, E: <1 ud)")
        
        col_graf1, col_graf2 = st.columns([1, 1])
        
        with col_graf1:
            cat_descripciones = {
                'A': 'Alta rotación (>260 uds/año)',
                'B': 'Rotación media-alta (52-260 uds/año)',
                'C': 'Rotación media (12-51 uds/año)',
                'D': 'Rotación baja (1-11 uds/año)',
                'E': 'Sin rotación (<1 ud/año)'
            }
            
            categorias_ordenadas = ['A', 'B', 'C', 'D', 'E']
            colores_ordenados = ['#667eea', '#764ba2', '#f093fb', '#4facfe', '#43e97b']
            
            categoria_counts = df['Categoria'].value_counts().reindex(categorias_ordenadas, fill_value=0)
            
            hover_text = [f"{cat}: {cat_descripciones[cat]}<br>Productos: {count}" 
                         for cat, count in categoria_counts.items()]
            
            fig_pie = go.Figure(data=[go.Pie(
                labels=categorias_ordenadas,
                values=categoria_counts.values,
                hole=0.4,
                marker=dict(colors=colores_ordenados),
                textinfo='label+percent',
                textfont=dict(size=14),
                hovertext=hover_text,
                hoverinfo='text',
                sort=False
            )])
            
            fig_pie.update_layout(
                title="Proporción de Productos por Categoría",
                height=400,
                showlegend=True,
                legend=dict(traceorder='normal')
            )
            
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col_graf2:
            resumen_cat = df.groupby('Categoria').agg({
                'Total_Ventas': 'sum',
                'Valor_Stock_Actual': 'sum',
                'Stock_Sobrante': 'sum'
            }).reindex(['A', 'B', 'C', 'D', 'E'], fill_value=0).round(2)
            
            resumen_cat_display = pd.DataFrame({
                'Cat.': resumen_cat.index,
                'Ventas Anuales': resumen_cat['Total_Ventas'].apply(formato_numero),
                'Inversión Stock': resumen_cat['Valor_Stock_Actual'].apply(formato_euros),
                'Exceso Stock': resumen_cat['Stock_Sobrante'].apply(formato_euros)
            })
            
            st.dataframe(resumen_cat_display, use_container_width=True, height=250, hide_index=True)
            
            # CORRECCIÓN: Botón Ver/Ocultar CNs con toggle
            productos_sobrantes = df[df['Stock_Sobrante_Uds'] > 0].copy()
            if len(productos_sobrantes) > 0 and col_cn:
                # Estado del botón
                mostrar_cns = st.session_state.get('mostrar_cns_sobrante', False)
                
                # Botón con texto dinámico
                boton_texto = "🔽 Ocultar CNs con Exceso" if mostrar_cns else f"👁️ Ver CNs con Exceso ({len(productos_sobrantes)} productos)"
                
                if st.button(boton_texto, use_container_width=True, key="btn_toggle_cns"):
                    st.session_state.mostrar_cns_sobrante = not mostrar_cns
                    st.rerun()
                
                # Mostrar CNs si está activado
                if mostrar_cns:
                    productos_sobrantes = productos_sobrantes.sort_values('Stock_Sobrante', ascending=False)
                    
                    st.markdown("#### 📋 Productos con Exceso de Stock")
                    cns_display = productos_sobrantes[[
                        col_cn, 
                        col_descripcion if col_descripcion else col_cn,
                        'Categoria',
                        col_stock_actual,
                        'Stock_Ideal',
                        'Stock_Sobrante_Uds',
                        'Stock_Sobrante'
                    ]].head(100)
                    
                    cns_display_formatted = cns_display.copy()
                    cns_display_formatted['Stock_Sobrante'] = cns_display_formatted['Stock_Sobrante'].apply(formato_euros)
                    cns_display_formatted.columns = ['CN', 'Descripción', 'Cat.', 'Stock Actual', 'Stock Ideal', 'Exceso (uds)', 'Valor Exceso']
                    
                    st.dataframe(cns_display_formatted, use_container_width=True, height=400, hide_index=True)
                    
                    # Botones de descarga
                    col_btn1, col_btn2 = st.columns(2)
                    
                    with col_btn1:
                        cns_sobrantes = "\n".join(productos_sobrantes[col_cn].astype(str).tolist())
                        st.download_button(
                            label="📄 Descargar CNs (TXT)",
                            data=cns_sobrantes,
                            file_name=f"CNs_exceso_stock_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                            mime="text/plain",
                            use_container_width=True
                        )
                    
                    with col_btn2:
                        output_cns = BytesIO()
                        with pd.ExcelWriter(output_cns, engine='openpyxl') as writer:
                            productos_sobrantes[[
                                col_cn, 
                                col_descripcion if col_descripcion else col_cn,
                                'Categoria',
                                col_stock_actual,
                                'Stock_Ideal',
                                'Stock_Sobrante_Uds',
                                'Stock_Sobrante'
                            ]].to_excel(writer, sheet_name='Exceso Stock', index=False)
                        output_cns.seek(0)
                        
                        st.download_button(
                            label="📊 Descargar Detalle (Excel)",
                            data=output_cns,
                            file_name=f"detalle_exceso_stock_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
        
        st.markdown("---")
        
        # ========== GRÁFICO 2: DESGLOSE POR DEMANDA ==========
        st.subheader("📊 Análisis Cuantitativo por Categoría de Rotación")
        st.caption("Comparativa de referencias vs stock físico según velocidad de rotación")
        
        if col_cn:
            col_tabla, col_grafico = st.columns([1, 1])
            
            with col_tabla:
                analisis_demanda = df.groupby('Categoria').agg({
                    col_cn: 'count',
                    col_stock_actual: 'sum',
                    'Total_Ventas': 'sum',
                    'Indice_Rotacion': 'mean'
                }).reindex(['A', 'B', 'C', 'D', 'E'], fill_value=0)
                
                total_refs = len(df)
                total_stock = df[col_stock_actual].sum()
                
                analisis_demanda_display = pd.DataFrame({
                    'Cat.': analisis_demanda.index,
                    'Referencias': analisis_demanda[col_cn].astype(int),
                    '% Refs': (analisis_demanda[col_cn] / total_refs * 100).round(1),
                    'Stock (uds)': analisis_demanda[col_stock_actual].round(0).astype(int),
                    '% Stock': (analisis_demanda[col_stock_actual] / total_stock * 100).round(1),
                    'Uds/Ref': (analisis_demanda[col_stock_actual] / analisis_demanda[col_cn].replace(0, 1)).round(1),
                    'IR Medio': analisis_demanda['Indice_Rotacion'].round(2)
                })
                
                st.dataframe(analisis_demanda_display, use_container_width=True, hide_index=True, height=250)
            
            with col_grafico:
                categorias = ['A', 'B', 'C', 'D', 'E']
                pct_refs = (analisis_demanda[col_cn] / total_refs * 100).values
                pct_stock = (analisis_demanda[col_stock_actual] / total_stock * 100).values
                
                fig_demanda = go.Figure()
                fig_demanda.add_trace(go.Bar(
                    name='% Referencias',
                    x=categorias,
                    y=pct_refs,
                    marker_color='#667eea',
                    text=[f'{v:.1f}%' for v in pct_refs],
                    textposition='auto'
                ))
                fig_demanda.add_trace(go.Bar(
                    name='% Stock (uds)',
                    x=categorias,
                    y=pct_stock,
                    marker_color='#764ba2',
                    text=[f'{v:.1f}%' for v in pct_stock],
                    textposition='auto'
                ))
                
                fig_demanda.update_layout(
                    title='Distribución: Referencias vs Stock',
                    barmode='group',
                    yaxis_title='Porcentaje (%)',
                    xaxis_title='Categoría',
                    height=350,
                    showlegend=True
                )
                
                st.plotly_chart(fig_demanda, use_container_width=True)
        
        st.markdown("---")
        
        # ========== GRÁFICO 3: ANÁLISIS POR CATEGORÍAS - COMPARATIVA STOCK ==========
        st.subheader("🎯 Comparativa Stock: Actual vs Ideal vs Límite")
        st.caption(f"Evaluación del inventario actual respecto a objetivos (Margen de seguridad: {margen_seguridad*100:.0f}%)")
        
        if col_cn:
            col_tabla_comp, col_grafico_comp = st.columns([1, 1])
            
            with col_tabla_comp:
                analisis_cat_ideal = df.groupby('Categoria').agg({
                    col_cn: 'count',
                    col_stock_actual: 'sum',
                    'Stock_Ideal': 'sum',
                    'Stock_Limite': 'sum',
                    'Stock_Sobrante_Uds': 'sum',
                    'Stock_Faltante_Uds': 'sum',
                    'Valor_Stock_Actual': 'sum',
                    'Valor_Stock_Limite': 'sum',
                    'Stock_Sobrante': 'sum',
                    'Stock_Faltante': 'sum'
                }).reindex(['A', 'B', 'C', 'D', 'E'], fill_value=0).round(2)
            
                analisis_cat_comp = pd.DataFrame({
                    'Cat.': analisis_cat_ideal.index,
                    'Refs': analisis_cat_ideal[col_cn].astype(int),
                    'Stock Actual': analisis_cat_ideal[col_stock_actual].round(0).astype(int),
                    'Stock Ideal': analisis_cat_ideal['Stock_Ideal'].round(0).astype(int),
                    'Stock Límite': analisis_cat_ideal['Stock_Limite'].round(0).astype(int),
                    'Exceso (uds)': analisis_cat_ideal['Stock_Sobrante_Uds'].round(0).astype(int),
                    'Déficit (uds)': analisis_cat_ideal['Stock_Faltante_Uds'].round(0).astype(int),
                    'Inversión Actual': analisis_cat_ideal['Valor_Stock_Actual'].apply(formato_euros),
                    'Inversión Límite': analisis_cat_ideal['Valor_Stock_Limite'].apply(formato_euros),
                    'Valor Exceso': analisis_cat_ideal['Stock_Sobrante'].apply(formato_euros),
                    'Valor Déficit': analisis_cat_ideal['Stock_Faltante'].apply(formato_euros)
                })
                
                st.dataframe(analisis_cat_comp, use_container_width=True, height=250, hide_index=True)
            
            with col_grafico_comp:
                categorias = ['A', 'B', 'C', 'D', 'E']
                stock_actual = analisis_cat_ideal[col_stock_actual].values
                stock_ideal = analisis_cat_ideal['Stock_Ideal'].values
                stock_limite = analisis_cat_ideal['Stock_Limite'].values
                
                fig_comp = go.Figure()
                fig_comp.add_trace(go.Bar(
                    name='Stock Actual',
                    x=categorias,
                    y=stock_actual,
                    marker_color='#4169E1'
                ))
                fig_comp.add_trace(go.Bar(
                    name='Stock Ideal',
                    x=categorias,
                    y=stock_ideal,
                    marker_color='#FFD700'
                ))
                fig_comp.add_trace(go.Bar(
                    name='Stock Límite',
                    x=categorias,
                    y=stock_limite,
                    marker_color='#FF6347'
                ))
                
                fig_comp.update_layout(
                    title='Comparativa Stock Actual vs Ideal vs Límite',
                    barmode='group',
                    yaxis_title='Unidades',
                    xaxis_title='Categoría',
                    height=350,
                    showlegend=True
                )
                
                st.plotly_chart(fig_comp, use_container_width=True)
        else:
            st.warning("⚠️ No se encontró la columna de código de artículo para realizar el análisis comparativo")
        
        st.markdown("---")
        
        # ========== GRÁFICO 4: ANÁLISIS POR FAMILIAS FUNCIONALES - STOCK ACTUAL ==========
        if 'Familia' in df.columns and col_cn:
            st.subheader("🏪 Análisis por Familias Terapéuticas - Inventario Actual")
            st.caption("Distribución del stock y ventas por categorías funcionales de productos")
            
            familias_todas = sorted(df['Familia'].dropna().unique().tolist())
            
            # Filtro de familias
            familias_selec_stock = st.multiselect(
                "Filtrar por familias específicas (dejar vacío para incluir todas):",
                options=familias_todas,
                default=[],
                key="familias_stock_actual",
                help="Seleccione una o varias familias para análisis específico"
            )
            
            if familias_selec_stock:
                df_analisis_familias = df[df['Familia'].isin(familias_selec_stock)]
            else:
                df_analisis_familias = df.copy()
            
            analisis_familias_actual = df_analisis_familias.groupby('Familia').agg({
                col_cn: 'count',
                col_stock_actual: 'sum',
                'Valor_Stock_Actual': 'sum',
                'Total_Ventas': 'sum',
                'Valor_Ventas': 'sum',
                'Indice_Rotacion': 'mean'
            }).sort_values('Valor_Stock_Actual', ascending=False).round(2)
            
            total_stock_uds = df_analisis_familias[col_stock_actual].sum()
            total_stock_valor = df_analisis_familias['Valor_Stock_Actual'].sum()
            total_ventas_uds = df_analisis_familias['Total_Ventas'].sum()
            total_ventas_valor = df_analisis_familias['Valor_Ventas'].sum()
            
            analisis_familias_display = pd.DataFrame({
                'Familia': analisis_familias_actual.index,
                'Nº Refs': analisis_familias_actual[col_cn].astype(int),
                'Stock (uds)': analisis_familias_actual[col_stock_actual].round(0).astype(int),
                '% Stock (uds)': (analisis_familias_actual[col_stock_actual] / total_stock_uds * 100).round(1),
                'Inversión (€)': analisis_familias_actual['Valor_Stock_Actual'].apply(formato_euros),
                '% Inversión': (analisis_familias_actual['Valor_Stock_Actual'] / total_stock_valor * 100).round(1),
                'Ventas (uds)': analisis_familias_actual['Total_Ventas'].round(0).astype(int),
                '% Ventas (uds)': (analisis_familias_actual['Total_Ventas'] / total_ventas_uds * 100).round(1),
                'Ventas (€)': analisis_familias_actual['Valor_Ventas'].apply(formato_euros),
                '% Ventas (€)': (analisis_familias_actual['Valor_Ventas'] / total_ventas_valor * 100).round(1),
                'Uds/Ref': (analisis_familias_actual[col_stock_actual] / analisis_familias_actual[col_cn]).round(1),
                'IR Medio': analisis_familias_actual['Indice_Rotacion'].round(2)
            })
            
            st.dataframe(analisis_familias_display, use_container_width=True, height=400, hide_index=True)
            
            st.markdown("---")
            
            # ========== GRÁFICO 5: FAMILIAS FUNCIONALES CON MAYOR EXCESO ==========
            st.subheader("🚨 Familias Terapéuticas con Mayor Exceso de Stock")
            st.caption("Top 15 familias con mayor inversión en exceso de inventario respecto al stock ideal")
            
            familias_unicas = sorted(df['Familia'].dropna().unique().tolist())
            familias_selec_grafico = st.multiselect(
                "Filtrar familias para análisis de exceso (dejar vacío para todas):",
                options=familias_unicas,
                default=[],
                key="familias_sobrestock",
                help="Útil para enfocarse en categorías específicas"
            )
            
            if familias_selec_grafico:
                df_filtrado_familias = df[df['Familia'].isin(familias_selec_grafico)]
            else:
                df_filtrado_familias = df.copy()
            
            top_sobrestock = df_filtrado_familias.groupby('Familia')['Stock_Sobrante'].sum().sort_values(ascending=False).head(15)
            
            if len(top_sobrestock) > 0 and top_sobrestock.sum() > 0:
                fig_sobrestock = go.Figure(data=[
                    go.Bar(
                        x=top_sobrestock.values,
                        y=top_sobrestock.index,
                        orientation='h',
                        marker=dict(
                            color=top_sobrestock.values,
                            colorscale='Reds',
                            showscale=True,
                            colorbar=dict(title="Valor (€)")
                        ),
                        text=[formato_euros(v) for v in top_sobrestock.values],
                        textposition='auto',
                        hovertemplate='<b>%{y}</b><br>Exceso: %{x:,.2f}€<extra></extra>'
                    )
                ])
                
                fig_sobrestock.update_layout(
                    title="Top 15 Familias con Mayor Exceso de Stock",
                    xaxis_title="Valor Exceso de Stock (€)",
                    yaxis_title="Familia Terapéutica",
                    height=500,
                    showlegend=False
                )
                
                st.plotly_chart(fig_sobrestock, use_container_width=True)
            else:
                st.info("✅ No se detecta exceso de stock significativo en las familias seleccionadas")
            
            st.markdown("---")
            
            # ========== GRÁFICO 6: DESFASE STOCK POR FAMILIAS ==========
            st.subheader("📊 Desfase de Inventario por Familias Terapéuticas")
            st.caption("Balance neto entre exceso y déficit de stock por familia (Rojo=Exceso, Azul=Déficit)")
            
            analisis_desfase = df.groupby('Familia').agg({
                col_cn: 'count',
                col_stock_actual: 'sum',
                'Stock_Ideal': 'sum',
                'Stock_Limite': 'sum',
                'Stock_Sobrante_Uds': 'sum',
                'Stock_Faltante_Uds': 'sum',
                'Valor_Stock_Actual': 'sum',
                'Valor_Stock_Ideal': 'sum',
                'Stock_Sobrante': 'sum',
                'Stock_Faltante': 'sum'
            }).round(2)
            
            analisis_desfase['Desfase_Uds'] = analisis_desfase['Stock_Sobrante_Uds'] - analisis_desfase['Stock_Faltante_Uds']
            analisis_desfase['Desfase_Valor'] = analisis_desfase['Stock_Sobrante'] - analisis_desfase['Stock_Faltante']
            analisis_desfase['Desfase_Abs'] = analisis_desfase['Desfase_Valor'].abs()
            
            top_desfase = analisis_desfase.nlargest(20, 'Desfase_Abs')
            
            if len(top_desfase) > 0:
                col_desfase_graf, col_desfase_tabla = st.columns([1, 1])
                
                with col_desfase_graf:
                    colores = ['#FF6347' if v > 0 else '#4169E1' for v in top_desfase['Desfase_Valor']]
                    
                    fig_desfase = go.Figure(data=[
                        go.Bar(
                            x=top_desfase['Desfase_Valor'].values,
                            y=top_desfase.index,
                            orientation='h',
                            marker=dict(color=colores),
                            text=[formato_euros(abs(v)) for v in top_desfase['Desfase_Valor'].values],
                            textposition='auto',
                            hovertemplate='<b>%{y}</b><br>Desfase: %{x:,.2f}€<extra></extra>'
                        )
                    ])
                    
                    fig_desfase.update_layout(
                        title="Top 20 Familias: Desfase Neto de Stock",
                        xaxis_title="Desfase Neto (€) - Exceso(+) / Déficit(-)",
                        yaxis_title="Familia Terapéutica",
                        height=600,
                        showlegend=False
                    )
                    
                    st.plotly_chart(fig_desfase, use_container_width=True)
                
                with col_desfase_tabla:
                    desfase_display = pd.DataFrame({
                        'Familia': top_desfase.index,
                        'Stock Actual': top_desfase[col_stock_actual].round(0).astype(int),
                        'Stock Ideal': top_desfase['Stock_Ideal'].round(0).astype(int),
                        'Desfase (uds)': top_desfase['Desfase_Uds'].round(0).astype(int),
                        'Desfase (€)': top_desfase['Desfase_Valor'].apply(formato_euros),
                        'Exceso': top_desfase['Stock_Sobrante'].apply(formato_euros),
                        'Déficit': top_desfase['Stock_Faltante'].apply(formato_euros)
                    })
                    
                    st.dataframe(desfase_display, use_container_width=True, height=600, hide_index=True)
            
            st.markdown("---")
            
            # ========== GRÁFICO 7: DESGLOSE POR FAMILIA Y SUBFAMILIAS ==========
            st.subheader("🔍 Análisis Detallado: Familia y Subfamilias")
            st.caption("Exploración profunda de una familia específica y su desglose por subfamilias")
            
            familias_disponibles = sorted(df['Familia'].dropna().unique().tolist())
            familia_seleccionada = st.selectbox(
                "Seleccione una familia terapéutica para análisis detallado:",
                familias_disponibles,
                help="Permite visualizar subfamilias y productos dentro de una categoría específica"
            )
            
            if familia_seleccionada:
                df_familia = df[df['Familia'] == familia_seleccionada]
                
                subfamilias_unicas = df_familia['Subfamilia'].unique()
                tiene_subfamilias = len(subfamilias_unicas) > 1 or (len(subfamilias_unicas) == 1 and subfamilias_unicas[0] != familia_seleccionada)
                
                if tiene_subfamilias:
                    st.markdown(f"### 📋 Desglose de Subfamilias: {familia_seleccionada}")
                    
                    analisis_subfamilias = df_familia.groupby('Subfamilia').agg({
                        col_cn: 'count',
                        col_stock_actual: 'sum',
                        'Stock_Ideal': 'sum',
                        'Stock_Limite': 'sum',
                        'Stock_Sobrante_Uds': 'sum',
                        'Stock_Faltante_Uds': 'sum',
                        'Valor_Stock_Actual': 'sum',
                        'Valor_Stock_Ideal': 'sum',
                        'Valor_Stock_Limite': 'sum',
                        'Stock_Sobrante': 'sum',
                        'Stock_Faltante': 'sum',
                        'Total_Ventas': 'sum',
                        'Valor_Ventas': 'sum',
                        'Indice_Rotacion': 'mean'
                    }).sort_values('Valor_Stock_Actual', ascending=False).round(2)
                    
                    analisis_subfamilias_display = pd.DataFrame({
                        'Subfamilia': analisis_subfamilias.index,
                        'Nº Refs': analisis_subfamilias[col_cn].astype(int),
                        'Stock Actual': analisis_subfamilias[col_stock_actual].round(0).astype(int),
                        'Stock Ideal': analisis_subfamilias['Stock_Ideal'].round(0).astype(int),
                        'Stock Límite': analisis_subfamilias['Stock_Limite'].round(0).astype(int),
                        'Exceso (uds)': analisis_subfamilias['Stock_Sobrante_Uds'].round(0).astype(int),
                        'Déficit (uds)': analisis_subfamilias['Stock_Faltante_Uds'].round(0).astype(int),
                        'Inversión Actual': analisis_subfamilias['Valor_Stock_Actual'].apply(formato_euros),
                        'Inversión Ideal': analisis_subfamilias['Valor_Stock_Ideal'].apply(formato_euros),
                        'Valor Exceso': analisis_subfamilias['Stock_Sobrante'].apply(formato_euros),
                        'Valor Déficit': analisis_subfamilias['Stock_Faltante'].apply(formato_euros),
                        'Ventas (uds)': analisis_subfamilias['Total_Ventas'].round(0).astype(int),
                        'Ventas (€)': analisis_subfamilias['Valor_Ventas'].apply(formato_euros),
                        'IR Medio': analisis_subfamilias['Indice_Rotacion'].round(2)
                    })
                    
                    st.dataframe(analisis_subfamilias_display, use_container_width=True, height=400, hide_index=True)
                    
                    col_graf_sub1, col_graf_sub2 = st.columns(2)
                    
                    with col_graf_sub1:
                        top_10_subfamilias = analisis_subfamilias.nlargest(10, 'Valor_Stock_Actual')
                        
                        fig_subfam_comp = go.Figure()
                        fig_subfam_comp.add_trace(go.Bar(
                            name='Stock Actual',
                            x=top_10_subfamilias.index,
                            y=top_10_subfamilias[col_stock_actual],
                            marker_color='#4169E1'
                        ))
                        fig_subfam_comp.add_trace(go.Bar(
                            name='Stock Ideal',
                            x=top_10_subfamilias.index,
                            y=top_10_subfamilias['Stock_Ideal'],
                            marker_color='#32CD32'
                        ))
                        fig_subfam_comp.add_trace(go.Bar(
                            name='Stock Límite',
                            x=top_10_subfamilias.index,
                            y=top_10_subfamilias['Stock_Limite'],
                            marker_color='#FFA500'
                        ))
                        
                        fig_subfam_comp.update_layout(
                            title='Top 10 Subfamilias: Comparativa de Stock',
                            barmode='group',
                            yaxis_title='Unidades',
                            xaxis_title='Subfamilia',
                            height=400,
                            showlegend=True,
                            xaxis={'tickangle': -45}
                        )
                        
                        st.plotly_chart(fig_subfam_comp, use_container_width=True)
                    
                    with col_graf_sub2:
                        top_sobrestock_subfam = analisis_subfamilias.nlargest(10, 'Stock_Sobrante')
                        
                        if len(top_sobrestock_subfam) > 0 and top_sobrestock_subfam['Stock_Sobrante'].sum() > 0:
                            fig_sobr_subfam = go.Figure(data=[
                                go.Bar(
                                    x=top_sobrestock_subfam['Stock_Sobrante'].values,
                                    y=top_sobrestock_subfam.index,
                                    orientation='h',
                                    marker=dict(
                                        color=top_sobrestock_subfam['Stock_Sobrante'].values,
                                        colorscale='Reds'
                                    ),
                                    text=[formato_euros(v) for v in top_sobrestock_subfam['Stock_Sobrante'].values],
                                    textposition='auto'
                                )
                            ])
                            
                            fig_sobr_subfam.update_layout(
                                title='Top 10 Subfamilias: Exceso de Stock',
                                xaxis_title='Valor Exceso (€)',
                                yaxis_title='Subfamilia',
                                height=400,
                                showlegend=False
                            )
                            
                            st.plotly_chart(fig_sobr_subfam, use_container_width=True)
                        else:
                            st.info("✅ No hay exceso de stock en las subfamilias de esta familia")
                
                else:
                    st.markdown(f"### 📋 Análisis de: {familia_seleccionada}")
                    st.info("ℹ️ Esta familia no tiene subfamilias diferenciadas en el sistema")
                    
                    resumen_familia = df_familia.agg({
                        col_cn: 'count',
                        col_stock_actual: 'sum',
                        'Stock_Ideal': 'sum',
                        'Stock_Limite': 'sum',
                        'Stock_Sobrante_Uds': 'sum',
                        'Stock_Faltante_Uds': 'sum',
                        'Valor_Stock_Actual': 'sum',
                        'Valor_Stock_Ideal': 'sum',
                        'Valor_Stock_Limite': 'sum',
                        'Stock_Sobrante': 'sum',
                        'Stock_Faltante': 'sum',
                        'Total_Ventas': 'sum',
                        'Valor_Ventas': 'sum',
                        'Indice_Rotacion': 'mean'
                    }).round(2)
                    
                    col_res1, col_res2, col_res3, col_res4 = st.columns(4)
                    with col_res1:
                        st.metric("Nº Referencias", resumen_familia[col_cn].astype(int))
                    with col_res2:
                        st.metric("Inversión Stock", formato_euros(resumen_familia['Valor_Stock_Actual']))
                    with col_res3:
                        st.metric("Exceso Stock", formato_euros(resumen_familia['Stock_Sobrante']))
                    with col_res4:
                        st.metric("Índice Rotación", f"{resumen_familia['Indice_Rotacion']:.2f}")
                    
                    st.markdown("#### Distribución por Categorías de Rotación")
                    
                    dist_cat_familia = df_familia.groupby('Categoria').agg({
                        col_cn: 'count',
                        col_stock_actual: 'sum',
                        'Valor_Stock_Actual': 'sum',
                        'Stock_Sobrante': 'sum'
                    }).reindex(['A', 'B', 'C', 'D', 'E'], fill_value=0).round(2)
                    
                    dist_cat_display = pd.DataFrame({
                        'Categoría': dist_cat_familia.index,
                        'Nº Refs': dist_cat_familia[col_cn].astype(int),
                        'Stock (uds)': dist_cat_familia[col_stock_actual].round(0).astype(int),
                        'Inversión': dist_cat_familia['Valor_Stock_Actual'].apply(formato_euros),
                        'Exceso': dist_cat_familia['Stock_Sobrante'].apply(formato_euros)
                    })
                    
                    st.dataframe(dist_cat_display, use_container_width=True, hide_index=True)
        else:
            st.info("ℹ️ No se detectaron familias funcionales en el archivo. Asegúrese de que existe una columna con categorías en formato: PREFIJO-DESCRIPCIÓN")
        
        # ========== EXPORTAR DATOS ==========
        st.markdown("---")
        st.subheader("📥 Exportación de Informes")
        st.caption("Descargue los análisis completos en formato Excel para archivo o análisis adicional")
        
        col_exp1, col_exp2, col_exp3 = st.columns(3)
        
        with col_exp1:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Datos Completos', index=False)
                
                if col_cn:
                    resumen_cat = df.groupby('Categoria').agg({
                        col_cn: 'count',
                        col_stock_actual: 'sum',
                        'Stock_Ideal': 'sum',
                        'Stock_Sobrante': 'sum',
                        'Stock_Faltante': 'sum',
                        'Valor_Stock_Actual': 'sum',
                        'Total_Ventas': 'sum'
                    }).round(2)
                    resumen_cat.to_excel(writer, sheet_name='Resumen Categorías')
                
                if 'Familia' in df.columns:
                    resumen_fam = df.groupby('Familia').agg({
                        col_cn: 'count',
                        col_stock_actual: 'sum',
                        'Stock_Ideal': 'sum',
                        'Stock_Limite': 'sum',
                        'Stock_Sobrante': 'sum',
                        'Stock_Faltante': 'sum',
                        'Valor_Stock_Actual': 'sum',
                        'Total_Ventas': 'sum'
                    }).round(2)
                    resumen_fam.to_excel(writer, sheet_name='Resumen Familias')
            
            output.seek(0)
            st.download_button(
                label="📊 Análisis Completo",
                data=output,
                file_name=f"analisis_stock_completo_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Incluye todos los datos procesados, resúmenes por categoría y familia"
            )
        
        with col_exp2:
            if len(productos_sobrantes) > 0:
                output_sobr = BytesIO()
                productos_sobrantes_export = productos_sobrantes[[
                    col_cn, col_descripcion if col_descripcion else col_cn, 
                    'Categoria', 'Familia' if 'Familia' in df.columns else col_cn,
                    col_stock_actual, 'Stock_Ideal', 'Stock_Sobrante_Uds', 'Stock_Sobrante'
                ]].copy()
                
                with pd.ExcelWriter(output_sobr, engine='openpyxl') as writer:
                    productos_sobrantes_export.to_excel(writer, sheet_name='Exceso Stock', index=False)
                
                output_sobr.seek(0)
                st.download_button(
                    label="⚠️ Exceso de Stock",
                    data=output_sobr,
                    file_name=f"exceso_stock_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="Listado de productos con stock por encima del nivel ideal"
                )
        
        with col_exp3:
            productos_faltantes = df[df['Stock_Faltante_Uds'] > 0].copy()
            if len(productos_faltantes) > 0:
                productos_faltantes = productos_faltantes.sort_values('Stock_Faltante', ascending=False)
                output_falt = BytesIO()
                productos_faltantes_export = productos_faltantes[[
                    col_cn, col_descripcion if col_descripcion else col_cn,
                    'Categoria', 'Familia' if 'Familia' in df.columns else col_cn,
                    col_stock_actual, 'Stock_Ideal', 'Stock_Faltante_Uds', 'Stock_Faltante'
                ]].copy()
                
                with pd.ExcelWriter(output_falt, engine='openpyxl') as writer:
                    productos_faltantes_export.to_excel(writer, sheet_name='Déficit Stock', index=False)
                
                output_falt.seek(0)
                st.download_button(
                    label="📈 Déficit de Stock",
                    data=output_falt,
                    file_name=f"deficit_stock_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    help="Listado de productos con stock por debajo del nivel ideal"
                )
        
        # ========== METODOLOGÍA ==========
        st.markdown("---")
        with st.expander("📖 Metodología y Criterios de Análisis"):
            st.markdown("""
            ### 🎯 Criterios de Clasificación por Rotación
            
            | Categoría | Ventas Anuales | Descripción |
            |-----------|----------------|-------------|
            | **A** | > 260 unidades | Alta rotación - Productos estrella con venta casi diaria |
            | **B** | 52-260 unidades | Rotación media-alta - Venta semanal o mayor |
            | **C** | 12-51 unidades | Rotación media - Venta mensual aproximadamente |
            | **D** | 1-11 unidades | Baja rotación - Venta ocasional durante el año |
            | **E** | < 1 unidad | Sin rotación - Productos sin ventas significativas |
            
            ### 📊 Cálculo de Stocks
            
            #### Para productos A y B (alta/media-alta rotación):
            - **Stock Mínimo**: Ventas diarias × Días mínimos configurados ({stock_min_dias} días por defecto)
            - **Stock Ideal**: Ventas diarias × Días de cobertura óptima ({dias_cobertura_optimo} días por defecto)
            - **Stock Límite**: Stock Ideal × (1 + Margen de seguridad)
            
            #### Para productos C (rotación media):
            - **Stock Mínimo**: 1 unidad
            - **Stock Ideal**: 1 unidad
            - **Stock Límite**: 1 unidad × (1 + Margen de seguridad)
            
            #### Para productos D (baja rotación):
            - **Stock Mínimo**: 0 unidades
            - **Stock Ideal**: 1 unidad
            - **Stock Límite**: 1 unidad × (1 + Margen de seguridad)
            
            #### Para productos E (sin rotación):
            - **Stock Mínimo**: 0 unidades
            - **Stock Ideal**: 0 unidades
            - **Stock Límite**: 0 unidades
            
            ### 💰 Cálculo de Excesos y Déficits
            
            - **Exceso de Stock**: Se produce cuando Stock Actual > Stock Ideal
              - Fórmula: Stock Actual - Stock Ideal
              - Indica inversión en inventario por encima del nivel óptimo
            
            - **Déficit de Stock**: Se produce cuando Stock Actual < Stock Ideal
              - Fórmula: Stock Ideal - Stock Actual
              - Indica oportunidades de venta perdidas por falta de producto
            
            ### 🔄 Índice de Rotación (IR)
            
            - **Fórmula**: Ventas Anuales / Stock Actual
            - **Interpretación**:
              - IR > 12: Excelente rotación (más de una vez al mes)
              - IR 4-12: Rotación adecuada
              - IR < 4: Rotación baja, revisar stock
            
            ### 🏥 Familias Terapéuticas
            
            El sistema reconoce automáticamente 31 familias funcionales desde el prefijo de la categoría:
            
            **Familias disponibles**: ADELGAZANTES, ANTICELULITICOS, AROMATERAPIA, DEPORTE, DERMO, 
            DIET SOE, DIETETICA, EFEC SOE, EFECTOS, EFP, ESPECIALIDAD, FITOTERAPIA, HIG.BUCAL, 
            HIG.CAPILAR, HIG.CORPORAL, HOMEOPATIA, INFANTIL, INFANTIL SOE, INSECTOS, NARIZ OIDOS, 
            OPTICA, ORTOPEDIA, ORTOPEDIA SOE, PIES/MANOS, SALUD GINECOLOGICA, SALUD SEXUAL, SOLARES, 
            VETERINARIA, VACUNAS, FORMULAS, ENVASE CLINICO
            
            ### ⚙️ Parámetros Configurables
            
            - **Días abierto al año**: Días laborables de la farmacia ({dias_abierto} configurado)
            - **Stock mínimo (A/B)**: Cobertura mínima en días ({stock_min_dias} configurado)
            - **Stock máximo (A/B)**: Cobertura máxima en días ({stock_max_dias} configurado)
            - **Días cobertura ideal**: Objetivo de stock para alta rotación ({dias_cobertura_optimo} configurado)
            - **Margen de seguridad**: Porcentaje adicional sobre stock ideal ({margen_seguridad*100:.0f}% configurado)
            
            ### 📌 Notas Importantes
            
            - Los cálculos se basan en datos históricos de ventas del período analizado
            - Se recomienda revisar periódicamente los parámetros según estacionalidad
            - Los productos nuevos pueden requerir análisis manual adicional
            - Considere promociones y campañas al interpretar excesos temporales
            """)
    
    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {str(e)}")
        st.exception(e)

else:
    st.info("👆 Por favor, cargue un archivo Excel para iniciar el análisis de stock farmacéutico")
    
    st.markdown("""
    ### 📋 Requisitos del Archivo Excel
    
    El archivo debe contener las siguientes columnas:
    
    #### Columnas Obligatorias:
    - **Ventas**: Columnas mensuales de ventas O una columna "TOTAL" con ventas anuales
    - **Stock Actual**: Inventario actual de cada producto
    - **PVP**: Precio de venta al público
    - **CN**: Código Nacional del producto
    
    #### Columnas Opcionales (recomendadas):
    - **Descripción**: Nombre descriptivo del producto
    - **Categoría Funcional**: En formato "PREFIJO-SUBFAMILIA" (ej: DERMO-ACNE, FITO-DIGESTION)
      - Si existe, se detectarán automáticamente las 31 familias terapéuticas
    
    ### 🎯 ¿Qué obtendrás?
    
    1. **Clasificación inteligente** de productos por velocidad de rotación (A, B, C, D, E)
    2. **Cálculo automático** de stocks ideal, mínimo y límite
    3. **Identificación** de excesos y déficits de inventario
    4. **Análisis financiero** del capital invertido en stock
    5. **Desglose por familias terapéuticas** (si está disponible la información)
    6. **Índices de rotación** para evaluar la eficiencia del inventario
    7. **Informes exportables** en Excel para seguimiento y toma de decisiones
    
    ### 💡 Consejo Profesional
    
    Para un análisis óptimo, asegúrese de que:
    - Los datos de ventas correspondan a un período completo (12 meses idealmente)
    - El stock actual esté actualizado a la fecha del análisis
    - Los precios (PVP) estén vigentes
    - Las categorías funcionales sigan el formato estándar de su sistema de gestión
    
    ### 🔧 Configuración Personalizable
    
    Una vez cargado el archivo, podrá ajustar:
    - Días de apertura anual de la farmacia
    - Niveles de stock mínimo y máximo para productos de alta rotación
    - Días de cobertura ideal según su estrategia comercial
    - Margen de seguridad para calcular el stock límite
    
    ---
    
    **📞 ¿Necesita ayuda?** Revise la sección "Metodología y Criterios" al final del análisis 
    para entender los cálculos y criterios utilizados.
    """)

# ========== PIE DE PÁGINA ==========
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <p><strong>Sistema de Análisis de Stock Farmacéutico</strong></p>
    <p style='font-size: 0.9em;'>Versión 2.0 | Optimización de Inventario Basada en Rotación y Demanda</p>
    <p style='font-size: 0.8em;'>Desarrollado para profesionales farmacéuticos | © 2025</p>
</div>
""", unsafe_allow_html=True)
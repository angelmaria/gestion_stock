# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

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
    
    # Calcular stocks
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
            stock_limite = 2
        elif cat == 'D':
            stock_ideal = 1
            stock_min = 0
            stock_limite = 1
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
        
        # CORREGIDO: Stock SOBRANTE cuando Stock Actual > Stock LÍMITE
        df['Stock_Sobrante_Uds'] = np.where(
            df[col_stock_actual] > df['Stock_Limite'],
            df[col_stock_actual] - df['Stock_Limite'],
            0
        )
        df['Stock_Sobrante'] = df['Stock_Sobrante_Uds'] * df[col_pvp]
        
        # CORREGIDO: Stock FALTANTE cuando Stock Actual < Stock LÍMITE
        df['Stock_Faltante_Uds'] = np.where(
            df[col_stock_actual] < df['Stock_Limite'],
            df['Stock_Limite'] - df[col_stock_actual],
            0
        )
        df['Stock_Faltante'] = df['Stock_Faltante_Uds'] * df[col_pvp]
        
        df['Reposicion'] = df['Stock_Ideal'] - df[col_stock_actual]
        
        df['Indice_Rotacion'] = df.apply(
            lambda row: calcular_indice_rotacion(row['Total_Ventas'], row[col_stock_actual]),
            axis=1
        )
        
        df['Valor_Ventas'] = df['Total_Ventas'] * df[col_pvp]
    
    # Procesar categoría funcional (FAMILIAS)
    if col_categoria_funcional:
        df['Familia'] = df[col_categoria_funcional].astype(str).str.split('-').str[0]
        df['Subfamilia'] = df[col_categoria_funcional]
    
    return df, col_stock_actual, col_pvp, col_cn, col_descripcion, col_categoria_funcional

# Sidebar
st.sidebar.header("⚙️ Configuración")
dias_abierto = st.sidebar.number_input("Días abierto al año", min_value=250, max_value=365, value=300, step=1)

st.sidebar.markdown("### Parámetros Stock (A y B)")
col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    stock_min_dias = st.number_input("Mín (días)", min_value=5, max_value=20, value=10, step=1)
with col_sb2:
    stock_max_dias = st.number_input("Máx (días)", min_value=15, max_value=40, value=30, step=1)

dias_cobertura_optimo = st.sidebar.slider("Días cobertura ideal (A y B)", min_value=10, max_value=30, value=15, step=1)

margen_seguridad = st.sidebar.slider("Margen de seguridad (%)", min_value=0.0, max_value=0.30, value=0.15, step=0.05, 
                                      help="Stock límite = Stock ideal × (1 + margen)")

# Upload Excel
uploaded_file = st.file_uploader("📁 Cargar archivo Excel con datos de ventas", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        df, col_stock_actual, col_pvp, col_cn, col_descripcion, col_categoria_funcional = procesar_excel(
            uploaded_file, dias_abierto, stock_min_dias, stock_max_dias, dias_cobertura_optimo, margen_seguridad
        )
        
        st.success(f"✅ Archivo procesado correctamente: {len(df):,} productos".replace(",", "."))
        
        # ========== KPIs PRINCIPALES ==========
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Productos", f"{len(df):,}".replace(",", "."))
        with col2:
            if 'Valor_Stock_Actual' in df.columns:
                st.metric("Valor Stock Actual", formato_euros(df['Valor_Stock_Actual'].sum()))
        with col3:
            if 'Stock_Sobrante' in df.columns:
                st.metric("💰 Stock Sobrante", formato_euros(df['Stock_Sobrante'].sum()))
        with col4:
            if 'Stock_Faltante' in df.columns:
                st.metric("⚠️ Stock Faltante", formato_euros(df['Stock_Faltante'].sum()))
        
        st.markdown("---")
        
        # ========== GRÁFICO 1: DISTRIBUCIÓN POR CATEGORÍAS DE ROTACIÓN ==========
        st.subheader("📈 Distribución por Categorías de Rotación (A, B, C, D, E)")
        
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
                'Total Ventas (uds)': resumen_cat['Total_Ventas'].apply(formato_numero),
                'Valor Stock': resumen_cat['Valor_Stock_Actual'].apply(formato_euros),
                'Stock Sobrante': resumen_cat['Stock_Sobrante'].apply(formato_euros)
            })
            
            st.dataframe(resumen_cat_display, use_container_width=True, height=250, hide_index=True)
            
            # MEJORADO: Botón para ver CNs de stock sobrante
            productos_sobrantes = df[df['Stock_Sobrante_Uds'] > 0].copy()
            if len(productos_sobrantes) > 0 and col_cn:
                if st.button(f"👁️ Ver CNs Stock Sobrante ({len(productos_sobrantes)} productos)", use_container_width=True):
                    st.session_state.mostrar_cns_sobrante = True
                
                # Mostrar CNs si se ha hecho clic
                if st.session_state.get('mostrar_cns_sobrante', False):
                    productos_sobrantes = productos_sobrantes.sort_values('Stock_Sobrante', ascending=False)
                    
                    # Mostrar tabla con info adicional
                    st.markdown("#### 📋 Productos con Stock Sobrante")
                    cns_display = productos_sobrantes[[
                        col_cn, 
                        col_descripcion if col_descripcion else col_cn,
                        'Categoria',
                        col_stock_actual,
                        'Stock_Limite',
                        'Stock_Sobrante_Uds',
                        'Stock_Sobrante'
                    ]].head(50)
                    
                    cns_display_formatted = cns_display.copy()
                    cns_display_formatted['Stock_Sobrante'] = cns_display_formatted['Stock_Sobrante'].apply(formato_euros)
                    cns_display_formatted.columns = ['CN', 'Descripción', 'Cat.', 'Stock Actual', 'Stock Límite', 'Sobrante (uds)', 'Valor Sobrante']
                    
                    st.dataframe(cns_display_formatted, use_container_width=True, height=400, hide_index=True)
                    
                    # Botones de descarga
                    col_btn1, col_btn2 = st.columns(2)
                    
                    with col_btn1:
                        # Descargar solo CNs en TXT
                        cns_sobrantes = "\n".join(productos_sobrantes[col_cn].astype(str).tolist())
                        st.download_button(
                            label="📄 Descargar CNs (TXT)",
                            data=cns_sobrantes,
                            file_name=f"CNs_stock_sobrante_{pd.Timestamp.now().strftime('%Y%m%d')}.txt",
                            mime="text/plain",
                            use_container_width=True
                        )
                    
                    with col_btn2:
                        # Descargar tabla completa en Excel
                        output_cns = BytesIO()
                        with pd.ExcelWriter(output_cns, engine='openpyxl') as writer:
                            productos_sobrantes[[
                                col_cn, 
                                col_descripcion if col_descripcion else col_cn,
                                'Categoria',
                                col_stock_actual,
                                'Stock_Limite',
                                'Stock_Sobrante_Uds',
                                'Stock_Sobrante'
                            ]].to_excel(writer, sheet_name='Stock Sobrante', index=False)
                        output_cns.seek(0)
                        
                        st.download_button(
                            label="📊 Descargar Detalle (Excel)",
                            data=output_cns,
                            file_name=f"detalle_stock_sobrante_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
        
        # ========== GRÁFICO 2: DESGLOSE POR DEMANDA ==========
        st.subheader("📊 Análisis por Tipo de Demanda (Categorías A-E)")
        
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
                    'Refs': analisis_demanda[col_cn].astype(int),
                    '% Refs': (analisis_demanda[col_cn] / total_refs * 100).round(1),
                    'Stock (uds)': analisis_demanda[col_stock_actual].round(0).astype(int),
                    '% Stock': (analisis_demanda[col_stock_actual] / total_stock * 100).round(1),
                    'Media uds/ref': (analisis_demanda[col_stock_actual] / analisis_demanda[col_cn].replace(0, 1)).round(1),
                    'IR': analisis_demanda['Indice_Rotacion'].round(2)
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
        else:
            st.warning("⚠️ No se encontró la columna de código de artículo (CN/IdArticu) para realizar el análisis por demanda")
        
        st.markdown("---")
        
        # ========== GRÁFICO 3: ANÁLISIS POR CATEGORÍAS - COMPARATIVA STOCK IDEAL ==========
        st.subheader("🎯 Análisis por Categorías de Rotación - Comparativa Stock Ideal vs Actual")
        
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
                    'Sobrante': analisis_cat_ideal['Stock_Sobrante_Uds'].round(0).astype(int),
                    'Faltante': analisis_cat_ideal['Stock_Faltante_Uds'].round(0).astype(int),
                    'Valor Actual': analisis_cat_ideal['Valor_Stock_Actual'].apply(formato_euros),
                    'Valor Límite': analisis_cat_ideal['Valor_Stock_Limite'].apply(formato_euros),
                    'Valor Sobrante': analisis_cat_ideal['Stock_Sobrante'].apply(formato_euros),
                    'Valor Faltante': analisis_cat_ideal['Stock_Faltante'].apply(formato_euros)
                })
                
                st.dataframe(analisis_cat_comp, use_container_width=True, height=250, hide_index=True)
            
            with col_grafico_comp:
                categorias = ['A', 'B', 'C', 'D', 'E']
                stock_actual = analisis_cat_ideal[col_stock_actual].values
                stock_ideal = analisis_cat_ideal['Stock_Ideal'].values
                stock_limite = analisis_cat_ideal['Stock_Limite'].values
                
                # COLORES MEJORADOS
                fig_comp = go.Figure()
                fig_comp.add_trace(go.Bar(
                    name='Stock Actual',
                    x=categorias,
                    y=stock_actual,
                    marker_color='#4169E1'  # Azul royal - actual
                ))
                fig_comp.add_trace(go.Bar(
                    name='Stock Ideal',
                    x=categorias,
                    y=stock_ideal,
                    marker_color='#FFD700'  # Dorado - objetivo ideal
                ))
                fig_comp.add_trace(go.Bar(
                    name='Stock Límite',
                    x=categorias,
                    y=stock_limite,
                    marker_color='#FF6347'  # Rojo tomate - límite de alerta
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
        if col_categoria_funcional and 'Familia' in df.columns and col_cn:
            st.subheader("🏪 Análisis por Familias Funcionales - Stock Actual")
            
            familias_todas = sorted(df['Familia'].dropna().unique().tolist())
            familias_selec_stock = st.multiselect(
                "Seleccionar familias para análisis (dejar vacío para ver todas):",
                options=familias_todas,
                default=[],
                key="familias_stock_actual"
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
            }).round(2)
            
            total_stock_uds = df_analisis_familias[col_stock_actual].sum()
            total_stock_valor = df_analisis_familias['Valor_Stock_Actual'].sum()
            total_ventas_uds = df_analisis_familias['Total_Ventas'].sum()
            total_ventas_valor = df_analisis_familias['Valor_Ventas'].sum()
            
            analisis_familias_display = pd.DataFrame({
                'Familia': analisis_familias_actual.index,
                'Nº Refs': analisis_familias_actual[col_cn].astype(int),
                'Stock (uds)': analisis_familias_actual[col_stock_actual].round(0).astype(int),
                '% Stock (uds)': (analisis_familias_actual[col_stock_actual] / total_stock_uds * 100).round(1),
                'Stock (€)': analisis_familias_actual['Valor_Stock_Actual'].apply(formato_euros),
                '% Stock (€)': (analisis_familias_actual['Valor_Stock_Actual'] / total_stock_valor * 100).round(1),
                'Ventas (uds)': analisis_familias_actual['Total_Ventas'].round(0).astype(int),
                '% Ventas (uds)': (analisis_familias_actual['Total_Ventas'] / total_ventas_uds * 100).round(1),
                'Ventas (€)': analisis_familias_actual['Valor_Ventas'].apply(formato_euros),
                '% Ventas (€)': (analisis_familias_actual['Valor_Ventas'] / total_ventas_valor * 100).round(1),
                'Media uds/ref': (analisis_familias_actual[col_stock_actual] / analisis_familias_actual[col_cn]).round(1),
                'IR': analisis_familias_actual['Indice_Rotacion'].round(2)
            })
            
            st.dataframe(analisis_familias_display, use_container_width=True, height=400, hide_index=True)
            
            st.markdown("---")
            
            # ========== GRÁFICO 5: FAMILIAS FUNCIONALES CON MAYOR SOBRESTOCK ==========
            st.subheader("🚨 Familias Funcionales con Mayor Stock Sobrante")
            
            familias_unicas = sorted(df['Familia'].dropna().unique().tolist())
            familias_selec_grafico = st.multiselect(
                "Seleccionar familias para análisis (dejar vacío para ver todas):",
                options=familias_unicas,
                default=[],
                key="familias_sobrestock"
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
                            showscale=True
                        ),
                        text=[formato_euros(v) for v in top_sobrestock.values],
                        textposition='auto'
                    )
                ])
                
                fig_sobrestock.update_layout(
                    title="Top 15 Familias con Mayor Stock Sobrante",
                    xaxis_title="Valor Stock Sobrante (€)",
                    yaxis_title="Familia",
                    height=500,
                    showlegend=False
                )
                
                st.plotly_chart(fig_sobrestock, use_container_width=True)
            else:
                st.info("No hay stock sobrante en las familias seleccionadas")
            
            st.markdown("---")
            
            # ========== GRÁFICO 6: DESFASE STOCK POR FAMILIAS (ACTUAL VS LÍMITE) ==========
            st.subheader("📊 Desfase de Stock por Familias Funcionales")
            
            # Calcular el desfase para cada familia
            analisis_desfase = df.groupby('Familia').agg({
                col_cn: 'count',
                col_stock_actual: 'sum',
                'Stock_Limite': 'sum',
                'Stock_Sobrante_Uds': 'sum',
                'Stock_Faltante_Uds': 'sum',
                'Valor_Stock_Actual': 'sum',
                'Valor_Stock_Limite': 'sum',
                'Stock_Sobrante': 'sum',
                'Stock_Faltante': 'sum'
            }).round(2)
            
            # Calcular el desfase neto (positivo = sobrante, negativo = faltante)
            analisis_desfase['Desfase_Uds'] = analisis_desfase['Stock_Sobrante_Uds'] - analisis_desfase['Stock_Faltante_Uds']
            analisis_desfase['Desfase_Valor'] = analisis_desfase['Stock_Sobrante'] - analisis_desfase['Stock_Faltante']
            
            # Ordenar por mayor desfase (en valor absoluto)
            analisis_desfase['Desfase_Abs'] = analisis_desfase['Desfase_Valor'].abs()
            top_desfase = analisis_desfase.nlargest(20, 'Desfase_Abs')
            
            if len(top_desfase) > 0:
                col_desfase_graf, col_desfase_tabla = st.columns([1, 1])
                
                with col_desfase_graf:
                    # Gráfico de barras con colores según si es sobrante o faltante
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
                        title="Top 20 Familias con Mayor Desfase (Rojo=Sobrante, Azul=Faltante)",
                        xaxis_title="Desfase en Valor (€)",
                        yaxis_title="Familia",
                        height=600,
                        showlegend=False
                    )
                    
                    st.plotly_chart(fig_desfase, use_container_width=True)
                
                with col_desfase_tabla:
                    # Tabla resumen del desfase
                    desfase_display = pd.DataFrame({
                        'Familia': top_desfase.index,
                        'Stock Actual': top_desfase[col_stock_actual].round(0).astype(int),
                        'Stock Límite': top_desfase['Stock_Limite'].round(0).astype(int),
                        'Desfase (uds)': top_desfase['Desfase_Uds'].round(0).astype(int),
                        'Desfase (€)': top_desfase['Desfase_Valor'].apply(formato_euros),
                        'Sobrante': top_desfase['Stock_Sobrante'].apply(formato_euros),
                        'Faltante': top_desfase['Stock_Faltante'].apply(formato_euros)
                    })
                    
                    st.dataframe(desfase_display, use_container_width=True, height=600, hide_index=True)
            
            st.markdown("---")
            
            # ========== GRÁFICO 7: DESGLOSE POR FAMILIA ESPECÍFICA Y SUBFAMILIAS ==========
            st.subheader("🔍 Análisis Detallado por Familia y sus Subfamilias")
            
            familias_disponibles = sorted(df['Familia'].dropna().unique().tolist())
            familia_seleccionada = st.selectbox("Selecciona una familia para análisis detallado:", familias_disponibles)
            
            if familia_seleccionada:
                df_familia = df[df['Familia'] == familia_seleccionada]
                
                # Verificar si hay subfamilias diferentes a la familia
                subfamilias_unicas = df_familia['Subfamilia'].unique()
                tiene_subfamilias = len(subfamilias_unicas) > 1 or (len(subfamilias_unicas) == 1 and subfamilias_unicas[0] != familia_seleccionada)
                
                if tiene_subfamilias:
                    st.markdown(f"### 📋 Análisis por Subfamilias de: {familia_seleccionada}")
                    
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
                        'Stock Actual (uds)': analisis_subfamilias[col_stock_actual].round(0).astype(int),
                        'Stock Ideal (uds)': analisis_subfamilias['Stock_Ideal'].round(0).astype(int),
                        'Stock Límite (uds)': analisis_subfamilias['Stock_Limite'].round(0).astype(int),
                        'Sobrante (uds)': analisis_subfamilias['Stock_Sobrante_Uds'].round(0).astype(int),
                        'Faltante (uds)': analisis_subfamilias['Stock_Faltante_Uds'].round(0).astype(int),
                        'Valor Stock Actual': analisis_subfamilias['Valor_Stock_Actual'].apply(formato_euros),
                        'Valor Stock Ideal': analisis_subfamilias['Valor_Stock_Ideal'].apply(formato_euros),
                        'Valor Sobrante': analisis_subfamilias['Stock_Sobrante'].apply(formato_euros),
                        'Valor Faltante': analisis_subfamilias['Stock_Faltante'].apply(formato_euros),
                        'Ventas (uds)': analisis_subfamilias['Total_Ventas'].round(0).astype(int),
                        'Ventas (€)': analisis_subfamilias['Valor_Ventas'].apply(formato_euros),
                        'IR': analisis_subfamilias['Indice_Rotacion'].round(2)
                    })
                    
                    st.dataframe(analisis_subfamilias_display, use_container_width=True, height=400, hide_index=True)
                    
                    # Gráficos de subfamilias
                    col_graf_sub1, col_graf_sub2 = st.columns(2)
                    
                    with col_graf_sub1:
                        # Gráfico de barras: Stock Actual vs Límite por subfamilia
                        top_10_subfamilias = analisis_subfamilias.nlargest(10, 'Valor_Stock_Actual')
                        
                        fig_subfam_comp = go.Figure()
                        fig_subfam_comp.add_trace(go.Bar(
                            name='Stock Actual',
                            x=top_10_subfamilias.index,
                            y=top_10_subfamilias[col_stock_actual],
                            marker_color='#4169E1'  # Azul - actual
                        ))
                        fig_subfam_comp.add_trace(go.Bar(
                            name='Stock Ideal',
                            x=top_10_subfamilias.index,
                            y=top_10_subfamilias['Stock_Ideal'],
                            marker_color='#32CD32'  # Verde lima - ideal
                        ))
                        fig_subfam_comp.add_trace(go.Bar(
                            name='Stock Límite',
                            x=top_10_subfamilias.index,
                            y=top_10_subfamilias['Stock_Limite'],
                            marker_color='#FFA500'  # Naranja - límite
                        ))
                        
                        fig_subfam_comp.update_layout(
                            title='Top 10 Subfamilias: Stock Actual vs Ideal vs Límite',
                            barmode='group',
                            yaxis_title='Unidades',
                            xaxis_title='Subfamilia',
                            height=400,
                            showlegend=True,
                            xaxis={'tickangle': -45}
                        )
                        
                        st.plotly_chart(fig_subfam_comp, use_container_width=True)
                    
                    with col_graf_sub2:
                        # Gráfico de barras horizontales: Stock sobrante por subfamilia
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
                                title='Top 10 Subfamilias: Stock Sobrante',
                                xaxis_title='Valor Stock Sobrante (€)',
                                yaxis_title='Subfamilia',
                                height=400,
                                showlegend=False
                            )
                            
                            st.plotly_chart(fig_sobr_subfam, use_container_width=True)
                        else:
                            st.info("No hay stock sobrante en las subfamilias de esta familia")
                
                else:
                    st.markdown(f"### 📋 Análisis de: {familia_seleccionada}")
                    st.info("Esta familia no tiene subfamilias definidas")
                    
                    # Mostrar resumen de la familia
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
                        st.metric("Stock Actual", formato_euros(resumen_familia['Valor_Stock_Actual']))
                    with col_res3:
                        st.metric("Stock Sobrante", formato_euros(resumen_familia['Stock_Sobrante']))
                    with col_res4:
                        st.metric("Índice Rotación", f"{resumen_familia['Indice_Rotacion']:.2f}")
                    
                    # Distribución por categorías dentro de la familia
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
                        'Valor Stock': dist_cat_familia['Valor_Stock_Actual'].apply(formato_euros),
                        'Stock Sobrante': dist_cat_familia['Stock_Sobrante'].apply(formato_euros)
                    })
                    
                    st.dataframe(dist_cat_display, use_container_width=True, hide_index=True)
        
        # ========== EXPORTAR DATOS ==========
        st.markdown("---")
        st.subheader("📥 Exportar Datos")
        
        col_exp1, col_exp2, col_exp3 = st.columns(3)
        
        with col_exp1:
            # Exportar análisis completo
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
                label="📊 Descargar Análisis Completo (Excel)",
                data=output,
                file_name=f"analisis_stock_completo_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col_exp2:
            # Exportar productos con stock sobrante
            if len(productos_sobrantes) > 0:
                output_sobr = BytesIO()
                productos_sobrantes_export = productos_sobrantes[[
                    col_cn, col_descripcion if col_descripcion else col_cn, 
                    'Categoria', 'Familia' if 'Familia' in df.columns else col_cn,
                    col_stock_actual, 'Stock_Limite', 'Stock_Sobrante_Uds', 'Stock_Sobrante'
                ]].copy()
                
                with pd.ExcelWriter(output_sobr, engine='openpyxl') as writer:
                    productos_sobrantes_export.to_excel(writer, sheet_name='Stock Sobrante', index=False)
                
                output_sobr.seek(0)
                st.download_button(
                    label="⚠️ Descargar Stock Sobrante (Excel)",
                    data=output_sobr,
                    file_name=f"stock_sobrante_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        with col_exp3:
            # Exportar productos con stock faltante
            productos_faltantes = df[df['Stock_Faltante_Uds'] > 0].copy()
            if len(productos_faltantes) > 0:
                productos_faltantes = productos_faltantes.sort_values('Stock_Faltante', ascending=False)
                output_falt = BytesIO()
                productos_faltantes_export = productos_faltantes[[
                    col_cn, col_descripcion if col_descripcion else col_cn,
                    'Categoria', 'Familia' if 'Familia' in df.columns else col_cn,
                    col_stock_actual, 'Stock_Limite', 'Stock_Faltante_Uds', 'Stock_Faltante'
                ]].copy()
                
                with pd.ExcelWriter(output_falt, engine='openpyxl') as writer:
                    productos_faltantes_export.to_excel(writer, sheet_name='Stock Faltante', index=False)
                
                output_falt.seek(0)
                st.download_button(
                    label="📈 Descargar Stock Faltante (Excel)",
                    data=output_falt,
                    file_name=f"stock_faltante_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    
    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {str(e)}")
        st.exception(e)

else:
    st.info("👆 Por favor, carga un archivo Excel para comenzar el análisis")
    
    st.markdown("""
    ### 📋 Instrucciones de uso:
    
    1. **Carga tu archivo Excel** con los datos de ventas de la farmacia
    2. **Ajusta los parámetros** en el panel lateral según tus necesidades
    3. **Explora los diferentes análisis** que se generarán automáticamente:
       - Distribución por categorías de rotación (A, B, C, D, E)
       - Análisis por tipo de demanda
       - Comparativa de stock ideal vs actual
       - Análisis por familias funcionales
       - Identificación de stock sobrante
       - Análisis detallado por subfamilias
    4. **Descarga los reportes** en formato Excel o las listas de CNs en texto
    
    ### 📊 El archivo debe contener:
    - Columnas de ventas mensuales o una columna TOTAL
    - Columna de Stock Actual
    - Columna de PVP
    - Columna de CN (Código Nacional)
    - Columna de Categoría funcional (opcional, para análisis por familias)
    """)
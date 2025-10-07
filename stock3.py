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
    
    # Categorizar productos
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
        if 'stock' in col_lower and ('actual' in col_lower or col_lower == 'stock actual'):
            col_stock_actual = col
        elif col_lower == 'pvp':
            col_pvp = col
        elif col_lower == 'cn':
            col_cn = col
        elif 'descripcion' in col_lower or 'descripción' in col_lower:
            col_descripcion = col
        elif 'categoria' in col_lower and 'funcional' in col_lower:
            col_categoria_funcional = col
        elif col_lower == 'categoria' or col_lower == 'categoría':
            if col_categoria_funcional is None:
                col_categoria_funcional = col
    
    # Limpiar PVP
    if col_pvp:
        df[col_pvp] = df[col_pvp].astype(str).str.replace('€', '').str.replace(',', '.').str.strip()
        df[col_pvp] = pd.to_numeric(df[col_pvp], errors='coerce').fillna(0)
    
    # Calcular valores
    if col_stock_actual and col_pvp:
        df[col_stock_actual] = pd.to_numeric(df[col_stock_actual], errors='coerce').fillna(0)
        
        # Valor del stock actual
        df['Valor_Stock_Actual'] = df[col_stock_actual] * df[col_pvp]
        
        # Valor del stock ideal y límite
        df['Valor_Stock_Ideal'] = df['Stock_Ideal'] * df[col_pvp]
        df['Valor_Stock_Limite'] = df['Stock_Limite'] * df[col_pvp]
        
        # Stock sobrante (en unidades y valor) - CORREGIDO
        df['Stock_Sobrante_Uds'] = np.where(
            df[col_stock_actual] > df['Stock_Ideal'],
            df[col_stock_actual] - df['Stock_Ideal'],
            0
        )
        df['Stock_Sobrante'] = df['Stock_Sobrante_Uds'] * df[col_pvp]
        
        # Stock faltante (en unidades y valor) - CORREGIDO
        df['Stock_Faltante_Uds'] = np.where(
            df[col_stock_actual] < df['Stock_Ideal'],
            df['Stock_Ideal'] - df[col_stock_actual],
            0
        )
        df['Stock_Faltante'] = df['Stock_Faltante_Uds'] * df[col_pvp]
        
        # Reposición necesaria
        df['Reposicion'] = df['Stock_Ideal'] - df[col_stock_actual]
        
        # Índice de rotación
        df['Indice_Rotacion'] = df.apply(
            lambda row: calcular_indice_rotacion(row['Total_Ventas'], row[col_stock_actual]),
            axis=1
        )
        
        # Valor de ventas
        df['Valor_Ventas'] = df['Total_Ventas'] * df[col_pvp]
    
    # Procesar categoría funcional
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
        
        # ========== GRÁFICO 1: DISTRIBUCIÓN POR CATEGORÍAS ==========
        st.subheader("📈 Distribución por Categorías de Rotación")
        
        col_graf1, col_graf2 = st.columns([1, 1])
        
        with col_graf1:
            # Descripción de categorías
            cat_descripciones = {
                'A': 'Alta rotación (>260 uds/año)',
                'B': 'Rotación media-alta (52-260 uds/año)',
                'C': 'Rotación media (12-51 uds/año)',
                'D': 'Rotación baja (1-11 uds/año)',
                'E': 'Sin rotación (<1 ud/año)'
            }
            
            categoria_counts = df['Categoria'].value_counts().reindex(['A', 'B', 'C', 'D', 'E'], fill_value=0)
            
            # Preparar texto hover
            hover_text = [f"{cat}: {cat_descripciones[cat]}<br>Productos: {count}" 
                         for cat, count in categoria_counts.items()]
            
            fig_pie = go.Figure(data=[go.Pie(
                labels=categoria_counts.index,
                values=categoria_counts.values,
                hole=0.4,
                marker=dict(colors=['#667eea', '#764ba2', '#f093fb', '#4facfe', '#43e97b']),
                textinfo='label+percent',
                textfont=dict(size=14),
                hovertext=hover_text,
                hoverinfo='text'
            )])
            
            fig_pie.update_layout(
                title="Proporción de Productos por Categoría",
                height=400,
                showlegend=True
            )
            
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col_graf2:
            # Tabla resumen con stock sobrante
            resumen_cat = df.groupby('Categoria').agg({
                'Total_Ventas': 'sum',
                'Valor_Stock_Actual': 'sum',
                'Stock_Sobrante': 'sum'
            }).reindex(['A', 'B', 'C', 'D', 'E'], fill_value=0).round(2)
            
            resumen_cat_display = pd.DataFrame({
                'Categoría': resumen_cat.index,
                'Total Ventas (uds)': resumen_cat['Total_Ventas'].apply(formato_numero),
                'Valor Stock': resumen_cat['Valor_Stock_Actual'].apply(formato_euros),
                'Stock Sobrante': resumen_cat['Stock_Sobrante'].apply(formato_euros)
            })
            
            st.dataframe(resumen_cat_display, use_container_width=True, height=400, hide_index=True)
        
        st.markdown("---")
        
        # ========== GRÁFICO 2: DESGLOSE POR DEMANDA ==========
        st.subheader("📊 Análisis por Tipo de Demanda")
        
        analisis_demanda = df.groupby('Categoria').agg({
            col_cn: 'count',
            col_stock_actual: 'sum',
            'Total_Ventas': 'sum',
            'Indice_Rotacion': 'mean'
        }).reindex(['A', 'B', 'C', 'D', 'E'], fill_value=0)
        
        total_refs = len(df)
        total_stock = df[col_stock_actual].sum()
        
        analisis_demanda_display = pd.DataFrame({
            'Categoría': analisis_demanda.index,
            'Nº Referencias': analisis_demanda[col_cn].astype(int),
            '% Refs': (analisis_demanda[col_cn] / total_refs * 100).round(1),
            'Stock (uds)': analisis_demanda[col_stock_actual].round(0).astype(int),
            '% Stock': (analisis_demanda[col_stock_actual] / total_stock * 100).round(1),
            'Media uds/ref': (analisis_demanda[col_stock_actual] / analisis_demanda[col_cn]).round(1),
            'Índice Rotación': analisis_demanda['Indice_Rotacion'].round(2)
        })
        
        st.dataframe(analisis_demanda_display, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        
        # ========== GRÁFICO 3: ANÁLISIS POR FAMILIAS - STOCK ACTUAL ==========
        if col_categoria_funcional and 'Familia' in df.columns:
            st.subheader("🏪 Análisis por Familias - Stock Actual")
            
            analisis_familias_actual = df.groupby('Familia').agg({
                col_cn: 'count',
                col_stock_actual: 'sum',
                'Valor_Stock_Actual': 'sum',
                'Total_Ventas': 'sum',
                'Valor_Ventas': 'sum',
                'Indice_Rotacion': 'mean'
            }).round(2)
            
            total_stock_uds = df[col_stock_actual].sum()
            total_stock_valor = df['Valor_Stock_Actual'].sum()
            total_ventas_uds = df['Total_Ventas'].sum()
            total_ventas_valor = df['Valor_Ventas'].sum()
            
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
            
            # ========== GRÁFICO 4: ANÁLISIS POR FAMILIAS - COMPARATIVA STOCK IDEAL ==========
            st.subheader("🎯 Análisis por Familias - Comparativa Stock Ideal vs Actual")
            
            analisis_familias_ideal = df.groupby('Familia').agg({
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
            }).round(2)
            
            analisis_familias_comp = pd.DataFrame({
                'Familia': analisis_familias_ideal.index,
                'Nº Refs Actual': analisis_familias_ideal[col_cn].astype(int),
                'Stock Actual': analisis_familias_ideal[col_stock_actual].round(0).astype(int),
                'Stock Ideal': analisis_familias_ideal['Stock_Ideal'].round(0).astype(int),
                'Stock Límite': analisis_familias_ideal['Stock_Limite'].round(0).astype(int),
                'Stock Sobrante': analisis_familias_ideal['Stock_Sobrante_Uds'].round(0).astype(int),
                'Stock Faltante': analisis_familias_ideal['Stock_Faltante_Uds'].round(0).astype(int),
                'Valor Actual': analisis_familias_ideal['Valor_Stock_Actual'].apply(formato_euros),
                'Valor Límite': analisis_familias_ideal['Valor_Stock_Limite'].apply(formato_euros),
                'Valor Sobrante': analisis_familias_ideal['Stock_Sobrante'].apply(formato_euros),
                'Valor Faltante': analisis_familias_ideal['Stock_Faltante'].apply(formato_euros)
            })
            
            st.dataframe(analisis_familias_comp, use_container_width=True, height=400, hide_index=True)
            
            st.markdown("---")
            
            # ========== GRÁFICO 5: FAMILIAS CON MAYOR SOBRESTOCK ==========
            st.subheader("🚨 Familias con Mayor Necesidad de Eliminación de Stock")
            
            top_sobrestock = df.groupby('Familia')['Stock_Sobrante'].sum().sort_values(ascending=False).head(10)
            
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
                title="Top 10 Familias con Mayor Stock Sobrante",
                xaxis_title="Valor Stock Sobrante (€)",
                yaxis_title="Familia",
                height=500,
                showlegend=False
            )
            
            st.plotly_chart(fig_sobrestock, use_container_width=True)
            
            st.markdown("---")
            
            # ========== GRÁFICO 6: DESGLOSE POR FAMILIA ESPECÍFICA ==========
            st.subheader("🔍 Análisis Detallado por Familia")
            
            familias_disponibles = sorted(df['Familia'].dropna().unique().tolist())
            familia_seleccionada = st.selectbox("Selecciona una familia para análisis detallado:", familias_disponibles)
            
            if familia_seleccionada:
                df_familia = df[df['Familia'] == familia_seleccionada]
                
                analisis_familia_cat = df_familia.groupby('Categoria').agg({
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
                }).reindex(['A', 'B', 'C', 'D', 'E'], fill_value=0).round(2)
                
                display_familia_cat = pd.DataFrame({
                    'Cat.': analisis_familia_cat.index,
                    'Refs': analisis_familia_cat[col_cn].astype(int),
                    'Stock Actual': analisis_familia_cat[col_stock_actual].round(0).astype(int),
                    'Stock Ideal': analisis_familia_cat['Stock_Ideal'].round(0).astype(int),
                    'Stock Límite': analisis_familia_cat['Stock_Limite'].round(0).astype(int),
                    'Sobrante': analisis_familia_cat['Stock_Sobrante_Uds'].round(0).astype(int),
                    'Faltante': analisis_familia_cat['Stock_Faltante_Uds'].round(0).astype(int),
                    'Valor Actual': analisis_familia_cat['Valor_Stock_Actual'].apply(formato_euros),
                    'Valor Ideal': analisis_familia_cat['Valor_Stock_Ideal'].apply(formato_euros),
                    'Valor Límite': analisis_familia_cat['Valor_Stock_Limite'].apply(formato_euros),
                    'Valor Sobrante': analisis_familia_cat['Stock_Sobrante'].apply(formato_euros),
                    'Valor Faltante': analisis_familia_cat['Stock_Faltante'].apply(formato_euros),
                    'Ventas (uds)': analisis_familia_cat['Total_Ventas'].round(0).astype(int),
                    'Ventas (€)': analisis_familia_cat['Valor_Ventas'].apply(formato_euros),
                    'Media uds/ref': (analisis_familia_cat[col_stock_actual] / analisis_familia_cat[col_cn]).round(1),
                    'IR': analisis_familia_cat['Indice_Rotacion'].round(2)
                })
                
                st.dataframe(display_familia_cat, use_container_width=True, hide_index=True)
                
                st.markdown("---")
                
                # ========== GRÁFICO 7: DESGLOSE POR SUBFAMILIAS ==========
                st.subheader(f"📋 Análisis por Subfamilias de: {familia_seleccionada}")
                
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
                }).round(2)
                
                display_subfamilias = pd.DataFrame({
                    'Subfamilia': analisis_subfamilias.index,
                    'Refs': analisis_subfamilias[col_cn].astype(int),
                    'Stock Actual': analisis_subfamilias[col_stock_actual].round(0).astype(int),
                    'Stock Ideal': analisis_subfamilias['Stock_Ideal'].round(0).astype(int),
                    'Stock Límite': analisis_subfamilias['Stock_Limite'].round(0).astype(int),
                    'Sobrante': analisis_subfamilias['Stock_Sobrante_Uds'].round(0).astype(int),
                    'Faltante': analisis_subfamilias['Stock_Faltante_Uds'].round(0).astype(int),
                    'Valor Actual': analisis_subfamilias['Valor_Stock_Actual'].apply(formato_euros),
                    'Valor Ideal': analisis_subfamilias['Valor_Stock_Ideal'].apply(formato_euros),
                    'Valor Límite': analisis_subfamilias['Valor_Stock_Limite'].apply(formato_euros),
                    'Valor Sobrante': analisis_subfamilias['Stock_Sobrante'].apply(formato_euros),
                    'Valor Faltante': analisis_subfamilias['Stock_Faltante'].apply(formato_euros),
                    'Ventas (uds)': analisis_subfamilias['Total_Ventas'].round(0).astype(int),
                    'Ventas (€)': analisis_subfamilias['Valor_Ventas'].apply(formato_euros),
                    'Media uds/ref': (analisis_subfamilias[col_stock_actual] / analisis_subfamilias[col_cn]).round(1),
                    'IR': analisis_subfamilias['Indice_Rotacion'].round(2)
                })
                
                st.dataframe(display_subfamilias, use_container_width=True, height=500, hide_index=True)
        
        st.markdown("---")
        
        # ========== FILTROS Y TABLA DE DETALLE ==========
        st.sidebar.markdown("---")
        st.sidebar.subheader("🔎 Filtros")
        
        categorias_seleccionadas = st.sidebar.multiselect(
            "Filtrar por categoría de rotación",
            options=['A', 'B', 'C', 'D', 'E'],
            default=['A', 'B', 'C', 'D', 'E']
        )
        
        # Filtros por familia/subfamilia
        if col_categoria_funcional and 'Familia' in df.columns:
            familias_disponibles = df['Familia'].dropna().unique().tolist()
            familias_seleccionadas = st.sidebar.multiselect(
                "Filtrar por Familia",
                options=sorted(familias_disponibles),
                default=[]
            )
            
            if familias_seleccionadas:
                subfamilias_disponibles = df[df['Familia'].isin(familias_seleccionadas)]['Subfamilia'].dropna().unique().tolist()
                subfamilias_seleccionadas = st.sidebar.multiselect(
                    "Filtrar por Subfamilia",
                    options=sorted(subfamilias_disponibles),
                    default=[]
                )
            else:
                subfamilias_seleccionadas = []
        else:
            familias_seleccionadas = []
            subfamilias_seleccionadas = []
        
        # Aplicar filtros
        df_filtrado = df[df['Categoria'].isin(categorias_seleccionadas)].copy()
        
        if familias_seleccionadas:
            df_filtrado = df_filtrado[df_filtrado['Familia'].isin(familias_seleccionadas)]
        
        if subfamilias_seleccionadas:
            df_filtrado = df_filtrado[df_filtrado['Subfamilia'].isin(subfamilias_seleccionadas)]
        
        # Ordenar por valor de stock descendente
        if 'Valor_Stock_Actual' in df_filtrado.columns:
            df_filtrado = df_filtrado.sort_values('Valor_Stock_Actual', ascending=False)
        
        # Tabla de resultados
        st.subheader(f"📋 Detalle de Productos ({len(df_filtrado):,} productos)".replace(",", "."))
        
        # Preparar columnas a mostrar
        columnas_mostrar = []
        
        if col_cn:
            columnas_mostrar.append(col_cn)
        if col_descripcion:
            columnas_mostrar.append(col_descripcion)
        if col_categoria_funcional:
            columnas_mostrar.extend(['Familia', 'Subfamilia'])
        
        columnas_mostrar.extend(['Categoria', 'Total_Ventas', 'Vtas_Dia'])
        
        if col_stock_actual:
            columnas_mostrar.append(col_stock_actual)
        
        columnas_mostrar.extend(['Stock_Min_Calc', 'Stock_Ideal', 'Stock_Limite'])
        
        if 'Indice_Rotacion' in df_filtrado.columns:
            columnas_mostrar.append('Indice_Rotacion')
        if 'Reposicion' in df_filtrado.columns:
            columnas_mostrar.append('Reposicion')
        if 'Stock_Sobrante_Uds' in df_filtrado.columns:
            columnas_mostrar.append('Stock_Sobrante_Uds')
        if 'Stock_Faltante_Uds' in df_filtrado.columns:
            columnas_mostrar.append('Stock_Faltante_Uds')
        if 'Valor_Stock_Actual' in df_filtrado.columns:
            columnas_mostrar.extend(['Valor_Stock_Actual', 'Stock_Sobrante', 'Stock_Faltante'])
        
        columnas_mostrar = [col for col in columnas_mostrar if col in df_filtrado.columns]
        
        # Mostrar tabla
        df_display = df_filtrado[columnas_mostrar].copy()
        
        # Formatear columnas numéricas para visualización
        for col in df_display.columns:
            if df_display[col].dtype in ['float64', 'int64'] and col not in [col_cn]:
                df_display[col] = df_display[col].round(2)
        
        st.dataframe(df_display, use_container_width=True, height=500)
        
        # Botones de descarga
        st.markdown("---")
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        
        with col_btn1:
            # Descargar CNs seleccionados
            if col_cn:
                cns_texto = "\n".join(df_filtrado[col_cn].astype(str).tolist())
                st.download_button(
                    label=f"📋 Descargar CNs ({len(df_filtrado)})",
                    data=cns_texto,
                    file_name=f"CNs_seleccionados_{pd.Timestamp.now().strftime('%Y%m%d')}.txt",
                    mime="text/plain"
                )
        
        with col_btn2:
            # Descargar Excel completo
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_filtrado.to_excel(writer, index=False, sheet_name='Analisis')
            
            st.download_button(
                label="📊 Descargar Excel Completo",
                data=output.getvalue(),
                file_name=f"analisis_stock_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col_btn3:
            # Descargar CSV
            csv = df_filtrado.to_csv(index=False, encoding='utf-8-sig', decimal=',', sep=';')
            st.download_button(
                label="📄 Descargar CSV",
                data=csv.encode('utf-8-sig'),
                file_name=f"analisis_stock_{pd.Timestamp.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        
    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {str(e)}")
        st.exception(e)
else:
    st.info("👋 Por favor, carga un archivo Excel para comenzar el análisis")
    st.markdown("""
    ### 📝 Instrucciones:
    
    **El archivo Excel debe contener:**
    1. **Columnas de ventas:** Mensuales O una columna TOTAL
    2. **Columnas obligatorias:** CN, Descripción, PVP, Stock Actual
    3. **Opcional:** Columna 'Categoria' con familias (ej: ADELG-ANTICELULÍTICOS)
    
    **Categorización automática de productos:**
    - **A:** Alta rotación (>260 uds/año) - ~5 ventas/semana
    - **B:** Rotación media-alta (52-260 uds/año) - 1-5 ventas/semana
    - **C:** Rotación media (12-51 uds/año) - 1-4 ventas/mes
    - **D:** Rotación baja (1-11 uds/año) - <1 venta/mes
    - **E:** Sin rotación (<1 ud/año)
    
    **Cálculo de stocks:**
    - **Stock Ideal:** Nivel óptimo basado en días de cobertura
    - **Stock Límite:** Stock ideal + margen de seguridad (configurable)
    - **Stock Sobrante:** Cuando stock actual > stock ideal
    - **Stock Faltante:** Cuando stock actual < stock ideal
    
    Los productos se ordenarán automáticamente por valor de stock descendente.
    """)
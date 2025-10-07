# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="Analisis Stock Farmacia", layout="wide")

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

st.title("Analisis de Stock Farmaceutico")
st.markdown("---")

# Funciones auxiliares
def formato_euros(valor):
    """Formatea un número con punto para miles y coma para decimales"""
    return f"{valor:,.2f}€".replace(",", "X").replace(".", ",").replace("X", ".")

def calcular_indice_rotacion(categoria, ventas_anuales, stock_actual):
    """Calcula el índice de rotación anual"""
    if stock_actual == 0 or pd.isna(stock_actual):
        return 0
    return round(ventas_anuales / stock_actual, 2)

@st.cache_data
def procesar_excel(uploaded_file, dias_abierto, stock_min_dias, stock_max_dias, dias_cobertura_optimo):
    """Procesa el archivo Excel y calcula todos los valores (con cache para velocidad)"""
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
            stock_min = vtas_dia * stock_min_dias
            stock_max = vtas_dia * stock_max_dias
            stock_opt = vtas_dia * dias_cobertura_optimo
        elif cat == 'C':
            stock_min = 1
            stock_max = 2
            stock_opt = 1
        elif cat == 'D':
            stock_min = 0
            stock_max = 1
            stock_opt = 1
        else:  # E
            stock_min = 0
            stock_max = 0
            stock_opt = 0
        
        return pd.Series({
            'Stock_Min_Calc': round(stock_min, 1),
            'Stock_Max_Calc': round(stock_max, 1),
            'Stock_Opt_Calc': round(stock_opt, 1)
        })
    
    df[['Stock_Min_Calc', 'Stock_Max_Calc', 'Stock_Opt_Calc']] = df.apply(calcular_stocks, axis=1)
    
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
            if col_categoria_funcional is None:  # Solo si no hemos encontrado ya una columna de categoría funcional
                col_categoria_funcional = col
    
    # Limpiar PVP
    if col_pvp:
        df[col_pvp] = df[col_pvp].astype(str).str.replace('€', '').str.replace(',', '.').str.strip()
        df[col_pvp] = pd.to_numeric(df[col_pvp], errors='coerce').fillna(0)
    
    # Calcular valores
    if col_stock_actual and col_pvp:
        df[col_stock_actual] = pd.to_numeric(df[col_stock_actual], errors='coerce').fillna(0)
        
        df['Valor_Stock_Actual'] = df[col_stock_actual] * df[col_pvp]
        df['Valor_Stock_Optimo'] = df['Stock_Opt_Calc'] * df[col_pvp]
        
        df['Stock_Sobrante'] = np.where(
            df[col_stock_actual] > df['Stock_Opt_Calc'],
            (df[col_stock_actual] - df['Stock_Opt_Calc']) * df[col_pvp],
            0
        )
        
        df['Stock_Faltante'] = np.where(
            df[col_stock_actual] < df['Stock_Opt_Calc'],
            (df['Stock_Opt_Calc'] - df[col_stock_actual]) * df[col_pvp],
            0
        )
        
        df['Reposicion'] = df['Stock_Opt_Calc'] - df[col_stock_actual]
        
        # Calcular índice de rotación
        df['Indice_Rotacion'] = df.apply(
            lambda row: calcular_indice_rotacion(row['Categoria'], row['Total_Ventas'], row[col_stock_actual]),
            axis=1
        )
    
    # Procesar categoría funcional
    if col_categoria_funcional:
        df['Familia'] = df[col_categoria_funcional].astype(str).str.split('-').str[0]
        df['Subfamilia'] = df[col_categoria_funcional]
    
    return df, col_stock_actual, col_pvp, col_cn, col_descripcion, col_categoria_funcional

# Sidebar
st.sidebar.header("Configuracion")
dias_abierto = st.sidebar.number_input("Dias abierto al año", min_value=250, max_value=365, value=300, step=1)

st.sidebar.markdown("### Parametros Stock (A y B)")
col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    stock_min_dias = st.number_input("Min (dias)", min_value=5, max_value=20, value=10, step=1)
with col_sb2:
    stock_max_dias = st.number_input("Max (dias)", min_value=15, max_value=40, value=30, step=1)

dias_cobertura_optimo = st.sidebar.slider("Dias cobertura optima (A y B)", min_value=10, max_value=30, value=15, step=1)

# Upload Excel
uploaded_file = st.file_uploader("Cargar archivo Excel con datos de ventas", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        df, col_stock_actual, col_pvp, col_cn, col_descripcion, col_categoria_funcional = procesar_excel(
            uploaded_file, dias_abierto, stock_min_dias, stock_max_dias, dias_cobertura_optimo
        )
        
        st.success(f"Archivo procesado correctamente")
        
        # KPIs principales
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Productos", f"{len(df):,}".replace(",", "."))
        with col2:
            if 'Valor_Stock_Actual' in df.columns:
                st.metric("Valor Stock Actual", formato_euros(df['Valor_Stock_Actual'].sum()))
        with col3:
            if 'Stock_Sobrante' in df.columns:
                st.metric("Stock Sobrante", formato_euros(df['Stock_Sobrante'].sum()))
        with col4:
            if 'Stock_Faltante' in df.columns:
                st.metric("Stock Faltante", formato_euros(df['Stock_Faltante'].sum()))
        
        st.markdown("---")
        
        # Grafico de distribución por categorías
        st.subheader("Distribucion por Categorias de Rotacion")
        
        col_graf1, col_graf2 = st.columns([1, 1])
        
        with col_graf1:
            # Gráfico de tarta
            categoria_counts = df['Categoria'].value_counts().sort_index()
            
            fig_pie = go.Figure(data=[go.Pie(
                labels=categoria_counts.index,
                values=categoria_counts.values,
                hole=0.4,
                marker=dict(colors=['#667eea', '#764ba2', '#f093fb', '#4facfe', '#43e97b']),
                textinfo='label+percent',
                textfont=dict(size=14)
            )])
            
            fig_pie.update_layout(
                title="Proporcion de Productos por Categoria",
                height=400,
                showlegend=True
            )
            
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col_graf2:
            # Tabla resumen con índice de rotación
            resumen_cat = df.groupby('Categoria').agg({
                'Total_Ventas': 'sum',
                'Valor_Stock_Actual': 'sum' if 'Valor_Stock_Actual' in df.columns else 'count',
                'Indice_Rotacion': 'mean' if 'Indice_Rotacion' in df.columns else 'count'
            }).round(2)
            
            resumen_cat.columns = ['Total Ventas', 'Valor Stock (€)', 'Indice Rotacion Medio']
            resumen_cat['Valor Stock (€)'] = resumen_cat['Valor Stock (€)'].apply(formato_euros)
            
            st.dataframe(resumen_cat, use_container_width=True, height=400)
        
        # Filtros
        st.sidebar.markdown("---")
        st.sidebar.subheader("Filtros")
        
        categorias_seleccionadas = st.sidebar.multiselect(
            "Filtrar por categoria de rotacion",
            options=['A', 'B', 'C', 'D', 'E'],
            default=['A', 'B', 'C', 'D', 'E']
        )
        
        # Filtros por familia/subfamilia si existen
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
        st.markdown("---")
        st.subheader(f"Detalle de Productos ({len(df_filtrado)} productos)")
        
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
        
        columnas_mostrar.extend(['Stock_Min_Calc', 'Stock_Max_Calc', 'Stock_Opt_Calc'])
        
        if 'Indice_Rotacion' in df_filtrado.columns:
            columnas_mostrar.append('Indice_Rotacion')
        if 'Reposicion' in df_filtrado.columns:
            columnas_mostrar.append('Reposicion')
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
        st.error(f"Error al procesar el archivo: {str(e)}")
        st.exception(e)
else:
    st.info("Por favor, carga un archivo Excel para comenzar el analisis")
    st.markdown("""
    ### Instrucciones:
    1. El archivo debe contener:
       - Columnas de ventas mensuales O una columna TOTAL
       - Columnas: CN, Descripcion, PVP, Stock Actual
    2. Opcional: Columna 'Categoria' con familias (ej: ADELG-ANTICELULITICOS)
    3. Los productos se ordenaran automaticamente por valor de stock
    4. Ajusta los parametros en el panel lateral
    """)
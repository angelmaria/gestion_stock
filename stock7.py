# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Análisis Stock Farmacia", layout="wide")

# ==================== CONFIGURACIÓN Y CONSTANTES ====================
FAMILIAS_MAP = {
    'ADELG': 'ADELGAZANTES', 'ANTICEL': 'ANTICELULITICOS', 'AROMA': 'AROMATERAPIA',
    'DEPORTE': 'DEPORTE', 'DERMO': 'DERMO', 'DIETSOE': 'DIET SOE', 'DIET': 'DIETETICA',
    'EFECSOE': 'EFEC SOE', 'EFEC': 'EFECTOS', 'EFP': 'EFP', 'ESPEC': 'ESPECIALIDAD',
    'ESPECSR': 'ESPECIALIDAD', 'FITO': 'FITOTERAPIA', 'HIGBUC': 'HIG.BUCAL',
    'HIGCAP': 'HIG.CAPILAR', 'HIGCORP': 'HIG.CORPORAL', 'HOMEO': 'HOMEOPATIA',
    'INFAN': 'INFANTIL', 'INFANSOE': 'INFANTIL SOE', 'INSEC': 'INSECTOS',
    'NASOI': 'NARIZ OIDOS', 'OPTIC': 'OPTICA', 'ORTO': 'ORTOPEDIA',
    'ORTOSOE': 'ORTOPEDIA SOE', 'PIEMAN': 'PIES/MANOS', 'GINEC': 'SALUD GINECOLOGICA',
    'SEX': 'SALUD SEXUAL', 'SOL': 'SOLARES', 'VET': 'VETERINARIA',
    'VACUNAS': 'VACUNAS', 'FORMULAS': 'FORMULAS', 'ENVASE': 'ENVASE CLINICO'
}

# ==================== FUNCIONES AUXILIARES ====================
def aplicar_estilos():
    st.markdown("""
    <style>
        .stButton>button {
            width: 100%;
            border-radius: 8px;
            height: 3em;
            font-weight: 500;
        }
    </style>
    """, unsafe_allow_html=True)

def formato_euros(valor):
    return f"{valor:,.2f}€".replace(",", "X").replace(".", ",").replace("X", ".")

def formato_numero(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def extraer_familia(categoria_str):
    if pd.isna(categoria_str):
        return 'SIN CLASIFICAR'
    
    # Intentar extraer el prefijo antes del guión
    partes = str(categoria_str).split('-')
    if len(partes) > 0:
        prefijo = partes[0].strip().upper()
        # Buscar coincidencia exacta o parcial
        for key, value in FAMILIAS_MAP.items():
            if prefijo.startswith(key) or key.startswith(prefijo):
                return value
    
    return 'OTROS'

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

def calcular_stocks_por_categoria(row, dias_cobertura_optimo, stock_min_dias, margen_seguridad):
    """CORRECCIÓN: Cálculo correcto de stocks según categoría"""
    cat = row['Categoria']
    vtas_dia = row['Vtas_Dia']
    
    if cat == 'A' or cat == 'B':
        stock_ideal = vtas_dia * dias_cobertura_optimo
        stock_min = vtas_dia * stock_min_dias
        stock_limite = stock_ideal * (1 + margen_seguridad)
    elif cat == 'C':
        stock_ideal = 1
        stock_min = 1
        stock_limite = 1 * (1 + margen_seguridad)
    elif cat == 'D':
        stock_ideal = 1
        stock_min = 0
        stock_limite = 1 * (1 + margen_seguridad)
    else:  # E
        stock_ideal = 0
        stock_min = 0
        stock_limite = 0
    
    return pd.Series({
        'Stock_Min_Calc': round(stock_min, 1),
        'Stock_Ideal': round(stock_ideal, 1),
        'Stock_Limite': round(stock_limite, 1)
    })

def detectar_columnas(df):
    """Detecta automáticamente las columnas relevantes del DataFrame"""
    cols = {
        'total': None, 'stock_actual': None, 'pvp': None, 
        'cn': None, 'descripcion': None, 'categoria_funcional': None
    }
    
    for col in df.columns:
        col_lower = str(col).lower()
        
        if 'total' in col_lower and 'ventas' not in col_lower and cols['total'] is None:
            cols['total'] = col
        elif 'stock' in col_lower and 'actual' in col_lower:
            cols['stock_actual'] = col
        elif col_lower == 'pvp':
            cols['pvp'] = col
        elif col_lower in ['cn', 'codigo'] or 'idarti' in col_lower:
            if cols['cn'] is None:
                cols['cn'] = col
        elif 'descripcion' in col_lower or 'descripción' in col_lower:
            cols['descripcion'] = col
        elif ('categoria' in col_lower and 'funcional' in col_lower) or col_lower in ['categoria', 'categoría']:
            if cols['categoria_funcional'] is None:
                cols['categoria_funcional'] = col
    
    return cols

def calcular_ventas_totales(df, col_total):
    """Calcula las ventas totales desde columna TOTAL o sumando meses"""
    if col_total:
        return pd.to_numeric(df[col_total], errors='coerce').fillna(0)
    
    # Buscar columnas mensuales
    meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
             'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
    
    columnas_ventas = []
    for col in df.columns:
        col_lower = str(col).lower()
        if 'ventas' in col_lower or any(mes in col_lower for mes in meses):
            columnas_ventas.append(col)
    
    if columnas_ventas:
        return df[columnas_ventas].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)
    
    return pd.Series(0, index=df.index)

# ==================== FUNCIÓN PRINCIPAL DE PROCESAMIENTO ====================
def procesar_excel(uploaded_file, dias_abierto, stock_min_dias, stock_max_dias, 
                   dias_cobertura_optimo, margen_seguridad):
    """Procesa el Excel y calcula todos los indicadores"""
    
    # Leer Excel
    df = pd.read_excel(uploaded_file)
    
    # Detectar columnas
    cols = detectar_columnas(df)
    
    # Calcular ventas totales
    df['Total_Ventas'] = calcular_ventas_totales(df, cols['total'])
    
    # Calcular ventas diarias
    df['Vtas_Dia'] = df['Total_Ventas'] / dias_abierto
    
    # Categorizar productos
    df['Categoria'] = df['Total_Ventas'].apply(categorizar_producto)
    
    # Calcular stocks según categoría
    df[['Stock_Min_Calc', 'Stock_Ideal', 'Stock_Limite']] = df.apply(
        lambda row: calcular_stocks_por_categoria(row, dias_cobertura_optimo, stock_min_dias, margen_seguridad),
        axis=1
    )
    
    # Limpiar y procesar PVP
    if cols['pvp']:
        df[cols['pvp']] = df[cols['pvp']].astype(str).str.replace('€', '').str.replace(',', '.').str.strip()
        df[cols['pvp']] = pd.to_numeric(df[cols['pvp']], errors='coerce').fillna(0)
    
    # Procesar stock actual
    if cols['stock_actual']:
        df[cols['stock_actual']] = pd.to_numeric(df[cols['stock_actual']], errors='coerce').fillna(0)
    
    # Calcular valores monetarios y excesos/déficits
    if cols['stock_actual'] and cols['pvp']:
        df['Valor_Stock_Actual'] = df[cols['stock_actual']] * df[cols['pvp']]
        df['Valor_Stock_Ideal'] = df['Stock_Ideal'] * df[cols['pvp']]
        df['Valor_Stock_Limite'] = df['Stock_Limite'] * df[cols['pvp']]
        
        # CORRECCIÓN: Stock sobrante cuando actual > ideal
        df['Stock_Sobrante_Uds'] = np.maximum(0, df[cols['stock_actual']] - df['Stock_Ideal'])
        df['Stock_Sobrante'] = df['Stock_Sobrante_Uds'] * df[cols['pvp']]
        
        # CORRECCIÓN: Stock faltante cuando actual < ideal
        df['Stock_Faltante_Uds'] = np.maximum(0, df['Stock_Ideal'] - df[cols['stock_actual']])
        df['Stock_Faltante'] = df['Stock_Faltante_Uds'] * df[cols['pvp']]
        
        df['Reposicion'] = df['Stock_Ideal'] - df[cols['stock_actual']]
        
        # Índice de rotación
        df['Indice_Rotacion'] = np.where(
            df[cols['stock_actual']] > 0,
            df['Total_Ventas'] / df[cols['stock_actual']],
            0
        ).round(2)
        
        df['Valor_Ventas'] = df['Total_Ventas'] * df[cols['pvp']]
    
    # Procesar familias funcionales
    if cols['categoria_funcional']:
        df['Familia'] = df[cols['categoria_funcional']].apply(extraer_familia)
        df['Subfamilia'] = df[cols['categoria_funcional']]
    else:
        df['Familia'] = 'SIN CLASIFICAR'
        df['Subfamilia'] = 'SIN CLASIFICAR'
    
    return df, cols

# ==================== COMPONENTES DE VISUALIZACIÓN ====================
def mostrar_resumen_ejecutivo(df, col_stock_actual):
    """Muestra el resumen ejecutivo con métricas principales"""
    with st.expander("📊 Resumen Ejecutivo", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        
        total_inversion = df['Valor_Stock_Actual'].sum()
        total_ideal = df['Valor_Stock_Ideal'].sum()
        total_sobrante = df['Stock_Sobrante'].sum()
        total_faltante = df['Stock_Faltante'].sum()
        
        with col1:
            st.metric("Inversión Total en Stock", formato_euros(total_inversion))
        with col2:
            delta = total_inversion - total_ideal
            st.metric("Inversión Ideal Objetivo", formato_euros(total_ideal),
                     delta=formato_euros(delta) if delta != 0 else None)
        with col3:
            pct = (total_sobrante / total_inversion * 100) if total_inversion > 0 else 0
            st.metric("💰 Exceso de Stock", formato_euros(total_sobrante),
                     delta=f"{pct:.1f}% del total")
        with col4:
            pct = (total_faltante / total_ideal * 100) if total_ideal > 0 else 0
            st.metric("⚠️ Déficit de Stock", formato_euros(total_faltante),
                     delta=f"{pct:.1f}% del ideal")
        
        # Alertas
        if total_sobrante > total_faltante:
            st.warning(f"⚠️ **Exceso de inventario**: El exceso supera al déficit en {formato_euros(total_sobrante - total_faltante)}")
        elif total_faltante > total_sobrante:
            st.info(f"📈 **Oportunidad de optimización**: Déficit de {formato_euros(total_faltante - total_sobrante)}")

def grafico_distribucion_categorias(df, cols):
    """Gráfico de distribución por categorías de rotación"""
    st.subheader("📈 Clasificación por Velocidad de Rotación")
    st.caption("Distribución de productos según su frecuencia de venta anual")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        categorias_ordenadas = ['A', 'B', 'C', 'D', 'E']
        colores = ['#667eea', '#764ba2', '#f093fb', '#4facfe', '#43e97b']
        
        categoria_counts = df['Categoria'].value_counts().reindex(categorias_ordenadas, fill_value=0)
        
        fig_pie = go.Figure(data=[go.Pie(
            labels=categorias_ordenadas,
            values=categoria_counts.values,
            hole=0.4,
            marker=dict(colors=colores),
            textinfo='label+percent',
            sort=False
        )])
        
        fig_pie.update_layout(title="Proporción de Productos por Categoría", height=400)
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with col2:
        resumen = df.groupby('Categoria').agg({
            'Total_Ventas': 'sum',
            'Valor_Stock_Actual': 'sum',
            'Stock_Sobrante': 'sum'
        }).reindex(categorias_ordenadas, fill_value=0)
        
        resumen_display = pd.DataFrame({
            'Cat.': resumen.index,
            'Ventas Anuales': resumen['Total_Ventas'].apply(formato_numero),
            'Inversión Stock': resumen['Valor_Stock_Actual'].apply(formato_euros),
            'Exceso Stock': resumen['Stock_Sobrante'].apply(formato_euros)
        })
        
        st.dataframe(resumen_display, use_container_width=True, height=250, hide_index=True)
        
        # Botón CNs sobrantes
        mostrar_cns_sobrantes(df, cols)

def mostrar_cns_sobrantes(df, cols):
    """Muestra productos con exceso de stock con toggle"""
    productos_sobrantes = df[df['Stock_Sobrante_Uds'] > 0].copy()
    
    if len(productos_sobrantes) > 0 and cols['cn']:
        mostrar = st.session_state.get('mostrar_cns_sobrante', False)
        
        boton_texto = "🔽 Ocultar CNs" if mostrar else f"👁️ Ver CNs con Exceso ({len(productos_sobrantes)})"
        
        if st.button(boton_texto, use_container_width=True, key="btn_cns"):
            st.session_state.mostrar_cns_sobrante = not mostrar
            st.rerun()
        
        if mostrar:
            productos_sobrantes = productos_sobrantes.sort_values('Stock_Sobrante', ascending=False)
            
            st.markdown("#### 📋 Productos con Exceso de Stock")
            
            display_cols = [cols['cn'], 'Categoria', cols['stock_actual'], 
                           'Stock_Ideal', 'Stock_Sobrante_Uds', 'Stock_Sobrante']
            if cols['descripcion']:
                display_cols.insert(1, cols['descripcion'])
            
            cns_display = productos_sobrantes[display_cols].head(100).copy()
            cns_display['Stock_Sobrante'] = cns_display['Stock_Sobrante'].apply(formato_euros)
            
            st.dataframe(cns_display, use_container_width=True, height=400, hide_index=True)
            
            # Botones descarga
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                cns_txt = "\n".join(productos_sobrantes[cols['cn']].astype(str).tolist())
                st.download_button(
                    "📄 Descargar CNs (TXT)", cns_txt,
                    f"CNs_exceso_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                    use_container_width=True
                )

def grafico_comparativa_stock(df, cols):
    """Gráfico comparativo Stock Actual vs Ideal vs Límite"""
    st.subheader("🎯 Comparativa Stock: Actual vs Ideal vs Límite")
    
    if not cols['cn']:
        st.warning("⚠️ No se encontró columna de código de artículo")
        return
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        analisis = df.groupby('Categoria').agg({
            cols['cn']: 'count',
            cols['stock_actual']: 'sum',
            'Stock_Ideal': 'sum',
            'Stock_Limite': 'sum',
            'Stock_Sobrante_Uds': 'sum',
            'Stock_Faltante_Uds': 'sum',
            'Valor_Stock_Actual': 'sum',
            'Stock_Sobrante': 'sum',
            'Stock_Faltante': 'sum'
        }).reindex(['A', 'B', 'C', 'D', 'E'], fill_value=0)
        
        display_df = pd.DataFrame({
            'Cat.': analisis.index,
            'Refs': analisis[cols['cn']].astype(int),
            'Stock Actual': analisis[cols['stock_actual']].round(0).astype(int),
            'Stock Ideal': analisis['Stock_Ideal'].round(0).astype(int),
            'Exceso (uds)': analisis['Stock_Sobrante_Uds'].round(0).astype(int),
            'Déficit (uds)': analisis['Stock_Faltante_Uds'].round(0).astype(int),
            'Valor Exceso': analisis['Stock_Sobrante'].apply(formato_euros),
            'Valor Déficit': analisis['Stock_Faltante'].apply(formato_euros)
        })
        
        st.dataframe(display_df, use_container_width=True, height=250, hide_index=True)
    
    with col2:
        categorias = ['A', 'B', 'C', 'D', 'E']
        
        fig = go.Figure()
        fig.add_trace(go.Bar(name='Stock Actual', x=categorias, 
                            y=analisis[cols['stock_actual']].values, marker_color='#4169E1'))
        fig.add_trace(go.Bar(name='Stock Ideal', x=categorias,
                            y=analisis['Stock_Ideal'].values, marker_color='#FFD700'))
        fig.add_trace(go.Bar(name='Stock Límite', x=categorias,
                            y=analisis['Stock_Limite'].values, marker_color='#FF6347'))
        
        fig.update_layout(title='Comparativa Stock', barmode='group',
                         yaxis_title='Unidades', height=350)
        st.plotly_chart(fig, use_container_width=True)

def analisis_familias(df, cols):
    """Análisis por familias funcionales"""
    if 'Familia' not in df.columns or cols['cn'] is None:
        st.info("ℹ️ No se detectaron familias funcionales")
        return
    
    st.markdown("---")
    st.subheader("🏪 Análisis por Familias Terapéuticas")
    
    familias_unicas = sorted(df['Familia'].unique().tolist())
    
    # Mostrar tabla resumen
    analisis = df.groupby('Familia').agg({
        cols['cn']: 'count',
        cols['stock_actual']: 'sum',
        'Valor_Stock_Actual': 'sum',
        'Stock_Sobrante': 'sum',
        'Stock_Faltante': 'sum',
        'Total_Ventas': 'sum',
        'Indice_Rotacion': 'mean'
    }).sort_values('Valor_Stock_Actual', ascending=False)
    
    display_df = pd.DataFrame({
        'Familia': analisis.index,
        'Nº Refs': analisis[cols['cn']].astype(int),
        'Stock (uds)': analisis[cols['stock_actual']].round(0).astype(int),
        'Inversión': analisis['Valor_Stock_Actual'].apply(formato_euros),
        'Exceso': analisis['Stock_Sobrante'].apply(formato_euros),
        'Déficit': analisis['Stock_Faltante'].apply(formato_euros),
        'Ventas (uds)': analisis['Total_Ventas'].round(0).astype(int),
        'IR Medio': analisis['Indice_Rotacion'].round(2)
    })
    
    st.dataframe(display_df, use_container_width=True, height=400, hide_index=True)
    
    # Gráfico top familias con exceso
    st.markdown("---")
    st.subheader("🚨 Top Familias con Mayor Exceso")
    
    top_exceso = df.groupby('Familia')['Stock_Sobrante'].sum().sort_values(ascending=False).head(15)
    
    if top_exceso.sum() > 0:
        fig = go.Figure(data=[go.Bar(
            x=top_exceso.values, y=top_exceso.index, orientation='h',
            marker=dict(color=top_exceso.values, colorscale='Reds'),
            text=[formato_euros(v) for v in top_exceso.values],
            textposition='auto'
        )])
        
        fig.update_layout(title="Top 15 Familias - Exceso de Stock",
                         xaxis_title="Valor Exceso (€)", height=500)
        st.plotly_chart(fig, use_container_width=True)

def botones_exportacion(df, cols):
    """Botones para exportar informes"""
    st.markdown("---")
    st.subheader("📥 Exportación de Informes")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Datos Completos', index=False)
        output.seek(0)
        
        st.download_button(
            "📊 Análisis Completo", output,
            f"analisis_completo_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ==================== INTERFAZ PRINCIPAL ====================
def main():
    aplicar_estilos()
    
    st.title("📊 Análisis de Stock Farmacéutico")
    st.markdown("---")
    
    # Sidebar
    st.sidebar.header("⚙️ Configuración")
    
    if st.sidebar.button("🔄 Resetear", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
    
    dias_abierto = st.sidebar.number_input("Días abierto al año", 250, 365, 300, 1)
    
    st.sidebar.markdown("### Parámetros Stock (A y B)")
    col1, col2 = st.sidebar.columns(2)
    with col1:
        stock_min_dias = st.number_input("Mín (días)", 5, 20, 10, 1)
    with col2:
        stock_max_dias = st.number_input("Máx (días)", 15, 40, 20, 1)
    
    dias_cobertura = st.sidebar.slider("Días cobertura ideal", 10, 30, 15, 1)
    margen_seguridad = st.sidebar.slider("Margen seguridad (%)", 0.0, 0.30, 0.0, 0.05)
    
    # Upload
    uploaded_file = st.file_uploader("📁 Cargar archivo Excel", type=['xlsx', 'xls'])
    
    if uploaded_file:
        try:
            # CORRECCIÓN: Crear clave única basada en parámetros para invalidar caché
            cache_key = f"{dias_abierto}_{stock_min_dias}_{stock_max_dias}_{dias_cobertura}_{margen_seguridad}"
            
            # Procesar datos
            df, cols = procesar_excel(uploaded_file, dias_abierto, stock_min_dias, 
                                     stock_max_dias, dias_cobertura, margen_seguridad)
            
            st.success(f"✅ Archivo procesado: {len(df):,} productos".replace(",", "."))
            
            # Mostrar componentes
            mostrar_resumen_ejecutivo(df, cols['stock_actual'])
            st.markdown("---")
            grafico_distribucion_categorias(df, cols)
            st.markdown("---")
            grafico_comparativa_stock(df, cols)
            analisis_familias(df, cols)
            botones_exportacion(df, cols)
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.exception(e)
    else:
        st.info("👆 Cargue un archivo Excel para comenzar")

if __name__ == "__main__":
    main()
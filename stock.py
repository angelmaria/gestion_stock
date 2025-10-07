# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Analisis Stock Farmacia", layout="wide")

st.title("Analisis de Stock Farmaceutico")
st.markdown("---")

# Sidebar para parametros
st.sidebar.header("Configuracion")
dias_abierto = st.sidebar.number_input("Dias abierto al año", min_value=250, max_value=365, value=300, step=1)
dias_cobertura_optimo = st.sidebar.number_input("Dias cobertura optima", min_value=10, max_value=30, value=15, step=1)

# Selectores de tipo de pedido
tipo_pedido = st.sidebar.selectbox(
    "Tipo de pedido (solo para A y B)",
    ["Directo Transfer", "Grupo Compras/Plataforma", "Mayorista Club Genericos", "Especiales"]
)

# Upload Excel
uploaded_file = st.file_uploader("Cargar archivo Excel con datos de ventas", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        # Leer Excel
        df = pd.read_excel(uploaded_file)
        
        st.success(f"Archivo cargado: {len(df)} productos")
        
        # Mostrar columnas detectadas para debug
        with st.expander("Columnas detectadas en el archivo"):
            st.write(df.columns.tolist())
        
        # Buscar columna TOTAL primero
        col_total = None
        for col in df.columns:
            if 'total' in str(col).lower():
                col_total = col
                break
        
        # Si no hay columna TOTAL, buscar columnas de ventas mensuales
        if col_total:
            df['Total_Ventas'] = df[col_total].fillna(0)
            st.info(f"Usando columna '{col_total}' para total de ventas")
        else:
            # Identificar columnas de ventas mensuales
            columnas_ventas = []
            meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 
                    'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre',
                    'en', 'fe', 'mr', 'ab', 'my', 'jn', 'jl', 'ag', 'se', 'oc', 'no', 'di']
            
            for col in df.columns:
                col_lower = str(col).lower()
                if 'ventas' in col_lower or any(mes in col_lower for mes in meses):
                    columnas_ventas.append(col)
            
            if not columnas_ventas:
                st.error("No se encontraron columnas de ventas mensuales ni columna TOTAL")
                st.stop()
            
            # Calcular total ventas anuales
            df['Total_Ventas'] = df[columnas_ventas].fillna(0).sum(axis=1)
            st.info(f"Calculando total desde {len(columnas_ventas)} columnas mensuales")
        
        # Calcular ventas diarias promedio
        df['Vtas_Dia'] = df['Total_Ventas'] / dias_abierto
        
        # Funcion para categorizar productos
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
        
        # Configuracion de stocks segun tipo de pedido y categoria
        def calcular_stocks(row, tipo_pedido, dias_cobertura):
            cat = row['Categoria']
            vtas_dia = row['Vtas_Dia']
            
            if cat == 'A':
                if tipo_pedido == 'Mayorista Club Genericos':
                    stock_min = vtas_dia * 2
                    stock_max = vtas_dia * 3
                    stock_opt = vtas_dia * 2.5
                elif tipo_pedido == 'Directo Transfer':
                    stock_min = vtas_dia * 8
                    stock_max = vtas_dia * 15
                    stock_opt = vtas_dia * dias_cobertura
                elif tipo_pedido == 'Grupo Compras/Plataforma':
                    stock_min = vtas_dia * 3
                    stock_max = vtas_dia * 15
                    stock_opt = vtas_dia * dias_cobertura
                else:  # Especiales
                    stock_min = vtas_dia * 1
                    stock_max = vtas_dia * 2
                    stock_opt = vtas_dia * 1.5
                    
            elif cat == 'B':
                if tipo_pedido == 'Mayorista Club Genericos':
                    stock_min = vtas_dia * 2
                    stock_max = vtas_dia * 3
                    stock_opt = vtas_dia * 2.5
                elif tipo_pedido == 'Directo Transfer':
                    stock_min = vtas_dia * 8
                    stock_max = vtas_dia * 15
                    stock_opt = vtas_dia * dias_cobertura
                elif tipo_pedido == 'Grupo Compras/Plataforma':
                    stock_min = vtas_dia * 3
                    stock_max = vtas_dia * 9
                    stock_opt = vtas_dia * min(dias_cobertura, 9)
                else:  # Especiales
                    stock_min = vtas_dia * 1
                    stock_max = vtas_dia * 2
                    stock_opt = vtas_dia * 1.5
                    
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
        
        # Aplicar calculos
        df[['Stock_Min_Calc', 'Stock_Max_Calc', 'Stock_Opt_Calc']] = df.apply(
            lambda row: calcular_stocks(row, tipo_pedido, dias_cobertura_optimo), axis=1
        )
        
        # Buscar columnas existentes en el Excel
        col_stock_actual = None
        col_pvp = None
        col_minf = None
        col_maxf = None
        col_cn = None
        col_descripcion = None
        
        for col in df.columns:
            col_lower = str(col).lower()
            if 'stock' in col_lower and 'actual' in col_lower:
                col_stock_actual = col
            elif col_lower == 'pvp':
                col_pvp = col
            elif col_lower == 'minf':
                col_minf = col
            elif col_lower == 'maxf':
                col_maxf = col
            elif col_lower == 'cn':
                col_cn = col
            elif 'descripcion' in col_lower:
                col_descripcion = col
        
        # Limpiar valores de PVP si tiene simbolo de euro
        if col_pvp:
            df[col_pvp] = df[col_pvp].astype(str).str.replace('€', '').str.replace(',', '.').str.strip()
            df[col_pvp] = pd.to_numeric(df[col_pvp], errors='coerce').fillna(0)
        
        if col_stock_actual and col_pvp:
            df['Valor_Stock_Actual'] = df[col_stock_actual] * df[col_pvp]
            df['Valor_Stock_Optimo'] = df['Stock_Opt_Calc'] * df[col_pvp]
            
            # Calcular sobrante y faltante
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
        
        # Mostrar resumen
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Productos", len(df))
        with col2:
            if 'Valor_Stock_Actual' in df.columns:
                st.metric("Valor Stock Actual", f"{df['Valor_Stock_Actual'].sum():,.2f} euros")
        with col3:
            if 'Stock_Sobrante' in df.columns:
                st.metric("Stock Sobrante", f"{df['Stock_Sobrante'].sum():,.2f} euros")
        with col4:
            if 'Stock_Faltante' in df.columns:
                st.metric("Stock Faltante", f"{df['Stock_Faltante'].sum():,.2f} euros")
        
        st.markdown("---")
        
        # Distribucion por categorias
        st.subheader("Distribucion por Categorias")
        col1, col2 = st.columns(2)
        
        with col1:
            categoria_counts = df['Categoria'].value_counts().sort_index()
            st.bar_chart(categoria_counts)
        
        with col2:
            resumen_cat = df.groupby('Categoria').agg({
                'Total_Ventas': 'sum',
                'Valor_Stock_Actual': 'sum' if 'Valor_Stock_Actual' in df.columns else 'count'
            }).round(2)
            st.dataframe(resumen_cat, use_container_width=True)
        
        # Filtros
        st.sidebar.markdown("---")
        st.sidebar.subheader("Filtros")
        categorias_seleccionadas = st.sidebar.multiselect(
            "Filtrar por categoria",
            options=['A', 'B', 'C', 'D', 'E'],
            default=['A', 'B', 'C', 'D', 'E']
        )
        
        df_filtrado = df[df['Categoria'].isin(categorias_seleccionadas)]
        
        # Tabla de resultados
        st.subheader("Detalle de Productos")
        
        # Seleccionar columnas a mostrar
        columnas_mostrar = []
        
        if col_cn:
            columnas_mostrar.append(col_cn)
        if col_descripcion:
            columnas_mostrar.append(col_descripcion)
        
        columnas_mostrar.extend(['Categoria', 'Total_Ventas', 'Vtas_Dia'])
        
        if col_minf:
            columnas_mostrar.append(col_minf)
        if col_maxf:
            columnas_mostrar.append(col_maxf)
        if col_stock_actual:
            columnas_mostrar.append(col_stock_actual)
        
        columnas_mostrar.extend(['Stock_Min_Calc', 'Stock_Max_Calc', 'Stock_Opt_Calc'])
        
        if 'Reposicion' in df_filtrado.columns:
            columnas_mostrar.append('Reposicion')
        if 'Valor_Stock_Actual' in df_filtrado.columns:
            columnas_mostrar.extend(['Valor_Stock_Actual', 'Stock_Sobrante', 'Stock_Faltante'])
        
        # Filtrar solo columnas existentes
        columnas_mostrar = [col for col in columnas_mostrar if col in df_filtrado.columns]
        
        st.dataframe(
            df_filtrado[columnas_mostrar].round(2),
            use_container_width=True,
            height=400
        )
        
        # Boton de descarga
        st.markdown("---")
        
        # Preparar DataFrame para descarga con todas las columnas
        df_descarga = df.copy()
        
        csv = df_descarga.to_csv(index=False, encoding='utf-8-sig', decimal=',', sep=';')
        
        st.download_button(
            label="Descargar analisis completo (CSV)",
            data=csv.encode('utf-8-sig'),
            file_name=f"analisis_stock_farmacia_{pd.Timestamp.now().strftime('%Y%m%d')}.csv",
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
       - Columnas de ventas mensuales (Ventas_Enero, Ventas_Febrero, etc.) O
       - Una columna TOTAL con las ventas anuales
    2. Columnas requeridas: CN, Descripcion, PVP, Stock Actual
    3. Columnas opcionales: MinF, MaxF
    4. Ajusta los parametros en el panel lateral
    """)
"""
Módulo de análisis de datos de ventas.
Contiene la clase AnalizadorVentas que procesa datos de Excel y genera métricas.
"""

import pandas as pd
from datetime import datetime
from typing import Dict, Any


class AnalizadorVentas:
    """
    Clase para analizar datos de ventas de un cliente específico.
    """
    
    def __init__(self, ruta_excel: str):
        """
        Inicializa el analizador con los datos del Excel.
        """
        self.df = pd.read_excel(ruta_excel)
        
        # Limpiar nombres de columnas
        self.df.columns = [col.strip() for col in self.df.columns]
        
        # Convertir fecha
        self.df['FECHA'] = pd.to_datetime(self.df['FECHA'])
        
        # CORRECCIÓN IMPORTANTE: Calcular el monto real facturado en dinero
        # La columna FACTURADO contiene unidades, no dinero
        # El monto real = FACTURADO (unidades) × PRECIO UNITARIO
        self.df['MONTO_FACTURADO'] = self.df['FACTURADO'] * self.df['PRECIO UNITARIO']
        
        # Obtener nombre del cliente
        self.cliente = self._obtener_nombre_cliente()
        
        # Limpiar datos
        self._limpiar_datos()
    
    def _obtener_nombre_cliente(self) -> str:
        """
        Obtiene el nombre del cliente desde la primera fila de datos.
        
        Returns:
            str: Nombre del cliente
        """
        if 'CLIENTE' in self.df.columns and len(self.df) > 0:
            return self.df['CLIENTE'].iloc[0]
        return "Cliente Desconocido"
    
    def _limpiar_datos(self):
        """
        Limpia y prepara los datos para el análisis.
        - Convierte fechas al formato correcto
        - Elimina filas vacías o con datos inválidos
        - Asegura que los valores numéricos sean del tipo correcto
        - Limpia nombres de columnas (elimina espacios extra)
        """
        # Limpiar nombres de columnas (eliminar espacios al inicio y final)
        self.df.columns = [col.strip() for col in self.df.columns]
        
        # Convertir FECHA a datetime si no lo es
        if 'FECHA' in self.df.columns:
            self.df['FECHA'] = pd.to_datetime(self.df['FECHA'], errors='coerce')
        
        # Asegurar que columnas numéricas sean del tipo correcto
        columnas_numericas = ['CANT', 'PESO NETO', 'PRECIO UNITARIO', 'FACTURADO', 'PESO TOTAL', 'MONTO_FACTURADO']
        for col in columnas_numericas:
            if col in self.df.columns:
                self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
        
        # Eliminar filas con valores nulos en columnas críticas
        self.df = self.df.dropna(subset=['CODIGO', 'NOMBRE'])
        
        # Crear columna de mes-año para análisis temporal
        if 'FECHA' in self.df.columns:
            self.df['MES_ANIO'] = self.df['FECHA'].dt.to_period('M')
    
    def resumen_general(self) -> Dict[str, Any]:
        """
        Genera un resumen general de las ventas del cliente.
        
        Returns:
            Dict con métricas generales: total facturado, órdenes, productos únicos, etc.
        """
        resumen = {
            'cliente': self.cliente,
            'total_facturado': self.df['MONTO_FACTURADO'].sum(),
            'total_ordenes': self.df['OV'].nunique(),
            'productos_unicos': self.df['CODIGO'].nunique(),
            'peso_total': self.df['PESO TOTAL'].sum(),
            'cantidad_total': self.df['CANT'].sum(),
            'ticket_promedio': self.df.groupby('OV')['MONTO_FACTURADO'].sum().mean(),
            'peso_promedio_orden': self.df.groupby('OV')['PESO TOTAL'].sum().mean(),
            'precio_promedio_kg': self.df['MONTO_FACTURADO'].sum() / self.df['PESO TOTAL'].sum() if self.df['PESO TOTAL'].sum() > 0 else 0
        }
        return resumen
    
    def top_productos_cantidad(self, top_n: int = 10) -> pd.DataFrame:
        """
        Obtiene los productos más vendidos por cantidad.
        
        Args:
            top_n (int): Número de productos a retornar
            
        Returns:
            DataFrame con los top productos por cantidad
        """
        top = self.df.groupby(['CODIGO', 'NOMBRE']).agg({
            'CANT': 'sum',
            'MONTO_FACTURADO': 'sum',
            'PESO TOTAL': 'sum'
        }).reset_index()
        
        top = top.sort_values('CANT', ascending=False).head(top_n)
        
        # CORRECCIÓN: Asegurar que las columnas estén en el orden correcto
        top = top[['CODIGO', 'NOMBRE', 'CANT', 'MONTO_FACTURADO', 'PESO TOTAL']]
        top.columns = ['Código', 'Nombre', 'Cantidad Total', 'Facturación Total', 'Peso Total']
        return top
    
    def top_productos_facturacion(self, top_n: int = 10) -> pd.DataFrame:
        """
        Obtiene los productos con mayor facturación.
        
        Args:
            top_n (int): Número de productos a retornar
            
        Returns:
            DataFrame con los top productos por facturación
        """
        top = self.df.groupby(['CODIGO', 'NOMBRE']).agg({
            'MONTO_FACTURADO': 'sum',
            'CANT': 'sum',
            'PESO TOTAL': 'sum'
        }).reset_index()
        
        top = top.sort_values('MONTO_FACTURADO', ascending=False).head(top_n)
        
        # CORRECCIÓN: Asegurar que las columnas estén en el orden correcto
        top = top[['CODIGO', 'NOMBRE', 'MONTO_FACTURADO', 'CANT', 'PESO TOTAL']]
        top.columns = ['Código', 'Nombre', 'Facturación Total', 'Cantidad Total', 'Peso Total']
        return top
    
    def analisis_categorias(self) -> pd.DataFrame:
        """
        Analiza las ventas por categoría de producto.
        
        Returns:
            DataFrame con métricas por categoría
        """
        if 'CATEGORIA' not in self.df.columns:
            return pd.DataFrame()
        
        categorias = self.df.groupby('CATEGORIA').agg({
            'MONTO_FACTURADO': 'sum',
            'CANT': 'sum',
            'PESO TOTAL': 'sum',
            'OV': 'nunique'
        }).reset_index()
        
        # Calcular porcentaje de facturación
        total_facturado = categorias['MONTO_FACTURADO'].sum()
        if total_facturado > 0:
            categorias['% Facturación'] = (categorias['MONTO_FACTURADO'] / total_facturado * 100).fillna(0).round(2)
        else:
            categorias['% Facturación'] = 0
        
        categorias = categorias.sort_values('MONTO_FACTURADO', ascending=False)
        categorias.columns = ['Categoría', 'Facturación', 'Cantidad', 'Peso Total', 'Órdenes', '% Facturación']
        
        return categorias
    
    def ventas_por_mes(self) -> pd.DataFrame:
        """
        Analiza las ventas agrupadas por mes.
        
        Returns:
            DataFrame con ventas mensuales
        """
        if 'MES_ANIO' not in self.df.columns:
            return pd.DataFrame()
        
        ventas_mes = self.df.groupby('MES_ANIO').agg({
            'MONTO_FACTURADO': 'sum',
            'CANT': 'sum',
            'OV': 'nunique',
            'PESO TOTAL': 'sum'
        }).reset_index()
        
        ventas_mes['MES_ANIO'] = ventas_mes['MES_ANIO'].astype(str)
        ventas_mes.columns = ['Mes', 'Facturación', 'Cantidad', 'Órdenes', 'Peso Total']
        
        return ventas_mes
    
    def productos_precio_alto_kg(self, top_n: int = 10) -> pd.DataFrame:
        """
        Obtiene los productos con mayor precio por kilogramo.
        
        Args:
            top_n (int): Número de productos a retornar
            
        Returns:
            DataFrame con productos de mayor precio/kg
        """
        # Calcular precio por kg para cada producto
        productos = self.df.groupby(['CODIGO', 'NOMBRE']).agg({
            'MONTO_FACTURADO': 'sum',
            'PESO TOTAL': 'sum',
            'CANT': 'sum'
        }).reset_index()
        
        # Evitar división por cero
        productos = productos[productos['PESO TOTAL'] > 0]
        productos['Precio/Kg'] = productos['MONTO_FACTURADO'] / productos['PESO TOTAL']
        
        productos = productos.sort_values('Precio/Kg', ascending=False).head(top_n)
        productos = productos[['CODIGO', 'NOMBRE', 'Precio/Kg', 'MONTO_FACTURADO', 'PESO TOTAL']]
        productos.columns = ['Código', 'Nombre', 'Precio/Kg', 'Facturación Total', 'Peso Total']
        
        return productos
    
    def analisis_pareto(self) -> Dict[str, Any]:
        """
        Aplica la Ley de Pareto (80/20) a los productos.
        Identifica qué porcentaje de productos genera el 80% de la facturación.
        
        Returns:
            Dict con análisis de Pareto y DataFrame con productos acumulados
        """
        # Agrupar por producto y calcular facturación total
        productos = self.df.groupby(['CODIGO', 'NOMBRE']).agg({
            'MONTO_FACTURADO': 'sum'
        }).reset_index()
        
        # Ordenar por facturación descendente
        productos = productos.sort_values('MONTO_FACTURADO', ascending=False)
        
        # Calcular facturación acumulada y porcentaje acumulado
        productos['Facturación Acumulada'] = productos['MONTO_FACTURADO'].cumsum()
        total_facturado = productos['MONTO_FACTURADO'].sum()
        productos['% Acumulado'] = (productos['Facturación Acumulada'] / total_facturado * 100).fillna(0).round(2)
        productos['% Individual'] = (productos['MONTO_FACTURADO'] / total_facturado * 100).fillna(0).round(2)
        
        # Identificar productos que representan el 80%
        productos_80 = productos[productos['% Acumulado'] <= 80]
        total_productos = len(productos)
        productos_80_count = len(productos_80)
        porcentaje_productos_80 = round((productos_80_count / total_productos * 100), 2) if total_productos > 0 else 0
        
        # Identificar productos que representan el 20% restante
        productos_20 = productos[productos['% Acumulado'] > 80]
        productos_20_count = len(productos_20)
        porcentaje_productos_20 = round((productos_20_count / total_productos * 100), 2) if total_productos > 0 else 0
        
        # Facturación del 80% y 20%
        facturacion_80 = productos_80['MONTO_FACTURADO'].sum()
        facturacion_20 = productos_20['MONTO_FACTURADO'].sum()
        
        # Renombrar columnas para el reporte
        productos.columns = ['Código', 'Nombre', 'Facturación', 'Facturación Acumulada', '% Acumulado', '% Individual']
        
        resultado = {
            'total_productos': total_productos,
            'productos_80_count': productos_80_count,
            'porcentaje_productos_80': porcentaje_productos_80,
            'facturacion_80': facturacion_80,
            'productos_20_count': productos_20_count,
            'porcentaje_productos_20': porcentaje_productos_20,
            'facturacion_20': facturacion_20,
            'dataframe': productos
        }
        
        return resultado
    
    def crecimiento_mensual(self) -> pd.DataFrame:
        """
        Calcula el crecimiento porcentual mes a mes.
        
        Returns:
            DataFrame con crecimiento mensual
        """
        if 'MES_ANIO' not in self.df.columns:
            return pd.DataFrame()
        
        ventas_mes = self.df.groupby('MES_ANIO').agg({
            'MONTO_FACTURADO': 'sum'
        }).reset_index()
        
        ventas_mes = ventas_mes.sort_values('MES_ANIO')
        
        # Calcular crecimiento porcentual
        ventas_mes['Crecimiento %'] = ventas_mes['MONTO_FACTURADO'].pct_change() * 100
        ventas_mes['Crecimiento %'] = ventas_mes['Crecimiento %'].fillna(0).round(2)
        
        # Calcular diferencia absoluta
        ventas_mes['Diferencia $'] = ventas_mes['MONTO_FACTURADO'].diff()
        
        ventas_mes['MES_ANIO'] = ventas_mes['MES_ANIO'].astype(str)
        ventas_mes.columns = ['Mes', 'Facturación', 'Crecimiento %', 'Diferencia $']
        
        return ventas_mes
    
    def frecuencia_compra(self) -> Dict[str, Any]:
        """
        Calcula la frecuencia de compra del cliente.
        
        Returns:
            Dict con métricas de frecuencia de compra
        """
        if 'FECHA' not in self.df.columns or self.df['FECHA'].isna().all():
            return {
                'dias_entre_compras': 0,
                'compras_por_mes': 0,
                'primera_compra': 'N/A',
                'ultima_compra': 'N/A',
                'dias_totales': 0,
                'total_ordenes': 0
            }
        
        # Obtener fechas únicas de órdenes
        ordenes = self.df.groupby('OV')['FECHA'].min().sort_values()
        
        if len(ordenes) < 2:
            return {
                'dias_entre_compras': 0,
                'compras_por_mes': len(ordenes),
                'primera_compra': ordenes.iloc[0].strftime('%Y-%m-%d') if len(ordenes) > 0 else 'N/A',
                'ultima_compra': ordenes.iloc[0].strftime('%Y-%m-%d') if len(ordenes) > 0 else 'N/A',
                'dias_totales': 0,
                'total_ordenes': len(ordenes)
            }
        
        # Calcular diferencias entre fechas
        diferencias = ordenes.diff().dt.days.dropna()
        dias_promedio = diferencias.mean()
        
        # Calcular métricas
        primera_compra = ordenes.iloc[0]
        ultima_compra = ordenes.iloc[-1]
        dias_totales = (ultima_compra - primera_compra).days
        compras_por_mes = (len(ordenes) / (dias_totales / 30.44)) if dias_totales > 0 else 0
        
        return {
            'dias_entre_compras': round(float(dias_promedio), 1) if not pd.isna(dias_promedio) else 0,
            'compras_por_mes': round(compras_por_mes, 2),
            'primera_compra': primera_compra.strftime('%Y-%m-%d'),
            'ultima_compra': ultima_compra.strftime('%Y-%m-%d'),
            'dias_totales': dias_totales,
            'total_ordenes': len(ordenes)
        }
    
    def obtener_dataframe_completo(self) -> pd.DataFrame:
        """
        Retorna el DataFrame completo procesado.
        
        Returns:
            DataFrame con todos los datos
        """
        return self.df.copy()
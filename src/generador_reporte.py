"""
Módulo para generar reportes en Excel con gráficos y tablas.
Utiliza xlsxwriter para crear archivos Excel con formato profesional.
"""

import pandas as pd
import xlsxwriter
from typing import Dict, Any
from datetime import datetime


class GeneradorReporte:
    """
    Clase para generar reportes en Excel con múltiples hojas, tablas y gráficos.
    """
    
    def __init__(self, ruta_salida: str):
        """
        Inicializa el generador de reportes.
        
        Args:
            ruta_salida (str): Ruta donde se guardará el archivo Excel
        """
        self.ruta_salida = ruta_salida
        self.workbook = None
        self.formatos = {}
    
    def _crear_formatos(self):
        """
        Crea los formatos de celda que se usarán en el reporte.
        """
        self.formatos = {
            'titulo': self.workbook.add_format({
                'bold': True,
                'font_size': 16,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#2C3E50',
                'font_color': 'white',
                'border': 1
            }),
            'subtitulo': self.workbook.add_format({
                'bold': True,
                'font_size': 12,
                'bg_color': '#34495E',
                'font_color': 'white',
                'border': 1
            }),
            'encabezado': self.workbook.add_format({
                'bold': True,
                'bg_color': '#3498DB',
                'font_color': 'white',
                'border': 1,
                'align': 'center'
            }),
            'moneda': self.workbook.add_format({
                'num_format': '$ #,##0.00',
                'border': 1
            }),
            'numero': self.workbook.add_format({
                'num_format': '#,##0.00',
                'border': 1
            }),
            'entero': self.workbook.add_format({
                'num_format': '#,##0',
                'border': 1
            }),
            'porcentaje': self.workbook.add_format({
                'num_format': '0.00%',
                'border': 1
            }),
            'normal': self.workbook.add_format({
                'border': 1
            })
        }
    
    def generar_reporte_completo(self, resumen: Dict[str, Any], 
                                  top_cantidad: pd.DataFrame,
                                  top_facturacion: pd.DataFrame,
                                  categorias: pd.DataFrame,
                                  ventas_mes: pd.DataFrame,
                                  precio_kg: pd.DataFrame,
                                  df_completo: pd.DataFrame,
                                  pareto: Dict[str, Any],
                                  crecimiento: pd.DataFrame,
                                  frecuencia: Dict[str, Any]):
        """
        Genera el reporte completo en Excel con todas las hojas y gráficos.
        
        Args:
            resumen: Diccionario con métricas generales
            top_cantidad: DataFrame con top productos por cantidad
            top_facturacion: DataFrame con top productos por facturación
            categorias: DataFrame con análisis de categorías
            ventas_mes: DataFrame con ventas mensuales
            precio_kg: DataFrame con productos de mayor precio/kg
            df_completo: DataFrame con todos los datos originales
            pareto: Diccionario con análisis de Pareto
            crecimiento: DataFrame con crecimiento mensual
            frecuencia: Diccionario con frecuencia de compra
        """
        self.workbook = xlsxwriter.Workbook(self.ruta_salida, {'nan_inf_to_errors': True})
        self._crear_formatos()
        
        # Crear hojas
        self._crear_hoja_resumen(resumen, frecuencia)
        self._crear_hoja_pareto(pareto)
        self._crear_hoja_productos(top_cantidad, top_facturacion, precio_kg)
        self._crear_hoja_categorias(categorias)
        self._crear_hoja_temporal(ventas_mes, crecimiento)
        self._crear_hoja_datos_completos(df_completo)
        
        self.workbook.close()
        print(f"✅ Reporte generado exitosamente: {self.ruta_salida}")
    
    def _crear_hoja_resumen(self, resumen: Dict[str, Any], frecuencia: Dict[str, Any]):
        """
        Crea la hoja de resumen general con KPIs principales.
        """
        sheet = self.workbook.add_worksheet('Resumen General')
        sheet.set_column('A:A', 35)
        sheet.set_column('B:B', 20)
        
        # Título
        sheet.merge_range('A1:B1', f'REPORTE DE VENTAS - {resumen["cliente"]}', self.formatos['titulo'])
        sheet.write('A2', f'Generado: {datetime.now().strftime("%Y-%m-%d %H:%M")}', self.formatos['normal'])
        
        # KPIs principales
        row = 4
        sheet.write(row, 0, 'INDICADORES PRINCIPALES', self.formatos['subtitulo'])
        sheet.write(row, 1, '', self.formatos['subtitulo'])
        
        # Limpiar valores NaN e infinitos del resumen
        for key, value in resumen.items():
            if isinstance(value, float) and (pd.isna(value) or value in [float('inf'), float('-inf')]):
                resumen[key] = 0
        
        kpis = [
            ('Total Facturado', resumen['total_facturado'], 'moneda'),
            ('Total de Órdenes', resumen['total_ordenes'], 'entero'),
            ('Productos Únicos', resumen['productos_unicos'], 'entero'),
            ('Peso Total (kg)', resumen['peso_total'], 'numero'),
            ('Cantidad Total', resumen['cantidad_total'], 'entero'),
            ('Ticket Promedio por Orden', resumen['ticket_promedio'], 'moneda'),
            ('Peso Promedio por Orden (kg)', resumen['peso_promedio_orden'], 'numero'),
            ('Precio Promedio por Kg', resumen['precio_promedio_kg'], 'moneda'),
        ]
        
        row += 1
        for nombre, valor, formato in kpis:
            sheet.write(row, 0, nombre, self.formatos['normal'])
            sheet.write(row, 1, valor, self.formatos[formato])
            row += 1
        
        # Frecuencia de compra
        row += 2
        sheet.write(row, 0, 'FRECUENCIA DE COMPRA', self.formatos['subtitulo'])
        sheet.write(row, 1, '', self.formatos['subtitulo'])
        
        frecuencia_kpis = [
            ('Primera Compra', frecuencia['primera_compra'], 'normal'),
            ('Última Compra', frecuencia['ultima_compra'], 'normal'),
            ('Días Totales de Relación', frecuencia['dias_totales'], 'entero'),
            ('Total de Órdenes', frecuencia['total_ordenes'], 'entero'),
            ('Días Promedio entre Compras', frecuencia['dias_entre_compras'], 'numero'),
            ('Compras Promedio por Mes', frecuencia['compras_por_mes'], 'numero'),
        ]
        
        row += 1
        for nombre, valor, formato in frecuencia_kpis:
            sheet.write(row, 0, nombre, self.formatos['normal'])
            sheet.write(row, 1, valor, self.formatos[formato])
            row += 1
    
    def _crear_hoja_pareto(self, pareto: Dict[str, Any]):
        """
        Crea la hoja de análisis de Pareto (80/20).
        """
        sheet = self.workbook.add_worksheet('Ley de Pareto (80-20)')
        sheet.set_column('A:F', 18)
        
        # Título
        sheet.merge_range('A1:F1', 'ANÁLISIS DE PARETO (LEY 80/20)', self.formatos['titulo'])
        
        # Resumen de Pareto
        row = 3
        sheet.write(row, 0, 'RESUMEN LEY DE PARETO', self.formatos['subtitulo'])
        sheet.merge_range(row, 1, row, 5, '', self.formatos['subtitulo'])
        
        row += 1
        sheet.write(row, 0, 'Total de Productos:', self.formatos['normal'])
        sheet.write(row, 1, pareto['total_productos'], self.formatos['entero'])
        
        row += 2
        sheet.write(row, 0, '🔵 PRODUCTOS VITALES (80% de facturación):', self.formatos['encabezado'])
        sheet.merge_range(row, 1, row, 5, '', self.formatos['encabezado'])
        
        row += 1
        sheet.write(row, 0, 'Cantidad de Productos:', self.formatos['normal'])
        sheet.write(row, 1, pareto['productos_80_count'], self.formatos['entero'])
        sheet.write(row, 2, f"{pareto['porcentaje_productos_80']}% del total", self.formatos['normal'])
        
        row += 1
        sheet.write(row, 0, 'Facturación Generada:', self.formatos['normal'])
        sheet.write(row, 1, pareto['facturacion_80'], self.formatos['moneda'])
        sheet.write(row, 2, '80% del total', self.formatos['normal'])
        
        row += 2
        sheet.write(row, 0, '🔴 PRODUCTOS TRIVIALES (20% de facturación):', self.formatos['encabezado'])
        sheet.merge_range(row, 1, row, 5, '', self.formatos['encabezado'])
        
        row += 1
        sheet.write(row, 0, 'Cantidad de Productos:', self.formatos['normal'])
        sheet.write(row, 1, pareto['productos_20_count'], self.formatos['entero'])
        sheet.write(row, 2, f"{pareto['porcentaje_productos_20']}% del total", self.formatos['normal'])
        
        row += 1
        sheet.write(row, 0, 'Facturación Generada:', self.formatos['normal'])
        sheet.write(row, 1, pareto['facturacion_20'], self.formatos['moneda'])
        sheet.write(row, 2, '20% del total', self.formatos['normal'])
        
        row += 3
        sheet.write(row, 0, '💡 INTERPRETACIÓN:', self.formatos['subtitulo'])
        sheet.merge_range(row, 1, row, 5, '', self.formatos['subtitulo'])
        
        row += 1
        interpretacion = f"El {pareto['porcentaje_productos_80']}% de tus productos ({pareto['productos_80_count']} productos) generan el 80% de tu facturación. Enfócate en estos productos vitales."
        sheet.merge_range(row, 0, row, 5, interpretacion, self.formatos['normal'])
        
        # Tabla detallada de productos con Pareto
        row += 3
        sheet.merge_range(row, 0, row, 5, 'DETALLE DE PRODUCTOS (ORDENADOS POR FACTURACIÓN)', self.formatos['titulo'])
        row += 1
        
        df_pareto = pareto['dataframe'].head(50)  # Mostrar top 50
        self._escribir_dataframe(sheet, df_pareto, row, 0)
        
        # Gráfico de Pareto
        if len(df_pareto) > 0:
            chart = self.workbook.add_chart({'type': 'column'})
            chart.add_series({
                'name': 'Facturación',
                'categories': ['Ley de Pareto (80-20)', row + 1, 1, row + min(20, len(df_pareto)), 1],
                'values': ['Ley de Pareto (80-20)', row + 1, 2, row + min(20, len(df_pareto)), 2],
                'y2_axis': False,
            })
            
            # Agregar línea de % acumulado
            chart.add_series({
                'name': '% Acumulado',
                'categories': ['Ley de Pareto (80-20)', row + 1, 1, row + min(20, len(df_pareto)), 1],
                'values': ['Ley de Pareto (80-20)', row + 1, 4, row + min(20, len(df_pareto)), 4],
                'y2_axis': True,
                'line': {'color': 'red', 'width': 2},
            })
            
            chart.set_title({'name': 'Diagrama de Pareto (Top 20 Productos)'})
            chart.set_x_axis({'name': 'Producto'})
            chart.set_y_axis({'name': 'Facturación ($)'})
            chart.set_y2_axis({'name': '% Acumulado'})
            chart.set_size({'width': 900, 'height': 500})
            chart.set_legend({'position': 'bottom'})
            
            sheet.insert_chart(row + min(20, len(df_pareto)) + 2, 0, chart)
    
    def _crear_hoja_productos(self, top_cantidad: pd.DataFrame, 
                               top_facturacion: pd.DataFrame,
                               precio_kg: pd.DataFrame):
        """
        Crea la hoja de análisis de productos con tablas y gráficos.
        """
        sheet = self.workbook.add_worksheet('Análisis de Productos')
        sheet.set_column('A:E', 15)
        
        row = 0
        
        # Top productos por cantidad
        sheet.merge_range(row, 0, row, 4, 'TOP 10 PRODUCTOS POR CANTIDAD', self.formatos['titulo'])
        row += 1
        self._escribir_dataframe(sheet, top_cantidad, row, 0)
        
        # Gráfico de barras para top cantidad
        if len(top_cantidad) > 0:
            chart = self.workbook.add_chart({'type': 'column'})
            chart.add_series({
                'name': 'Cantidad Total',
                'categories': ['Análisis de Productos', row + 1, 1, row + len(top_cantidad), 1],
                'values': ['Análisis de Productos', row + 1, 2, row + len(top_cantidad), 2],
            })
            chart.set_title({'name': 'Top 10 Productos por Cantidad'})
            chart.set_x_axis({'name': 'Producto'})
            chart.set_y_axis({'name': 'Cantidad'})
            chart.set_size({'width': 720, 'height': 400})
            sheet.insert_chart(row + len(top_cantidad) + 2, 0, chart)
        
        row += len(top_cantidad) + 20
        
        # Top productos por facturación
        sheet.merge_range(row, 0, row, 4, 'TOP 10 PRODUCTOS POR FACTURACIÓN', self.formatos['titulo'])
        row += 1
        self._escribir_dataframe(sheet, top_facturacion, row, 0)
        
        row += len(top_facturacion) + 2
        
        # Productos con mayor precio/kg
        sheet.merge_range(row, 0, row, 4, 'TOP 10 PRODUCTOS POR PRECIO/KG', self.formatos['titulo'])
        row += 1
        self._escribir_dataframe(sheet, precio_kg, row, 0)
    
    def _crear_hoja_categorias(self, categorias: pd.DataFrame):
        """
        Crea la hoja de análisis de categorías con gráfico de torta.
        """
        sheet = self.workbook.add_worksheet('Análisis de Categorías')
        sheet.set_column('A:F', 18)
        
        # Título
        sheet.merge_range('A1:F1', 'ANÁLISIS POR CATEGORÍA', self.formatos['titulo'])
        
        # Tabla de categorías
        self._escribir_dataframe(sheet, categorias, 2, 0)
        
        # Gráfico de torta
        if len(categorias) > 0:
            chart = self.workbook.add_chart({'type': 'pie'})
            chart.add_series({
                'name': 'Facturación por Categoría',
                'categories': ['Análisis de Categorías', 3, 0, 2 + len(categorias), 0],
                'values': ['Análisis de Categorías', 3, 1, 2 + len(categorias), 1],
                'data_labels': {'percentage': True},
            })
            chart.set_title({'name': 'Distribución de Facturación por Categoría'})
            chart.set_size({'width': 600, 'height': 400})
            sheet.insert_chart(len(categorias) + 4, 0, chart)
    
    def _crear_hoja_temporal(self, ventas_mes: pd.DataFrame, crecimiento: pd.DataFrame):
        """
        Crea la hoja de análisis temporal con gráfico de línea y crecimiento.
        """
        sheet = self.workbook.add_worksheet('Análisis Temporal')
        sheet.set_column('A:E', 18)
        
        # Título
        sheet.merge_range('A1:E1', 'VENTAS POR MES', self.formatos['titulo'])
        
        # Tabla de ventas mensuales
        self._escribir_dataframe(sheet, ventas_mes, 2, 0)
        
        # Gráfico de línea
        if len(ventas_mes) > 0:
            chart = self.workbook.add_chart({'type': 'line'})
            chart.add_series({
                'name': 'Facturación Mensual',
                'categories': ['Análisis Temporal', 3, 0, 2 + len(ventas_mes), 0],
                'values': ['Análisis Temporal', 3, 1, 2 + len(ventas_mes), 1],
            })
            chart.set_title({'name': 'Tendencia de Facturación Mensual'})
            chart.set_x_axis({'name': 'Mes'})
            chart.set_y_axis({'name': 'Facturación ($)'})
            chart.set_size({'width': 720, 'height': 400})
            sheet.insert_chart(len(ventas_mes) + 4, 0, chart)
        
        # Crecimiento mensual
        row = len(ventas_mes) + 25
        sheet.merge_range(row, 0, row, 4, 'CRECIMIENTO MENSUAL (%)', self.formatos['titulo'])
        row += 1
        
        if len(crecimiento) > 0:
            self._escribir_dataframe(sheet, crecimiento, row, 0)
            
            # Gráfico de crecimiento
            chart2 = self.workbook.add_chart({'type': 'column'})
            chart2.add_series({
                'name': 'Crecimiento %',
                'categories': ['Análisis Temporal', row + 1, 0, row + len(crecimiento), 0],
                'values': ['Análisis Temporal', row + 1, 2, row + len(crecimiento), 2],
            })
            chart2.set_title({'name': 'Crecimiento Mensual (%)'})
            chart2.set_x_axis({'name': 'Mes'})
            chart2.set_y_axis({'name': 'Crecimiento (%)'})
            chart2.set_size({'width': 720, 'height': 400})
            sheet.insert_chart(row + len(crecimiento) + 2, 0, chart2)
    
    def _crear_hoja_datos_completos(self, df: pd.DataFrame):
        """
        Crea la hoja con todos los datos originales.
        """
        sheet = self.workbook.add_worksheet('Datos Completos')
        
        # Escribir encabezados
        for col_num, col_name in enumerate(df.columns):
            sheet.write(0, col_num, col_name, self.formatos['encabezado'])
        
        # Limpiar DataFrame de valores problemáticos
        df_clean = df.fillna('')
        df_clean = df_clean.replace([float('inf'), float('-inf')], 0)
        
        # Escribir datos
        for row_num, row_data in enumerate(df_clean.values, start=1):
            for col_num, cell_value in enumerate(row_data):
                # Convertir tipos especiales a string para evitar errores
                if pd.isna(cell_value) or cell_value == '':
                    sheet.write(row_num, col_num, '', self.formatos['normal'])
                elif isinstance(cell_value, (int, float)) and cell_value not in [float('inf'), float('-inf')]:
                    try:
                        sheet.write(row_num, col_num, cell_value, self.formatos['numero'])
                    except Exception:
                        sheet.write(row_num, col_num, str(cell_value), self.formatos['normal'])
                else:
                    sheet.write(row_num, col_num, str(cell_value), self.formatos['normal'])
        
        # Autoajustar columnas
        for col_num in range(len(df.columns)):
            sheet.set_column(col_num, col_num, 15)
    
    def _escribir_dataframe(self, sheet, df: pd.DataFrame, start_row: int, start_col: int):
        """
        Escribe un DataFrame en una hoja de Excel con formato.
        
        Args:
            sheet: Hoja de Excel donde escribir
            df: DataFrame a escribir
            start_row: Fila inicial
            start_col: Columna inicial
        """
        # Limpiar DataFrame de valores NaN e infinitos
        df_clean = df.fillna(0)  # Reemplazar NaN con 0
        
        # Reemplazar valores infinitos
        df_clean = df_clean.replace([float('inf'), float('-inf')], 0)
        
        # Escribir encabezados
        for col_num, col_name in enumerate(df_clean.columns):
            sheet.write(start_row, start_col + col_num, col_name, self.formatos['encabezado'])
        
        # Escribir datos
        for row_num, row_data in enumerate(df_clean.values, start=start_row + 1):
            for col_num, cell_value in enumerate(row_data):
                # Manejar valores especiales
                if pd.isna(cell_value) or cell_value in [float('inf'), float('-inf')]:
                    cell_value = 0
                    formato = self.formatos['normal']
                elif isinstance(cell_value, str):
                    formato = self.formatos['normal']
                elif 'Factura' in df_clean.columns[col_num] or 'Precio' in df_clean.columns[col_num]:
                    formato = self.formatos['moneda']
                elif '%' in df_clean.columns[col_num]:
                    formato = self.formatos['porcentaje']
                    if isinstance(cell_value, (int, float)) and not pd.isna(cell_value):
                        cell_value = cell_value / 100  # Convertir a decimal para formato porcentaje
                    else:
                        cell_value = 0
                elif isinstance(cell_value, float):
                    formato = self.formatos['numero']
                else:
                    formato = self.formatos['entero']
                
                try:
                    sheet.write(row_num, start_col + col_num, cell_value, formato)
                except Exception:
                    # En caso de error, escribir como texto
                    sheet.write(row_num, start_col + col_num, str(cell_value), self.formatos['normal'])

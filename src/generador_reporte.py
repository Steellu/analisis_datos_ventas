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
            }),
            'verde': self.workbook.add_format({
                'bg_color': '#C6EFCE',
                'font_color': '#006100',
                'border': 1,
                'bold': True
            }),
            'amarillo': self.workbook.add_format({
                'bg_color': '#FFEB9C',
                'font_color': '#9C6500',
                'border': 1,
                'bold': True
            }),
            'rojo': self.workbook.add_format({
                'bg_color': '#FFC7CE',
                'font_color': '#9C0006',
                'border': 1,
                'bold': True
            }),
            'texto_wrap': self.workbook.add_format({
                'border': 1,
                'text_wrap': True,
                'valign': 'vcenter'
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
                                frecuencia: Dict[str, Any],
                                # NUEVOS PARÁMETROS
                                pareto_peso: pd.DataFrame,
                                pareto_cantidad: pd.DataFrame,
                                pareto_facturacion: pd.DataFrame,
                                matriz_decision: pd.DataFrame,
                                segmentacion_bcg: pd.DataFrame):
        """
        Genera el reporte completo en Excel con todas las hojas y gráficos.
        """
        self.workbook = xlsxwriter.Workbook(self.ruta_salida, {'nan_inf_to_errors': True})
        self._crear_formatos()
      
        # Crear hojas (HOJAS DE PRIORIZACIÓN PRIMERO)
        # self._crear_hoja_matriz_decision(matriz_decision)
        # self._crear_hoja_segmentacion_bcg(segmentacion_bcg)
        # self._crear_hoja_pareto_peso(pareto_peso)
        # self._crear_hoja_pareto_cantidad(pareto_cantidad)
        # self._crear_hoja_pareto_facturacion_priorizacion(pareto_facturacion)
      
        # Hojas originales
        self._crear_hoja_resumen(resumen, frecuencia)
        self._crear_hoja_pareto(pareto)
        self._crear_hoja_productos(top_cantidad, top_facturacion, precio_kg)
        self._crear_hoja_categorias(categorias)
        self._crear_hoja_temporal(ventas_mes, crecimiento)
        self._crear_hoja_datos_completos(df_completo)
      
        # Crear hojas (HOJAS DE PRIORIZACIÓN PRIMERO)
        self._crear_hoja_matriz_decision(matriz_decision)
        self._crear_hoja_segmentacion_bcg(segmentacion_bcg)
        self._crear_hoja_pareto_peso(pareto_peso)
        self._crear_hoja_pareto_cantidad(pareto_cantidad)
        self._crear_hoja_pareto_facturacion_priorizacion(pareto_facturacion)
        self._crear_hoja_comparativa_peso_cantidad(df_completo)
      
        self.workbook.close()
        print(f"✅ Reporte generado exitosamente: {self.ruta_salida}")
  
    # ========== NUEVAS HOJAS DE PRIORIZACIÓN ==========
  
    def _crear_hoja_matriz_decision(self, matriz_decision: pd.DataFrame):
        """
        Crea la hoja de Matriz de Decisión (LA MÁS IMPORTANTE).
        """
        sheet = self.workbook.add_worksheet('1. Matriz de Decisión')
        sheet.set_column('A:B', 12)
        sheet.set_column('C:I', 14)
        sheet.set_column('J:J', 12)
        sheet.set_column('K:K', 50)
      
        row = 0
      
        # Título
        sheet.merge_range(row, 0, row, 10,
                         '🎯 MATRIZ DE DECISIÓN - ÍNDICE GLOBAL DE PRIORIZACIÓN',
                         self.formatos['titulo'])
        row += 1
        sheet.merge_range(row, 0, row, 10,
                         'Índice Global = 60% Eficiencia Fundición ($ por Kg) + 40% Eficiencia Mano de Obra (menos piezas)',
                         self.formatos['subtitulo'])
        row += 2
      
        # Explicación
        sheet.merge_range(row, 0, row, 10,
                         '💡 CÓMO USAR: Prioriza productos con Índice Global alto (verde). Estos maximizan $ por kg fundido y minimizan mano de obra.',
                         self.formatos['normal'])
        row += 2
      
        # Tabla
        self._escribir_dataframe_priorizacion(sheet, matriz_decision, row, 0)
      
        # Gráfico de dispersión: $ por Kg vs Cantidad
        if len(matriz_decision) > 0:
            chart = self.workbook.add_chart({'type': 'scatter'})
          
            # Obtener índices de columnas
            col_cantidad = 2  # Cantidad Total
            col_precio_kg = 7  # $ por Kg
          
            chart.add_series({
                'name': 'Productos',
                'categories': ['1. Matriz de Decisión', row + 1, col_cantidad, row + min(30, len(matriz_decision)), col_cantidad],
                'values': ['1. Matriz de Decisión', row + 1, col_precio_kg, row + min(30, len(matriz_decision)), col_precio_kg],
                'marker': {'type': 'circle', 'size': 8},
            })
          
            chart.set_title({'name': 'Eficiencia: $ por Kg vs Cantidad de Piezas'})
            chart.set_x_axis({'name': 'Cantidad de Piezas (menos es mejor)'})
            chart.set_y_axis({'name': '$ por Kg (más es mejor)'})
            chart.set_size({'width': 900, 'height': 500})
            chart.set_legend({'position': 'none'})
          
            sheet.insert_chart(row + min(30, len(matriz_decision)) + 2, 0, chart)
          
            # Gráfico de barras: Top 15 por Índice Global
            chart2 = self.workbook.add_chart({'type': 'bar'})
            chart2.add_series({
                'name': 'Índice Global',
                'categories': ['1. Matriz de Decisión', row + 1, 1, row + min(15, len(matriz_decision)), 1],
                'values': ['1. Matriz de Decisión', row + 1, 8, row + min(15, len(matriz_decision)), 8],
                'data_labels': {'value': True},
            })
          
            chart2.set_title({'name': 'Top 15 Productos por Índice Global'})
            chart2.set_x_axis({'name': 'Índice Global (0-100)'})
            chart2.set_y_axis({'name': 'Producto'})
            chart2.set_size({'width': 900, 'height': 500})
            chart2.set_legend({'position': 'none'})
          
            sheet.insert_chart(row + min(30, len(matriz_decision)) + 2, 6, chart2)
  
    def _crear_hoja_segmentacion_bcg(self, segmentacion_bcg: pd.DataFrame):
        """
        Crea la hoja de Segmentación BCG.
        """
        sheet = self.workbook.add_worksheet('2. Segmentación BCG')
        sheet.set_column('A:B', 12)
        sheet.set_column('C:F', 16)
        sheet.set_column('G:G', 20)
        sheet.set_column('H:H', 45)
      
        row = 0
      
        # Título
        sheet.merge_range(row, 0, row, 7,
                         '📈 SEGMENTACIÓN BCG - ESTRATEGIA POR CUADRANTE',
                         self.formatos['titulo'])
        row += 1
        sheet.merge_range(row, 0, row, 7,
                         '🐄 Vacas Lecheras = MÁXIMA PRIORIDAD | ⭐ Estrellas = Mantener | ⚡ Desafiantes = Revisar | 🐕 Perros = Descontinuar',
                         self.formatos['subtitulo'])
        row += 2
      
        # Explicación de cuadrantes
        sheet.merge_range(row, 0, row, 7,
                         '💡 INTERPRETACIÓN: Vacas Lecheras son ideales (bajo peso, alta facturación). Estrellas son buenos pero consumen más fundición.',
                         self.formatos['normal'])
        row += 2
      
        # Tabla
        self._escribir_dataframe(sheet, segmentacion_bcg, row, 0)
      
        # Contar productos por segmento
        if len(segmentacion_bcg) > 0:
            conteo = segmentacion_bcg['Segmento BCG'].value_counts()
          
            # Gráfico de torta: Distribución por segmento
            chart = self.workbook.add_chart({'type': 'pie'})
          
            # Crear datos para el gráfico
            row_temp = row + len(segmentacion_bcg) + 3
            sheet.write(row_temp, 0, 'Segmento', self.formatos['encabezado'])
            sheet.write(row_temp, 1, 'Cantidad', self.formatos['encabezado'])
          
            for idx, (segmento, cantidad) in enumerate(conteo.items(), start=1):
                sheet.write(row_temp + idx, 0, segmento, self.formatos['normal'])
                sheet.write(row_temp + idx, 1, cantidad, self.formatos['entero'])
          
            chart.add_series({
                'name': 'Distribución por Segmento BCG',
                'categories': ['2. Segmentación BCG', row_temp + 1, 0, row_temp + len(conteo), 0],
                'values': ['2. Segmentación BCG', row_temp + 1, 1, row_temp + len(conteo), 1],
                'data_labels': {'value': True, 'category': True},
                'points': [
                    {'fill': {'color': '#90EE90'}},  # Verde para Vacas
                    {'fill': {'color': '#FFD700'}},  # Amarillo para Estrellas
                    {'fill': {'color': '#FFA500'}},  # Naranja para Desafiantes
                    {'fill': {'color': '#FF6B6B'}},  # Rojo para Perros
                ],
            })
          
            chart.set_title({'name': 'Distribución de Productos por Segmento BCG'})
            chart.set_size({'width': 600, 'height': 450})
            chart.set_legend({'position': 'bottom'})
          
            sheet.insert_chart(row + len(segmentacion_bcg) + 2, 0, chart)
          
            # Gráfico de dispersión BCG: Peso vs Facturación
            chart2 = self.workbook.add_chart({'type': 'scatter'})
            chart2.add_series({
                'name': 'Productos',
                'categories': ['2. Segmentación BCG', row + 1, 2, row + len(segmentacion_bcg), 2],  # Peso
                'values': ['2. Segmentación BCG', row + 1, 3, row + len(segmentacion_bcg), 3],  # Facturación
                'marker': {'type': 'circle', 'size': 10},
            })
          
            chart2.set_title({'name': 'Matriz BCG: Peso Total vs Facturación'})
            chart2.set_x_axis({'name': 'Peso Total (kg)'})
            chart2.set_y_axis({'name': 'Facturación Total ($)'})
            chart2.set_size({'width': 900, 'height': 500})
            chart2.set_legend({'position': 'none'})
          
            sheet.insert_chart(row + len(segmentacion_bcg) + 2, 6, chart2)
  
    def _crear_hoja_pareto_peso(self, pareto_peso: pd.DataFrame):
        """
        Crea la hoja de Pareto por Peso (Capacidad de Fundición).
        """
        sheet = self.workbook.add_worksheet('3. Pareto por Peso')
        sheet.set_column('A:H', 16)
      
        row = 0
      
        # Título
        sheet.merge_range(row, 0, row, 7,
                         '⚙️ PARETO POR PESO - Capacidad de Fundición (kg)',
                         self.formatos['titulo'])
        row += 1
        sheet.merge_range(row, 0, row, 7,
                         'Identifica qué productos consumen el 80% de tu capacidad de fundición',
                         self.formatos['subtitulo'])
        row += 2
      
        # Resumen
        if len(pareto_peso) > 0:
            productos_80 = pareto_peso[pareto_peso['% Acumulado'] <= 80]
            sheet.merge_range(row, 0, row, 7,
                             f'📊 INSIGHT: {len(productos_80)} productos ({len(productos_80)/len(pareto_peso)*100:.1f}%) consumen el 80% del peso total',
                             self.formatos['normal'])
            row += 2
      
        # Tabla
        self._escribir_dataframe(sheet, pareto_peso.head(30), row, 0)
      
        # Gráfico de Pareto
        if len(pareto_peso) > 0:
            chart = self.workbook.add_chart({'type': 'column'})
            chart.add_series({
                'name': 'Peso Total (kg)',
                'categories': ['3. Pareto por Peso', row + 1, 1, row + min(20, len(pareto_peso)), 1],
                'values': ['3. Pareto por Peso', row + 1, 2, row + min(20, len(pareto_peso)), 2],
                'y2_axis': False,
            })
          
            # Línea de % acumulado
            chart.add_series({
                'name': '% Acumulado',
                'categories': ['3. Pareto por Peso', row + 1, 1, row + min(20, len(pareto_peso)), 1],
                'values': ['3. Pareto por Peso', row + 1, 6, row + min(20, len(pareto_peso)), 6],
                'y2_axis': True,
                'line': {'color': 'red', 'width': 2.5},
                'marker': {'type': 'circle', 'size': 5},
            })
          
            chart.set_title({'name': 'Diagrama de Pareto - Peso (Top 20 Productos)'})
            chart.set_x_axis({'name': 'Producto'})
            chart.set_y_axis({'name': 'Peso Total (kg)'})
            chart.set_y2_axis({'name': '% Acumulado', 'min': 0, 'max': 100})
            chart.set_size({'width': 1000, 'height': 500})
            chart.set_legend({'position': 'top'})
          
            sheet.insert_chart(row + min(30, len(pareto_peso)) + 2, 0, chart)
  
    def _crear_hoja_pareto_cantidad(self, pareto_cantidad: pd.DataFrame):
        """
        Crea la hoja de Pareto por Cantidad (Mano de Obra).
        """
        sheet = self.workbook.add_worksheet('4. Pareto por Cantidad')
        sheet.set_column('A:H', 16)
      
        row = 0
      
        # Título
        sheet.merge_range(row, 0, row, 7,
                         '👷 PARETO POR CANTIDAD - Mano de Obra (unidades)',
                         self.formatos['titulo'])
        row += 1
        sheet.merge_range(row, 0, row, 7,
                         'Identifica qué productos requieren el 80% de tu mano de obra (más piezas = más trabajo)',
                         self.formatos['subtitulo'])
        row += 2
      
        # Resumen
        if len(pareto_cantidad) > 0:
            productos_80 = pareto_cantidad[pareto_cantidad['% Acumulado'] <= 80]
            sheet.merge_range(row, 0, row, 7,
                             f'📊 INSIGHT: {len(productos_80)} productos ({len(productos_80)/len(pareto_cantidad)*100:.1f}%) representan el 80% de las piezas producidas',
                             self.formatos['normal'])
            row += 2
      
        # Tabla
        self._escribir_dataframe(sheet, pareto_cantidad.head(30), row, 0)
      
        # Gráfico de Pareto
        if len(pareto_cantidad) > 0:
            chart = self.workbook.add_chart({'type': 'column'})
            chart.add_series({
                'name': 'Cantidad Total',
                'categories': ['4. Pareto por Cantidad', row + 1, 1, row + min(20, len(pareto_cantidad)), 1],
                'values': ['4. Pareto por Cantidad', row + 1, 2, row + min(20, len(pareto_cantidad)), 2],
                'y2_axis': False,
            })
          
            # Línea de % acumulado
            chart.add_series({
                'name': '% Acumulado',
                'categories': ['4. Pareto por Cantidad', row + 1, 1, row + min(20, len(pareto_cantidad)), 1],
                'values': ['4. Pareto por Cantidad', row + 1, 6, row + min(20, len(pareto_cantidad)), 6],
                'y2_axis': True,
                'line': {'color': 'red', 'width': 2.5},
                'marker': {'type': 'circle', 'size': 5},
            })
          
            chart.set_title({'name': 'Diagrama de Pareto - Cantidad (Top 20 Productos)'})
            chart.set_x_axis({'name': 'Producto'})
            chart.set_y_axis({'name': 'Cantidad Total'})
            chart.set_y2_axis({'name': '% Acumulado', 'min': 0, 'max': 100})
            chart.set_size({'width': 1000, 'height': 500})
            chart.set_legend({'position': 'top'})
          
            sheet.insert_chart(row + min(30, len(pareto_cantidad)) + 2, 0, chart)
  
    def _crear_hoja_pareto_facturacion_priorizacion(self, pareto_facturacion: pd.DataFrame):
        """
        Crea la hoja de Pareto por Facturación (Ingresos).
        """
        sheet = self.workbook.add_worksheet('5. Pareto por Facturación')
        sheet.set_column('A:H', 16)
      
        row = 0
      
        # Título
        sheet.merge_range(row, 0, row, 7,
                         '💰 PARETO POR FACTURACIÓN - Ingresos ($)',
                         self.formatos['titulo'])
        row += 1
        sheet.merge_range(row, 0, row, 7,
                         'Identifica qué productos generan el 80% de tus ingresos',
                         self.formatos['subtitulo'])
        row += 2
      
        # Resumen
        if len(pareto_facturacion) > 0:
            productos_80 = pareto_facturacion[pareto_facturacion['% Acumulado'] <= 80]
            sheet.merge_range(row, 0, row, 7,
                             f'📊 INSIGHT: {len(productos_80)} productos ({len(productos_80)/len(pareto_facturacion)*100:.1f}%) generan el 80% de la facturación',
                             self.formatos['normal'])
            row += 2
      
        # Tabla
        self._escribir_dataframe(sheet, pareto_facturacion.head(30), row, 0)
      
        # Gráfico de Pareto
        if len(pareto_facturacion) > 0:
            chart = self.workbook.add_chart({'type': 'column'})
            chart.add_series({
                'name': 'Facturación Total ($)',
                'categories': ['5. Pareto por Facturación', row + 1, 1, row + min(20, len(pareto_facturacion)), 1],
                'values': ['5. Pareto por Facturación', row + 1, 2, row + min(20, len(pareto_facturacion)), 2],
                'y2_axis': False,
            })
          
            # Línea de % acumulado
            chart.add_series({
                'name': '% Acumulado',
                'categories': ['5. Pareto por Facturación', row + 1, 1, row + min(20, len(pareto_facturacion)), 1],
                'values': ['5. Pareto por Facturación', row + 1, 6, row + min(20, len(pareto_facturacion)), 6],
                'y2_axis': True,
                'line': {'color': 'red', 'width': 2.5},
                'marker': {'type': 'circle', 'size': 5},
            })
          
            chart.set_title({'name': 'Diagrama de Pareto - Facturación (Top 20 Productos)'})
            chart.set_x_axis({'name': 'Producto'})
            chart.set_y_axis({'name': 'Facturación Total ($)'})
            chart.set_y2_axis({'name': '% Acumulado', 'min': 0, 'max': 100})
            chart.set_size({'width': 1000, 'height': 500})
            chart.set_legend({'position': 'top'})
          
            sheet.insert_chart(row + min(30, len(pareto_facturacion)) + 2, 0, chart)
  
    # ========== MÉTODOS AUXILIARES ==========
  
    def _escribir_dataframe_priorizacion(self, sheet, df: pd.DataFrame, start_row: int, start_col: int):
        """
        Escribe el DataFrame de Matriz de Decisión con colores según prioridad.
        """
        df_clean = df.fillna(0)
        df_clean = df_clean.replace([float('inf'), float('-inf')], 0)
      
        # Escribir encabezados
        for col_num, col_name in enumerate(df_clean.columns):
            sheet.write(start_row, start_col + col_num, col_name, self.formatos['encabezado'])
      
        # Escribir datos con formato condicional
        for row_num, row_data in enumerate(df_clean.values, start=start_row + 1):
            prioridad = row_data[9]  # Columna "Prioridad"
          
            for col_num, cell_value in enumerate(row_data):
                if pd.isna(cell_value) or cell_value in [float('inf'), float('-inf')]:
                    cell_value = 0
              
                # Determinar formato
                if col_num == 9:  # Columna Prioridad
                    if prioridad == 'Alta':
                        formato = self.formatos['verde']
                    elif prioridad == 'Media':
                        formato = self.formatos['amarillo']
                    else:
                        formato = self.formatos['rojo']
                elif col_num == 10:  # Columna Recomendación
                    formato = self.formatos['texto_wrap']
                elif isinstance(cell_value, str):
                    formato = self.formatos['normal']
                elif 'Facturación' in df_clean.columns[col_num] or '$' in df_clean.columns[col_num]:
                    formato = self.formatos['moneda']
                elif 'Índice' in df_clean.columns[col_num]:
                    formato = self.formatos['numero']
                elif isinstance(cell_value, float):
                    formato = self.formatos['numero']
                else:
                    formato = self.formatos['entero']
              
                try:
                    sheet.write(row_num, start_col + col_num, cell_value, formato)
                except Exception:
                    sheet.write(row_num, start_col + col_num, str(cell_value), self.formatos['normal'])
  
    # ========== HOJAS ORIGINALES (SIN CAMBIOS) ==========
  
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
        sheet.write(row, 0, 'PRODUCTOS VITALES (80% de facturación):', self.formatos['encabezado'])
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
        sheet.write(row, 0, 'PRODUCTOS TRIVIALES (20% de facturación):', self.formatos['encabezado'])
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
      
        df_pareto = pareto['dataframe'].head(50)
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
        """
        df_clean = df.fillna(0)
        df_clean = df_clean.replace([float('inf'), float('-inf')], 0)
      
        # Escribir encabezados
        for col_num, col_name in enumerate(df_clean.columns):
            sheet.write(start_row, start_col + col_num, col_name, self.formatos['encabezado'])
      
        # Escribir datos
        for row_num, row_data in enumerate(df_clean.values, start=start_row + 1):
            for col_num, cell_value in enumerate(row_data):
                if pd.isna(cell_value) or cell_value in [float('inf'), float('-inf')]:
                    cell_value = 0
              
                formato = self.formatos['normal']
                if isinstance(cell_value, str):
                    formato = self.formatos['normal']
                elif 'Factura' in df_clean.columns[col_num] or '$' in df_clean.columns[col_num] or 'Precio' in df_clean.columns[col_num]:
                    formato = self.formatos['moneda']
                elif '%' in df_clean.columns[col_num]:
                    formato = self.formatos['porcentaje']
                    if isinstance(cell_value, (int, float)) and not pd.isna(cell_value):
                        cell_value = cell_value / 100
                    else:
                        cell_value = 0
                elif isinstance(cell_value, float):
                    formato = self.formatos['numero']
                else:
                    formato = self.formatos['entero']
              
                try:
                    sheet.write(row_num, start_col + col_num, cell_value, formato)
                except Exception:
                    sheet.write(row_num, start_col + col_num, str(cell_value), self.formatos['normal'])
  
    def _crear_hoja_comparativa_peso_cantidad(self, df_completo: pd.DataFrame):
    # """
    # Crea una hoja comparativa de Peso vs Cantidad por producto.
    # """
        sheet = self.workbook.add_worksheet('6. Peso vs Cantidad')
        sheet.set_column('A:E', 18)

        row = 0
        sheet.merge_range(row, 0, row, 4,
                        '⚖️ COMPARATIVA: PESO vs CANTIDAD por Producto',
                        self.formatos['titulo'])
        row += 1
        sheet.merge_range(row, 0, row, 4,
                        'Identifica productos que pesan mucho pero se piden poco (o viceversa)',
                        self.formatos['subtitulo'])
        row += 2

        # ✅ Usar la columna MONTO_FACTURADO ya calculada
        comparativa = df_completo.groupby(['CODIGO', 'NOMBRE']).agg({
            'CANT': 'sum',
            'PESO TOTAL': 'sum',
            'MONTO_FACTURADO': 'sum'
        }).reset_index()

        comparativa.columns = ['Código', 'Nombre', 'Cantidad Total', 'Peso Total (kg)', 'Facturación Total ($)']
        comparativa = comparativa.sort_values('Peso Total (kg)', ascending=False).head(20)

        # Insight
        if len(comparativa) > 0:
            producto_mas_pesado = comparativa.iloc[0]
            producto_mas_cantidad = comparativa.sort_values('Cantidad Total', ascending=False).iloc[0]
            sheet.merge_range(row, 0, row, 4,
                            f'📊 INSIGHT: "{producto_mas_pesado["Nombre"]}" es el más pesado '
                            f'({producto_mas_pesado["Peso Total (kg)"]:.0f} kg). '
                            f'"{producto_mas_cantidad["Nombre"]}" tiene más piezas '
                            f'({producto_mas_cantidad["Cantidad Total"]:.0f} unidades).',
                            self.formatos['normal'])
            row += 2

        # Tabla
        self._escribir_dataframe(sheet, comparativa, row, 0)

        # (mantén los gráficos igual que estaban)

        
            # Tabla
        self._escribir_dataframe(sheet, comparativa, row, 0)
        
        # GRÁFICO COMBINADO: Barras para Cantidad + Línea para Peso
        if len(comparativa) > 0:
            chart = self.workbook.add_chart({'type': 'column'})
        
            # Serie 1: Cantidad (barras azules)
            chart.add_series({
                'name': 'Cantidad Total',
                'categories': ['6. Peso vs Cantidad', row + 1, 1, row + len(comparativa), 1],  # Nombres
                'values': ['6. Peso vs Cantidad', row + 1, 2, row + len(comparativa), 2],  # Cantidad
                'fill': {'color': '#3498DB'},
                'y2_axis': False,
            })
        
            # Serie 2: Peso (línea roja en eje secundario)
            chart.add_series({
                'name': 'Peso Total (kg)',
                'categories': ['6. Peso vs Cantidad', row + 1, 1, row + len(comparativa), 1],  # Nombres
                'values': ['6. Peso vs Cantidad', row + 1, 3, row + len(comparativa), 3],  # Peso
                'line': {'color': '#E74C3C', 'width': 3},
                'marker': {'type': 'circle', 'size': 7, 'fill': {'color': '#E74C3C'}},
                'y2_axis': True,
            })
        
            chart.set_title({'name': 'Comparativa: Cantidad vs Peso (Top 20 Productos)'})
            chart.set_x_axis({'name': 'Producto', 'label_position': 'low'})
            chart.set_y_axis({'name': 'Cantidad Total (unidades)', 'major_gridlines': {'visible': True}})
            chart.set_y2_axis({'name': 'Peso Total (kg)', 'major_gridlines': {'visible': False}})
            chart.set_size({'width': 1100, 'height': 550})
            chart.set_legend({'position': 'top'})
        
            sheet.insert_chart(row + len(comparativa) + 2, 0, chart)
        
            # GRÁFICO DE DISPERSIÓN: Peso vs Cantidad
            chart2 = self.workbook.add_chart({'type': 'scatter'})
            chart2.add_series({
                'name': 'Productos',
                'categories': ['6. Peso vs Cantidad', row + 1, 2, row + len(comparativa), 2],  # Cantidad
                'values': ['6. Peso vs Cantidad', row + 1, 3, row + len(comparativa), 3],  # Peso
                'marker': {'type': 'circle', 'size': 10, 'fill': {'color': '#9B59B6'}},
            })
        
            chart2.set_title({'name': 'Relación: Cantidad vs Peso'})
            chart2.set_x_axis({'name': 'Cantidad Total (unidades)'})
            chart2.set_y_axis({'name': 'Peso Total (kg)'})
            chart2.set_size({'width': 900, 'height': 500})
            chart2.set_legend({'position': 'none'})
        
            sheet.insert_chart(row + len(comparativa) + 32, 0, chart2)
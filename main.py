"""
Script principal para generar reportes de análisis de ventas por cliente.
Ejecuta este archivo para procesar un Excel y generar el reporte completo.
"""

import os
from src.analisis import AnalizadorVentas
from src.generador_reporte import GeneradorReporte


def main():
    """
    Función principal que ejecuta el análisis y genera el reporte.
    """
    print("=" * 60)
    print("🔧 SISTEMA DE ANÁLISIS DE VENTAS - EMPRESA METALÚRGICA")
    print("=" * 60)
    print()
    
    # Configuración de rutas
    carpeta_input = "data/input"
    carpeta_output = "data/output"
    
    # Crear carpetas si no existen
    os.makedirs(carpeta_input, exist_ok=True)
    os.makedirs(carpeta_output, exist_ok=True)
    
    # Listar archivos Excel disponibles en la carpeta input
    archivos_excel = [f for f in os.listdir(carpeta_input) if f.endswith(('.xlsx', '.xls'))]
    
    if not archivos_excel:
        print("❌ No se encontraron archivos Excel en la carpeta 'data/input'")
        print(f"📁 Por favor, coloca tu archivo Excel en: {os.path.abspath(carpeta_input)}")
        return
    
    # Mostrar archivos disponibles
    print("📂 Archivos Excel encontrados:")
    for i, archivo in enumerate(archivos_excel, 1):
        print(f"   {i}. {archivo}")
    print()
    
    # Seleccionar archivo
    if len(archivos_excel) == 1:
        archivo_seleccionado = archivos_excel[0]
        print(f"✅ Procesando único archivo: {archivo_seleccionado}")
    else:
        try:
            seleccion = int(input("Selecciona el número del archivo a procesar: "))
            if 1 <= seleccion <= len(archivos_excel):
                archivo_seleccionado = archivos_excel[seleccion - 1]
            else:
                print("❌ Selección inválida")
                return
        except ValueError:
            print("❌ Entrada inválida")
            return
    
    # Rutas completas
    ruta_input = os.path.join(carpeta_input, archivo_seleccionado)
    nombre_sin_extension = os.path.splitext(archivo_seleccionado)[0]
    ruta_output = os.path.join(carpeta_output, f"reporte_{nombre_sin_extension}.xlsx")
    
    print()
    print("=" * 60)
    print("📊 INICIANDO ANÁLISIS...")
    print("=" * 60)
    
    try:
        # 1. Cargar y analizar datos
        print("\n🔍 Paso 1: Cargando datos del Excel...")
        analizador = AnalizadorVentas(ruta_input)
        print(f"   ✓ Cliente: {analizador.cliente}")
        print(f"   ✓ Registros cargados: {len(analizador.df)}")
        
        # 2. Generar análisis
        print("\n📈 Paso 2: Generando análisis...")
        resumen = analizador.resumen_general()
        print(f"   ✓ Total facturado: $ {resumen['total_facturado']:,.2f}")
        print(f"   ✓ Total órdenes: {resumen['total_ordenes']}")
        print(f"   ✓ Productos únicos: {resumen['productos_unicos']}")
        
        top_cantidad = analizador.top_productos_cantidad(10)
        print(f"   ✓ Top productos por cantidad: {len(top_cantidad)} productos")
        
        top_facturacion = analizador.top_productos_facturacion(10)
        print(f"   ✓ Top productos por facturación: {len(top_facturacion)} productos")
        
        categorias = analizador.analisis_categorias()
        print(f"   ✓ Análisis de categorías: {len(categorias)} categorías")
        
        ventas_mes = analizador.ventas_por_mes()
        print(f"   ✓ Análisis temporal: {len(ventas_mes)} meses")
        
        precio_kg = analizador.productos_precio_alto_kg(10)
        print(f"   ✓ Productos precio/kg: {len(precio_kg)} productos")
        
        # NUEVOS ANÁLISIS
        pareto = analizador.analisis_pareto()
        print(f"   ✓ Análisis de Pareto: {pareto['porcentaje_productos_80']}% de productos generan 80% de ventas")
        
        crecimiento = analizador.crecimiento_mensual()
        print(f"   ✓ Crecimiento mensual: {len(crecimiento)} meses analizados")
        
        frecuencia = analizador.frecuencia_compra()
        print(f"   ✓ Frecuencia de compra: cada {frecuencia['dias_entre_compras']} días")
        
        df_completo = analizador.obtener_dataframe_completo()
        
        # 3. Generar reporte Excel
        print("\n📝 Paso 3: Generando reporte Excel...")
        generador = GeneradorReporte(ruta_output)
        generador.generar_reporte_completo(
            resumen=resumen,
            top_cantidad=top_cantidad,
            top_facturacion=top_facturacion,
            categorias=categorias,
            ventas_mes=ventas_mes,
            precio_kg=precio_kg,
            df_completo=df_completo,
            pareto=pareto,
            crecimiento=crecimiento,
            frecuencia=frecuencia
        )
        
        print()
        print("=" * 60)
        print("✅ PROCESO COMPLETADO EXITOSAMENTE")
        print("=" * 60)
        print(f"\n📄 Reporte generado en:")
        print(f"   {os.path.abspath(ruta_output)}")
        print()
        
    except Exception as e:
        print()
        print("=" * 60)
        print("❌ ERROR AL PROCESAR EL ARCHIVO")
        print("=" * 60)
        print(f"\n{type(e).__name__}: {str(e)}")
        print()
        print("💡 Verifica que:")
        print("   - El archivo Excel tenga el formato correcto")
        print("   - Las columnas necesarias estén presentes")
        print("   - El archivo no esté abierto en otro programa")
        print()


if __name__ == "__main__":
    main()
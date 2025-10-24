# """
# Paquete de análisis de datos para empresa metalúrgica.
# Contiene módulos para analizar ventas por cliente y generar reportes en Excel.
# """

__version__ = "1.0.0"
__author__ = "Tu Nombre"

# Importaciones para facilitar el uso del paquete
from .analisis import AnalizadorVentas
from .generador_reporte import GeneradorReporte

__all__ = ['AnalizadorVentas', 'GeneradorReporte']
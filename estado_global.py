# estado_global.py
import pandas as pd

CARPETA_ARCHIVOS = "archivos_subidos"  # carpeta donde guardas archivos

# Variables globales para compartir estado y datos
usuarios_df = pd.DataFrame()
estado_excel = "📊 Archivo Excel: No cargado ❌"
ultima_ruta_archivo = ""

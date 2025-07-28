import pandas as pd
import os
import base64
import pdfkit
from jinja2 import Environment, FileSystemLoader
import re

# Configuración de rutas
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BANNER_PATH = os.path.join(BASE_DIR, "img", "banner.jpg")
FOOTER_PATH = os.path.join(BASE_DIR, "img", "footer.jpg")

# Función para convertir imágenes a formato Data URI
def imagen_a_data_uri(ruta_archivo):
    """Convierte una imagen a Data URI para incrustarla en el HTML"""
    if not os.path.exists(ruta_archivo):
        print(f"⚠️ Advertencia: No se encontró {ruta_archivo}")
        return ""
    
    try:
        # Determinar tipo MIME
        ext = os.path.splitext(ruta_archivo)[1].lower().replace(".", "")
        mime = f"image/{'jpeg' if ext == 'jpg' else ext}"
        
        # Leer y codificar la imagen
        with open(ruta_archivo, "rb") as img_file:
            encoded = base64.b64encode(img_file.read()).decode('utf-8')
            
        return f"data:{mime};base64,{encoded}"
    except Exception as e:
        print(f"❌ Error procesando imagen {ruta_archivo}: {e}")
        return ""

# Convertir las imágenes
BANNER_URI = imagen_a_data_uri(BANNER_PATH)
FOOTER_URI = imagen_a_data_uri(FOOTER_PATH)

# Funciones de formateo de datos
def formato_moneda(valor):
    """Formatea números a formato moneda: $ 1.234.567,89"""
    if valor in ["--", "", None] or pd.isna(valor):
        return "--"
    try:
        # Manejar strings con separadores
        if isinstance(valor, str):
            valor = valor.replace(".", "").replace(",", ".")
        valor = float(valor)
        return f"${valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception as e:
        print(f"⚠️ Error formateando moneda: {e} | Valor: {valor}")
        return str(valor)
    
def formato_porcentaje(valor):
    """Formatea porcentajes a xx,xx% siempre con dos decimales"""
    if valor in ["--", "", None] or pd.isna(valor):
        return "--"
    
    try:
        # Convertir a string y limpiar
        valor_str = str(valor).strip().replace("%", "").replace(",", ".")
        valor_float = float(valor_str)
        
        # Convertir decimales a porcentaje (0.5 -> 50%)
        if 0 <= valor_float <= 1:
            valor_float *= 100
        
        # Formatear a dos decimales con coma
        return f"{valor_float:.2f}".replace(".", ",") + "%"
    except Exception as e:
        print(f"⚠️ Error formateando porcentaje: {e} | Valor: {valor}")
        return str(valor)

def formato_numero(valor):
    """Formatea números con separadores de miles: 1.234.567,89"""
    if valor in ["--", "", None] or pd.isna(valor):
        return "--"
    try:
        # Manejar strings con separadores
        if isinstance(valor, str):
            valor = valor.replace(".", "").replace(",", ".")
        valor = float(valor)
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception as e:
        print(f"⚠️ Error formateando número: {e} | Valor: {valor}")
        return str(valor)
    
# Configurar wkhtmltopdf
config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")

# Cargar Excel
df = pd.read_excel("Reporte de obras abreviado.xlsx", engine='openpyxl')

# Configurar Jinja2
env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("plantilla.html")

# Crear carpeta de salida
os.makedirs("informes", exist_ok=True)

# Generar PDFs
for idx, fila in df.iterrows():
    try:
        # Buscar imagen de obra como JPG o PNG y convertir a base64
        ruta_imagen_obra = os.path.join(BASE_DIR, "imagenes_obras", f"{fila['ID obra']}.jpg")
        if not os.path.exists(ruta_imagen_obra):
            ruta_imagen_obra = os.path.join(BASE_DIR, "imagenes_obras", f"{fila['ID obra']}.png")

        imagen_obra_uri = imagen_a_data_uri(ruta_imagen_obra)

        # Preparar datos para la plantilla
        datos = {
            "banner_path": BANNER_URI,
            "footer_path": FOOTER_URI,
            "Memoria_Descriptiva": fila.get("Descripción", "--"),
            "Imagen_Obra": imagen_obra_uri,
            "ID_obra": fila.get("ID obra", "--"),
            "Estado": fila.get("Estado", "--"),
            "Solicitante_Financiamiento": fila.get("Solicitante financiamiento", "--"),
            "Solicitante_Presupuestario": fila.get("Solicitante presupuestario", "--"),
            "Municipio": fila.get("Municipio/s", "--"),
            "Localidad": "--",  # Campo no disponible
            "Modalidad": fila.get("Modalidad", "--"),
            "Programa": "Programa COMPLETAR",  # Valor fijo
            "Cod_emprendimiento": fila.get("Código emprendimiento", "--"),
            "Cod_obra": fila.get("Código de obra", "--"),
            "Monto_Convenio": formato_moneda(fila.get("Monto actualizado (ARS)", "--")),
            "Fecha_UVI": "--",  # Campo no disponible
            "Total_UVI": formato_numero(fila.get("Total UVI", "--")),
            "Exp_GDEBA": fila.get("Expediente GDEBA", "--"),
            "Avance_físico": formato_porcentaje(fila.get("% Av. físico", "--")),
            "Avance_financiero": formato_porcentaje(fila.get("% Av. financiero", "--")),
        }
        
        # Generar HTML
        html = template.render(**datos)
        
        # Crear nombre de archivo seguro
        nombre_base = f"informe_{fila['ID obra']}"
        nombre_base = re.sub(r'[\\/*?:"<>|]', "", nombre_base)
        nombre_archivo = os.path.join("informes", f"{nombre_base}.pdf")
        
        # Generar PDF
        options = {
            'enable-local-file-access': None  # necesario para cargar fuentes locales y archivos locales
        }
        pdfkit.from_string(html, nombre_archivo, configuration=config, options=options)
        print(f"✅ Generado: {nombre_archivo}")
        
    except Exception as e:
        print(f"❌ Error en la fila {idx}: {str(e)}")
        # Guardar HTML para depuración
        with open(f"error_{idx}.html", "w", encoding="utf-8") as f:
            f.write(html)

print("🎉 Todos los informes fueron generados.")
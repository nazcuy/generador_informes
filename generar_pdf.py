import pandas as pd
import os
import base64
import pdfkit
from jinja2 import Environment, FileSystemLoader
import re

# Configuración de rutas
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BANNER_PATH = os.path.join(BASE_DIR, "img", "banner.jpg")
HEADER_HTML_PATH = os.path.join(BASE_DIR, "header.html")
HEADER_RENDERED_HTML = os.path.join(BASE_DIR, "header_rendered.html")
FOOTER_PATH = os.path.abspath(os.path.join(BASE_DIR, "img", "footer.jpg"))
FOOTER_HTML_PATH = os.path.abspath(os.path.join(BASE_DIR, "footer.html"))
FOOTER_RENDERED_HTML = os.path.abspath(os.path.join(BASE_DIR, "footer_rendered.html"))
DOBLE_FLECHA_PATH = os.path.join(BASE_DIR, "img", "doble_flecha.jpg")

with open(BANNER_PATH, "rb") as image_file:
    banner_base64 = base64.b64encode(image_file.read()).decode('utf-8')
with open(HEADER_HTML_PATH, "r", encoding="utf-8") as f:
    header_html_content = f.read()
header_html_rendered = header_html_content.replace(
    "{{ banner_base64 }}", f"data:image/jpeg;base64,{banner_base64}"  # O usa BANNER_URI si lo tenés sin encabezado MIME
)
with open(HEADER_RENDERED_HTML, "w", encoding="utf-8") as f:
    f.write(header_html_rendered)


with open(FOOTER_PATH, "rb") as image_file:
    footer_base64 = base64.b64encode(image_file.read()).decode('utf-8')
# Leer footer.html base y reemplazar {{ footer_base64 }}
with open(FOOTER_HTML_PATH, "r", encoding="utf-8") as f:
    footer_html_content = f.read()

footer_html_rendered = footer_html_content.replace(
    "{{ footer_base64 }}", f"data:image/jpeg;base64,{footer_base64}"
)
with open(FOOTER_RENDERED_HTML, "w", encoding="utf-8") as f:
    f.write(footer_html_rendered)
   
def fuente_a_base64(ruta_fuente):
    #Convierte una fuente a base64 para incrustarla en el CSS
    if not os.path.exists(ruta_fuente):
        print(f"⚠️ Advertencia: No se encontró {ruta_fuente}")
        return ""
    
    try:
        # Leer y codificar la fuente
        with open(ruta_fuente, "rb") as font_file:
            encoded = base64.b64encode(font_file.read()).decode('utf-8')
        return encoded
    except Exception as e:
        print(f"❌ Error procesando fuente {ruta_fuente}: {e}")
        return ""
    
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
DOBLE_FLECHA_URI = imagen_a_data_uri(DOBLE_FLECHA_PATH)

# Convertir las fuentes a Base64
FUENTE_REGULAR_PATH = os.path.join(BASE_DIR, "fonts", "EncodeSans-Regular.ttf")
FUENTE_BOLD_PATH = os.path.join(BASE_DIR, "fonts", "EncodeSans-Bold.ttf")

fuente_regular_base64 = fuente_a_base64(FUENTE_REGULAR_PATH)
fuente_bold_base64 = fuente_a_base64(FUENTE_BOLD_PATH)

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
        return f"$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception as e:
        print(f"⚠️ Error formateando moneda: {e} | Valor: {valor}")
        return str(valor)
    
def formato_moneda_sin_decimales(valor):
    """Formatea números a formato moneda: $ 1.234.567 (sin decimales)"""
    if valor in ["--", "", None] or pd.isna(valor):
        return "--"
    try:
        if isinstance(valor, str):
            valor = valor.replace(".", "").replace(",", ".")
        valor = float(valor)
        return f"$ {valor:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
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
        return f"{valor:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception as e:
        print(f"⚠️ Error formateando número: {e} | Valor: {valor}")
        return str(valor)
    
def formato_numero_sin_decimales(valor):
    """Devuelve el número sin decimales y sin separadores de miles."""
    if valor in ["--", "", None] or pd.isna(valor):
        return "--"
    try:
        # Manejar strings con separadores
        if isinstance(valor, str):
            valor = valor.replace(".", "").replace(",", ".")
        valor = float(valor)
        valor = int(valor)
        return str(valor)
    except Exception as e:
        print(f"⚠️ Error formateando número: {e} | Valor: {valor}")
        return str(valor)
    
def formato_fecha(fecha):
    """Formatea fechas a DD/MM/YYYY. Si no hay fecha, devuelve '--'."""
    if fecha in ["--", "", None] or pd.isna(fecha):
        return "--"
    try:
        return fecha.strftime("%d/%m/%Y")
    except Exception as e:
        print(f"⚠️ Error formateando fecha: {e} | Valor: {fecha}")
        return str(fecha)

def chunk_text(text, size=20):
    if not text:
        return ""
    return "<br>".join(text[i:i+size] for i in range(0, len(text), size))

def dividir_en_grupos(lista, tamaño=4):
    """Divide una lista en sublistas de largo `tamaño`"""
    return [lista[i:i+tamaño] for i in range(0, len(lista), tamaño)]

# Configurar wkhtmltopdf
config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")

# Cargar Excel
df = pd.read_excel("pdf_generator_3000.xlsx", engine='openpyxl')

# Configurar Jinja2
env = Environment(loader=FileSystemLoader('.'))
env.filters['chunk'] = chunk_text
env.filters['dividir'] = dividir_en_grupos
template = env.get_template("plantilla.html")

# Crear carpeta de salida
os.makedirs("informes", exist_ok=True)

# Generar PDFs
for idx, fila in df.iterrows():
    try:
        # Buscar imagen de obra como JPG o PNG y convertir a base64
        ruta_imagen_obra = os.path.join(BASE_DIR, "imagenes_obras", f"{fila['id_obra']}.jpg")
        if not os.path.exists(ruta_imagen_obra):
            ruta_imagen_obra = os.path.join(BASE_DIR, "imagenes_obras", f"{fila['id_obra']}.png")

        imagen_obra_uri = imagen_a_data_uri(ruta_imagen_obra)
        
        # 1. Buscar todos los archivos que empiecen con ID_obra + "_"
        carpeta_img = os.path.join(BASE_DIR, "imagenes_obras")
        imagenes_extra = []
        for nombre in sorted(os.listdir(carpeta_img)):
            if nombre.startswith(f"{fila['id_obra']}_"):
                ruta = os.path.join(carpeta_img, nombre)
                uri = imagen_a_data_uri(ruta)
                if uri:
                    imagenes_extra.append(uri)
        
        # Preparar datos para la plantilla
        datos = {
            "banner_path": BANNER_URI,
            "footer_path": FOOTER_URI,
            "doble_flecha": DOBLE_FLECHA_URI,
            "fuente_regular": fuente_regular_base64,
            "fuente_bold": fuente_bold_base64,
            "Memoria_Descriptiva": fila.get("descripcion", "--"),
            "Imagen_Obra": imagen_obra_uri,
            "Imagenes_Extra": imagenes_extra,
            "ID_obra": fila.get("id_obra", "--"),
            "ID_historico": fila.get("id_historico", "--"),
            "Viviendas": formato_numero(fila.get("viv_totales", "--")),
            "Estado": fila.get("estado", "--"),
            "Solicitante_Financiamiento": fila.get("solicitante_financiero", "--"),
            "Solicitante_Presupuestario": fila.get("solicitante_presupuestario", "--"),
            "Municipio": fila.get("municipio", "--"),
            "Localidad":  fila.get("localidad", "--"),
            "Modalidad": fila.get("modalidad", "--"),
            "Programa": "Programa COMPLETAR",  # Valor fijo
            "Cod_emprendimiento": formato_numero_sin_decimales(fila.get("emprendimiento_incluidos", "--")),
            "Cod_obra": formato_numero_sin_decimales(fila.get("codigos_incluidos", "--")),
            "Monto_Convenio": formato_moneda(fila.get("monto_convenio", "--")),
            "Fecha_UVI": formato_fecha(fila.get("fecha_cotizacion_uvi_convenio")),
            "Total_UVI": formato_numero(fila.get("cantidad_uvis", "--")),
            "Exp_GDEBA": "" if pd.isna(fila.get("expediente_gdeba")) else str(fila.get("expediente_gdeba")),
            "Avance_físico": formato_porcentaje(fila.get("porcentaje_avance_fisico_anterior", "--")),
            "Avance_financiero": formato_porcentaje(fila.get("avance_financiero", "--")),
            "Monto_actualizado": formato_moneda_sin_decimales(fila.get("monto_actualizado", "--")),
            "Monto_Devengado": formato_moneda(fila.get("monto_devengado", "--")),
            "Monto_Pagado": formato_moneda(fila.get("monto_pagado", "--")),
            "Fecha_ultimo_pago": formato_fecha(fila.get("fecha_ultimo_pago")),
        }
        
        # Generar HTML
        html = template.render(**datos)
        
        # Crear nombre de archivo seguro
        nombre_base = f"informe_{fila['id_obra']}"
        nombre_base = re.sub(r'[\\/*?:"<>|]', "", nombre_base)
        nombre_archivo = os.path.join("informes", f"{nombre_base}.pdf")
        
        # Generar PDF
        options = {
            'enable-local-file-access': None,
            'margin-top': '30mm',
            'margin-bottom': '20mm',
            'margin-left': '4mm',
            'margin-right': '4mm',
            'footer-html': FOOTER_RENDERED_HTML,
            'header-html': HEADER_RENDERED_HTML
        }

        pdfkit.from_string(html, nombre_archivo, configuration=config, options=options)
        print(f"✅ Generado: {nombre_archivo}")
                
    except Exception as e:
        print(f"❌ Error en la fila {idx}: {str(e)}")
        # Guardar HTML para depuración
        with open(f"error_{idx}.html", "w", encoding="utf-8") as f:
            f.write(html)

print("🎉 Todo el proceso ha sido finalizado.")
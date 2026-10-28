import pdfplumber
import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
import datetime
import os
import io
import shutil
import win32com.client
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from pypdf import PdfReader, PdfWriter

PDF_INPUT = "Lote_Actual.pdf"
EXCEL_OUTPUT = "Lote_Produccion_Procesado.xlsx"
# Carpetas del buzón de entrada
INPUT_RAW_DIR = "input_raw"
ARCHIVE_RAW_DIR = os.path.join(INPUT_RAW_DIR, "archive")

# Palabras clave para excluir departamentos ajenos
EXCLUDED_KEYWORDS = ["25mm","Light Shield Molding","TKM","DOOR", "DRAWER", "TOEKICK", "Rip", "DWEP Panel_FF", "Shelf", "Overlay","FLTCROWN","Finished End","Universal Filler","Valance","Skin", "Baffle","P4896", "P34.5X96_1/4_FF", "FFS Panel"]
ALLOWED_EXCEPTIONS = ["WALL SIDE", "BASE SIDE", "BASE BOTTOM", "WALL BOTTOM", "OW BOTTOM", "DIVIDER","DWEP Universal Filler", "OW Side Adjustable Shelf", "TK-Bench Cabinet Door","TK-Bench Cabinet No Door" ] 

# Mapeo de colores/acabados a contracciones cortas de canto (Edge)
MAPA_COLOR = {
# 1. Tratamientos especiales y procesos (Primero los más específicos)
    "DOLPHIN_GREY_PAINTED": "DG-P",       
    "PERFECT_WHITE_PAINTED": "PW-P",
    "PERFECT_WHITE_VINYL2D": "PW-V2D",
    "CHARLESTON_GREY_VINYL2D": "CH.G.-V2D",
    "HONEY_MAPLE_STAINED": "HMS",
    # 2. Acabados y colores estándar
    "VELVET_BLUE": "VB",
    "CHARLESTON_GREY": "CH.G.",            
    "CARMEL": "CAR",
    "CINNAMON": "CNN",
    "DOLPHIN_GREY": "DG",
    "FOSSIL_GREY": "FG",
    "GIBRALTAR": "GBR",
    "JAVA": "JV",
    "SOMERSET": "SOM",
    "HONEY_MAPLE": "HM",
    "NATURAL_OAK": "NOK", 
    "PERFECT_WHITE": "PW",
    "WHITE_61": "W_61",                 # <--- AGREGAR AQUÍ
    "WHITE 61": "W_61",
    "WHITE_121": "PW",
    "VENETIAN": "VTN",
    "MAPLE": "M",    
    "MAPLE_111": "M",
    "BLACK": "BK",
    "SLATE": "SLT",
}


def extraer_color(nombre):
    """Detecta el color en el nombre y devuelve su acrónimo para la columna Edge."""
    if not isinstance(nombre, str):
        return ""
    nombre_norm = str(nombre).upper().replace(" ", "_")
    for color, acronimo in MAPA_COLOR.items():
        if color.replace(" ", "_") in nombre_norm:
            return acronimo
    return ""  # Vacío si es blanco genérico o no especifica

def unificar_y_archivar_pdfs():
    """
    Escanea la carpeta input_raw/. Si encuentra archivos PDF:
    1. Los une en orden en 'Lote_Actual.pdf'.
    2. Los mueve a input_raw/archive/ con timestamp para no duplicar mañana.
    Si no hay archivos en input_raw/, verifica si ya existe Lote_Actual.pdf en la raíz.
    """
    os.makedirs(INPUT_RAW_DIR, exist_ok=True)
    os.makedirs(ARCHIVE_RAW_DIR, exist_ok=True)

    # Buscar todos los PDFs sueltos en input_raw/ (ignora subcarpetas)
    archivos_pdf = [
        os.path.join(INPUT_RAW_DIR, f)
        for f in os.listdir(INPUT_RAW_DIR)
        if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(INPUT_RAW_DIR, f))
    ]

    if archivos_pdf:
        # Ordenar alfabéticamente para mantener el orden de páginas
        archivos_pdf.sort()
        print(f"Detectados {len(archivos_pdf)} archivo(s) en '{INPUT_RAW_DIR}'. Unificando...")

        # Unir todos los PDFs encontrados con pypdf
        merger = PdfWriter()
        for pdf_path in archivos_pdf:
            merger.append(pdf_path)

        with open(PDF_INPUT, "wb") as f_out:
            merger.write(f_out)
        merger.close()
        print(f"-> Archivo unificado creado: '{PDF_INPUT}'")

        # Mover los originales a input_raw/archive/ con fecha/hora
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        for pdf_path in archivos_pdf:
            nombre_archivo = os.path.basename(pdf_path)
            destino = os.path.join(ARCHIVE_RAW_DIR, f"{timestamp}_{nombre_archivo}")
            shutil.move(pdf_path, destino)
        print(f"-> Archivos originales archivados en '{ARCHIVE_RAW_DIR}'.")

    elif not os.path.exists(PDF_INPUT):
        print(f"Error Poka-Yoke: No hay archivos PDF en '{INPUT_RAW_DIR}' ni existe '{PDF_INPUT}' en la raíz.")
        exit()
    else:
        print(f"Aviso: Usando '{PDF_INPUT}' existente en la raíz (no se detectaron PDFs nuevos en '{INPUT_RAW_DIR}').")
def extraer_caras(linea):
    """Extrae el acabado con diagonal o el nombre tras la medida 4x8 / 5x8 del tablero."""
    # Descartar líneas técnicas de ARDIS y encabezados de tabla
    if any(k in linea.upper() for k in ["OPTIMIZER", "PAGE", "CUTTING", "DATA\\", "PART NO"]):
        return None
        
    # Descartar piezas de corte (líneas que inician con número de partida)
    if re.match(r'^\d+\s+[A-Za-z]', linea):
        return None

    # Purgar palabras operativas/máquinas para que no contaminen las caras
    linea_limpia = re.sub(r'\b(CNC|RIP|- Shelf|SHELF)\b', '', linea, flags=re.IGNORECASE).strip()

    # 1. Acabados con diagonal (ej. White/White, Fossil Grey/Fossil Grey)
    match_slash = re.search(
        r'([A-Za-z][A-Za-z0-9_]*(?:\s+[A-Za-z][A-Za-z0-9_]*)*\s*/\s*[A-Za-z][A-Za-z0-9_]*(?:\s+[A-Za-z][A-Za-z0-9_]*)*)',
        linea_limpia
    )
    if match_slash:
        return match_slash.group(1).strip()
    
    # 2. Acabados directos tras medida de tablero (ej. 4 X 8 Somerset 7)
    match_dim = re.search(r'\b[45]\s*[Xx]\s*[89]\s+([A-Za-z][A-Za-z0-9_ -]+?)(?:\s+\d+)?$', linea_limpia)
    if match_dim:
        return match_dim.group(1).strip()
    return None

def abreviar_caras(caras_texto):
    """
    Convierte 'Dolphin/Dolphin' -> 'DG/DG', 'Fossil Grey/Fossil Grey' -> 'FG/FG',
    'White/White' -> 'W/W', 'CHARLESTON_GREY/CHARLESTON_GREY' -> 'CH.G./CH.G.', etc.
    """
    if not caras_texto or caras_texto == "SIN_ESPECIFICAR":
        return caras_texto

    # 1. Mapeo mandatorio para caras de tablero (cubre variaciones con o sin 'GREY', blanco y crudo)
    alias_tableros = {
        "WHITE": "W",
        "PERFECT_WHITE": "W",
        "WHITE_121": "W",
        "WHITE_61": "W_61",             # <--- AGREGAR AQUÍ
        "WHITE 61": "W_61",
        "RAW": "RAW",
        "DOLPHIN": "DG",
        "DOLPHIN GREY": "DG",
        "DOLPHIN_GREY": "DG",
        "FOSSIL": "FG",
        "FOSSIL GREY": "FG",
        "FOSSIL_GREY": "FG",
        "CHARLESTON": "CH.G.",
        "CHARLESTON GREY": "CH.G.",
        "CHARLESTON_GREY": "CH.G.",
        "CARMEL": "CAR",
        "CINNAMON": "CNN",
        "SOMERSET": "SOM",
        "JAVA": "JV",
        "VELVET BLUE": "VB",
        "VELVET_BLUE": "VB",
        "VENETIAN": "VTN"
    }

    def obtener_sigla(lado):
        lado_limpio = lado.strip().upper()
        
        # Coincidencia directa contra tabla de acabados de tablero
        if lado_limpio in alias_tableros:
            return alias_tableros[lado_limpio]

        # Coincidencia contra MAPA_COLOR general (admitiendo espacio o guion bajo)
        lado_normalizado = lado_limpio.replace(" ", "_")
        for patron, sigla in MAPA_COLOR.items():
            if patron == lado_normalizado or patron in lado_normalizado:
                return sigla
        
        # Fallback para colores compuestos nuevos no registrados
        palabras = [p for p in re.split(r'[\s_]+', lado_limpio) if p]
        if len(palabras) > 1:
            return "".join([p[0] for p in palabras])
        elif len(palabras) == 1:
            p = palabras[0]
            return p[:2] if len(p) >= 2 else p
        return lado.strip()

    if "/" in caras_texto:
        partes = caras_texto.split("/")
        siglas = [obtener_sigla(p) for p in partes]
        return "/".join(siglas)
    else:
        sigla = obtener_sigla(caras_texto)
        return f"{sigla}/{sigla}"

    
def generar_pdf_con_marca_de_agua(excel_path, nombre_corrida, color_lote):
    """
    Exporta 'Catalogo_Lote' a PDF y estampa el nombre del color en
    blanco y negro a gran escala sobre filas blancas y cebra (over=True).
    """

    # 0. Asegurar existencia de la carpeta destino
    carpeta_pdf = "print_ready"
    os.makedirs(carpeta_pdf, exist_ok=True)
    excel_abs = os.path.abspath(excel_path)
    temp_pdf = os.path.abspath("temp_catalogo_base.pdf")
    final_pdf = os.path.abspath(os.path.join(carpeta_pdf, f"Catalogo_{nombre_corrida}_{color_lote}.pdf"))
    
    # 1. Exportación nativa a PDF desde Excel COM
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb_com = excel.Workbooks.Open(excel_abs)
        ws_com = wb_com.Worksheets("Catalogo_Lote")
        ws_com.ExportAsFixedFormat(0, temp_pdf)
        wb_com.Close(False)
    finally:
        excel.Quit()
        
    # 2. Generación del sello B&W de gran escala
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.saveState()
    
    # Centro exacto de hoja Carta Portrait (612 x 792 pt)
    can.translate(612 / 2.0, 792 / 2.0)
    can.rotate(53)
    
    # Tipografía grande y monocromática (Negro con 9% de opacidad)
    # Permite verse sobre la cebra gris #E5E5E5 sin interferir con la tinta de 14 pt
# Mapeo de repeticiones para cubrir la diagonal en nombres cortos
    patron_repeticion = {
        "BLUE": "BLUE   BLUE   BLUE",
        "PINK": "PINK   PINK   PINK",
        "GREEN": "GREEN   GREEN",
    }
    texto_marca = patron_repeticion.get(str(color_lote).upper(), str(color_lote).upper())

    # Calibrar tamaño de fuente para que 2 o 3 palabras quepan de esquina a esquina
    tamano_fuente = 74 if len(texto_marca) > 20 else 95
    can.setFont("Helvetica-Bold", tamano_fuente)
    can.setFillColor(colors.Color(0.0, 0.0, 0.0, alpha=0.09))
    can.drawCentredString(0, -15, texto_marca)
    
    can.restoreState()
    can.save()
    packet.seek(0)
    
    # 3. Fusión en capa superior para vencer los fondos opacos de Excel
    watermark_pdf = PdfReader(packet)
    watermark_page = watermark_pdf.pages[0]
    
    reader = PdfReader(temp_pdf)
    writer = PdfWriter()
    
    for page in reader.pages:
        # over=True proyecta la marca sobre celdas blancas y grises
        page.merge_page(watermark_page, over=True)
        writer.add_page(page)
        
    with open(final_pdf, "wb") as f_out:
        writer.write(f_out)
        
    if os.path.exists(temp_pdf):
        os.remove(temp_pdf)
        
    print(f"PDF generado exitosamente: '{final_pdf}' [B&W Grande - {color_lote}].")
def obtener_color_por_fecha(texto_corrida):
    """
    Determina el color de etiqueta según el día hábil (L-V).
    Patrón: Lunes=Azul, Martes=Rosa, Miércoles=Verde, Jueves=Naranja, Viernes=Morado.
    """
    colores_semana = {
        0: "BLUE",
        1: "PINK",     # 9-29 cae en Martes
        2: "GREEN",    # 9-30 cae en Miércoles
        3: "ORANGE",
        4: "LAVANDER"
    }
    
    # 1. Limpiar prefijos de corrida como 'r1-', 'r2-', 'R1_' para no confundir el número de corrida con el mes
    texto_limpio = re.sub(r'^[A-Za-z]+\d*[-_]', '', str(texto_corrida).strip())
    
    # 2. Buscar patrón Mes-Día con Año opcional (ej. '9-30-26', '9-30-2026' o solo '9-30')
    match = re.search(r'(\d{1,2})-(\d{1,2})(?:-(\d{2,4}))?', texto_limpio)
    if match:
        m = int(match.group(1))
        d = int(match.group(2))
        y_str = match.group(3)
        
        # Si no trae año en el nombre, tomar el año actual
        if y_str:
            y = int(y_str)
            if y < 100:
                y += 2000
        else:
            y = datetime.date.today().year
            
        try:
            fecha_lote = datetime.date(y, m, d)
            dia_sem = fecha_lote.weekday()
            if dia_sem in colores_semana:
                return colores_semana[dia_sem]
        except ValueError:
            pass
            
    # 3. Respaldo por defecto si no se pudo identificar la fecha
    dia_actual = datetime.date.today().weekday()
    return colores_semana.get(dia_actual, "Rosa")


def simplificar_partname(nombre, length="", width="", thickness=0):
    """
    Genera la versión compacta para piso eliminando sustratos, calibres mm
    y matrices de corte repetidas (L x W x T), preservando intactas las cotas en pulgadas.
    """
    if not isinstance(nombre, str):
        return ""
    texto = nombre
    
    # 1. Abreviar prefijo operativo
    texto = re.sub(r'\bFEP_FOIL\b', 'F.Foil', texto, flags=re.IGNORECASE)
    
    # 2. Remover sustratos, acabados y ruidos redundantes de catálogo
    palabras_ruido = [
        r'\bMDF\b',
        r'\bW/W\b', r'\bStr\b', r'\bStretch\b', r'Quarter\s+Bac\w*',
        r'\bFull\s+Bac\w*',
        r'\bFull\b'
    ]
    for r in palabras_ruido:
        texto = re.sub(r, '', texto, flags=re.IGNORECASE)
        # Reemplazar colores largos por su contracción de 2 letras
    for color in MAPA_COLOR.keys():
        patron_color = re.escape(color).replace(r'\_', r'[_\s]+')
        #texto = re.sub(rf'\b{color}\b', '', texto, flags=re.IGNORECASE)
        # Reemplazar colores admitiendo espacios y guiones bajos
        texto = re.sub(rf'\b{patron_color}\b', '', texto, flags=re.IGNORECASE)
        
    # 3. Remover calibres explícitos en milímetros (ej. 16mm, 19mm, 18 mm)
    texto = re.sub(r'\b\d+\s*mm\b', '', texto, flags=re.IGNORECASE)
    
    # 4. Remover matrices de corte en mm con x/X (ej. 865 x 354 x 18 o 610 x 305)
    # No toca pulgadas porque las pulgadas tienen comillas o letras (25.5W, 14.75"H)
    texto = re.sub(r'\b\d+(?:\.\d+)?\s*[xX]\s*\d+(?:\.\d+)?(?:\s*[xX]\s*\d+(?:\.\d+)?)?\b', '', texto)
    texto = re.sub(r'\b\d+(?:\.\d+)?\s*[xX]\b', '', texto)  # Corta colgantes tipo "2286 x"
    
    # 5. Remover valores idénticos a L, W o T cuando queden sueltos en mm
    # El lookahead negativo (?!["A-Za-z]) garantiza NO tocar 24"W, 24"H, 21"D, etc.
    for val in [length, width, thickness]:
        val_str = str(val).split('.')[0].strip()
        if val_str and val_str != '0':
            texto = re.sub(rf'\b{re.escape(val_str)}\b(?!["A-Za-z])', '', texto)
            
    # 6. Limpiar espacios múltiples y bordes
    return re.sub(r'\s+', ' ', texto).strip()
def formatear_faces(acabado):
    """Convierte acabados a iniciales por palabra (ej. Fossil Grey/Fossil Grey -> FG/FG)."""
    if not acabado or acabado == "SIN_ESPECIFICAR":
        return "N/A"
    
    if "/" in acabado:
        lados = acabado.split("/")
        acronimos = []
        for lado in lados:
            # Extrae secuencias alfabéticas y toma la primera letra en mayúscula
            palabras = re.findall(r'[A-Za-z]+', lado)
            if palabras:
                acronimos.append("".join(p[0].upper() for p in palabras))
        # Solo unir con diagonal si se detectaron letras válidas en ambas caras
        if len(acronimos) == 2:
            return "/".join(acronimos)
        elif len(acronimos) == 1:
            return acronimos[0]
        return "N/A"
    
    # En caso de una sola cara (ej. Somerset)
    palabras = re.findall(r'[A-Za-z]+', acabado)
    return "".join(p[0].upper() for p in palabras) if palabras else "N/A"
def extraer_grosor(texto):
    """Detecta el calibre en milímetros dentro del texto de la pieza."""
    match = re.search(r'(\d+)\s*mm', texto, flags=re.IGNORECASE)
    if match:
        return int(match.group(1))
    match_corte = re.search(r'x\s*(\d+)\s*$', texto.strip())
    if match_corte:
        return int(match_corte.group(1))
    return 0

def es_material_blanco(caras="", part_name=""):
    """
    Prioriza estrictamente el acabado de las caras del tablero.
    Si las caras son W/W (White/White), va a KOCH sin importar el color en la etiqueta.
    """
    caras_upper = str(caras).upper()
    
    # 1. Regla Mandatoria por Caras del Tablero
    if any(w in caras_upper for w in ["W/W", "WHITE/WHITE", "WHITE / WHITE", "PERFECT_WHITE", "W_61"]):
        return True
    if caras_upper not in ["SIN_ESPECIFICAR", "N/A", ""] and not any(w in caras_upper for w in ["W/W", "W_61"]):
        return False
        
    # 2. Respaldo por nombre solo si las caras no vinieron especificadas en el PDF
    nombre_upper = str(part_name).upper()
    colores_madera = [
        "FOSSIL", "SOMERSET", "MAPLE", "HONEY", "STAINED", 
        "GREY", "GRAY", "SLATE", "FOIL", "BLACK", "OAK", "GIBRALTAR", "JAVA", "CINNAMON", "CHARLESTON", 
        "VELVET", "BLUE", "VENETIAN"
    ]
    if any(c in nombre_upper for c in colores_madera):
        return False
        
    return "WHITE" in nombre_upper or "W/W" in nombre_upper or "W_61" in nombre_upper

def determinar_routing(part_name, thickness=0, caras="", es_cnc=False):
    """
    Calcula la secuencia de producción en piso (Routing) para una pieza
    a partir de su nombre, calibre (mm) y acabado/caras.
    """
    nombre = part_name.upper().strip()
    grosor = float(thickness) if thickness else 0.0
# -------------------------------------------------------------
    # 0. REGLA MAESTRA DE HOJA CNC (Tableros optimizados para Router CNC)
    # -------------------------------------------------------------
    if es_cnc:
        # Si la pieza requiere canteado estructural (>13mm en costados/repisas)
        if grosor > 13 and any(k in nombre for k in ["SIDE", "SHELF", "BOTTOM", "TOP", "DIVIDER"]):
            return "CNC-EDG"
        return "CNC"
    
    # -------------------------------------------------------------
    # 1. PANELES ESPECIALES (HASB-RF, SB)
    # -------------------------------------------------------------
    if any(k in nombre for k in ["HASB-RF", "SB-RF", "SB SIDE PANEL", "HASB"]):
        return "CNC-EDG"
# Bases especiales de microondas (ruteado complejo):
    if "MWC" in nombre and "BOTTOM" in nombre:
        return "CNC"

    # -------------------------------------------------------------
    # 2. RESPALDOS / ESPALDAS (Back)
    # -------------------------------------------------------------
    if "OW BACK" in nombre:
        # Respaldo abierto expuesto: requiere barrenado vertical y ensamble
        return "BHX-ABD"
        
    if "BACK" in nombre and "SIDE" not in nombre:
        # Respaldos regulares (Back, RF Back, RFB Back): directos a Kitting
        if any(k in nombre for k in ["BASE BACK", "WALL BACK", "MWC"]):
            return "BHX"
        return ""

    # -------------------------------------------------------------
    # 3. FILLERS Y ESQUINEROS (DWEP / Universal Filler)
    # -------------------------------------------------------------
    if any(k in nombre for k in ["DWEP", "UNIVERSAL FILLER", "FILLER"]):
        return "ABD" if grosor <= 13 else "ABD"

    # -------------------------------------------------------------
    # 4. TIRANTES Y REFUERZOS (Stretcher)
    # -------------------------------------------------------------
    if any(k in nombre for k in ["STRETCHER", "STR"]):
        return "" if grosor <= 13 else "ABD"

    # -------------------------------------------------------------
    # 5. DIVISIONES VERTICALES (Divider)
    # -------------------------------------------------------------
    if "DIVIDER" in nombre:
        if grosor <= 13:
            return ""
        # En calibres estructurales: W/W va a KOCH, color va a ABD
        return "EDG-KOCH" if es_material_blanco(caras=caras, part_name=nombre) else "EDG-ABD"

# -------------------------------------------------------------
    # 6. PISOS Y TECHOS (Bottom / Top)
    # -------------------------------------------------------------
    if any(k in nombre for k in ["BOTTOM", "TOP"]):
        # A) Excepción RFB (Base Bottom RFB PB 19mm): No va a ABD/KOCH, solo barrenado vertical
        if "RFB" in nombre:
            return "EDG-BHX" if grosor > 13 else "BHX"

        # B) Calibres delgados (12mm / 13mm)
        if grosor <= 13:
            if any(k in nombre for k in ["TK-", "TOEKICK", "TOP-BOTTOM", "BOTTOM TOP", "TOP BOTTOM"]):
                return "NTCH"
            return "BHX"

        # C) Tall Cabinet Bottoms (sustituye BHX por DADO; DADO va antes que EDG)
        if "TALL" in nombre:
            return "DADO-EDG-KOCH" if es_material_blanco(caras=caras, part_name=nombre) else "DADO-EDG-ABD"

        # D) Wall Bottoms & Open Wall Bottoms (van a DADO en lugar de BHX; DADO antes que EDG)
        if any(k in nombre for k in ["WALL BOTTOM", "OW BOTTOM"]) or ("WALL" in nombre and "BOTTOM" in nombre):
            return "DADO-EDG-KOCH" if es_material_blanco(caras=caras, part_name=nombre) else "DADO-EDG-ABD"

        # E) Bases con zoclo integrado gruesas
        if "TK-" in nombre or "TOEKICK" in nombre:
            return "EDG-KOCH"
            
        # F) Bases y techos regulares (Base Bottom, Wall Top, etc.)
        if es_material_blanco(caras=caras, part_name=nombre):
            return "EDG-BHX-KOCH"
        else:
            return "EDG-BHX-ABD"

    # -------------------------------------------------------------
    # 7. COSTADOS (Side / Tall Side)
    # -------------------------------------------------------------
    if "SIDE" in nombre:
        # Canteado (EDG) aplica para tableros de 13 mm o mayores
        if grosor > 13:
            return "EDG-BHX"
        else:
            return "BHX"

    # -------------------------------------------------------------
    # 8. REPISAS (Shelf / OW Shelf) — Provisional
    # -------------------------------------------------------------
    if "SHELF" in nombre:
        return "EDG"

    # Retorno de control para piezas no identificadas en la matriz
    return "REVISAR"
# ==========================================
# 1. PARSEO DIRECTO LÍNEA POR LÍNEA
# ==========================================

unificar_y_archivar_pdfs()
filas_extraidas = []
filas_omitidas = []
caras_actual = "SIN_ESPECIFICAR"
es_cnc_actual = False
nombre_corrida = "LOTE_ACTUAL"  # <--- Valor por defecto de seguridad

# Regex para partidas: [Part No] [PartName + Comodín D] [Length] [Width] [Qty]
patron_partida = re.compile(
    r'^(\d+)\s+(.+?)\s+(?:([A-Za-z](?:\s+[A-Za-z])?)\s+)?(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+)$'
)

with pdfplumber.open(PDF_INPUT) as pdf:
    for page in pdf.pages:
        texto = page.extract_text() or ""
        lineas = [l.strip() for l in texto.split("\n") if l.strip()]
        
        # Filtro Poka-Yoke: solo procesar páginas oficiales de corte
        if not any("PARTS_REVISED_NEW" in l.upper() for l in lineas[:3]):
            continue
    # Capturar la corrida de la cabecera si aún no se ha detectado
        if nombre_corrida == "LOTE_ACTUAL":
            for l in lineas[:3]:
                match_c = re.search(r'Data\\([A-Za-z0-9_-]+)\.R\d+', l, flags=re.IGNORECASE)
                if match_c:
                    nombre_corrida = match_c.group(1)
                    break
        # Detectar el material/caras de la página (habitualmente en línea 2 o 3)
        for l in lineas[:5]:
            # Si topa con el encabezado de columnas o una pieza, la página es continuación: DETENER
            if "PART NO" in l.upper() or patron_partida.match(l):
                break           
            acabado = extraer_caras(l)
            if acabado and "CUTTING" not in acabado.upper() and "PAGE" not in acabado.upper():
                caras_actual = abreviar_caras(acabado)
                es_cnc_actual = bool(re.search(r'\bCNC\b', l, re.IGNORECASE))
                break


                
        # Parsear partidas de corte útiles
        for l in lineas:
            match = patron_partida.match(l)

            if match:
                part_no = match.group(1)
                part_name = match.group(2).strip()
                # Grupo 3 corresponde al comodín 'D' (L, Y L) y se ignora intencionalmente
                length = match.group(4)
                width = match.group(5)
                qty = int(match.group(6))
                
                # Excluir piezas de otros departamentos (Doors, Drawers)
                if any(k.upper() in part_name.upper() for k in EXCLUDED_KEYWORDS):
                    name_upper = part_name.upper()
                    palabra_filtro = next((k for k in EXCLUDED_KEYWORDS if k.upper() in part_name.upper()), None)
                    if palabra_filtro:
                        es_excepcion = any(exc.upper() in name_upper for exc in ALLOWED_EXCEPTIONS)
                        if not es_excepcion:
                                filas_omitidas.append({
                                    "Part No": part_no,
                                    "PartName": part_name,
                                    "L": length,
                                    "W": width,
                                    "Qty": qty,
                                    "Faces": caras_actual,  # Guarda el acrónimo (ej. W/W o FG/FG)
                                    "Filtro": palabra_filtro.upper()
                                })
                                continue
                
                thickness = extraer_grosor(part_name)
                color_val = extraer_color(part_name)
                routing_teorico = determinar_routing(
                part_name=part_name, 
                thickness=thickness, 
                caras=caras_actual,
                es_cnc=es_cnc_actual
                
                )
                
                # 1. Tomar el último dígito del espesor (ej. 19 -> '9', 16 -> '6', 13 -> '3')
                thk_digito = str(thickness)[-1] if thickness > 0 else ""
                
                # 2. Tomar la primera letra mayúscula del nombre de la pieza
                mayusculas = re.findall(r'[A-Z]', part_name)
                primera_letra = mayusculas[0] if mayusculas else "X"
                
# 3. Clave compuesta temporal para desempate
                clave_compuesta = f"{part_no}{thk_digito}{primera_letra}"
                
                filas_extraidas.append({
                    "Clave_Compuesta": clave_compuesta,   # <--- Temporal para consolidar
                    "Part No": part_no,
                    "T": thickness,
                    "PartName": part_name,
                    "PartName_Simplified": simplificar_partname(part_name, length, width, thickness),
                    "Routing": routing_teorico,
                    "L": length,
                    "W": width,
                    "Qty": qty,
                    "Faces": caras_actual,
                    "Color": color_val                    # (O 'Edge' según tu nombre actual)
                })
    # Al terminar el parseo del PDF, calculamos el color correspondiente
color_lote = obtener_color_por_fecha(nombre_corrida)
ruta_record = os.path.join("record", f"Lote_{nombre_corrida}_{color_lote}.xlsx")
df = pd.DataFrame(filas_extraidas)
df_omitidos = pd.DataFrame(filas_omitidas)

if not df_omitidos.empty:
    df_omitidos = df_omitidos.groupby(["Part No", "PartName", "L", "W", "Filtro"], as_index=False).agg({
        "Qty": "sum"
    }).sort_values(by=["Filtro", "PartName"], ascending=[True, True])

if df.empty:
    print("Error: No se lograron extraer filas útiles. Verifica el archivo.")
    exit()

# Deduplicar por llave compuesta Unique y sumar Qty
# 1. Deduplicar por clave compuesta temporal y sumar Qty
df = df.groupby("Clave_Compuesta", as_index=False).agg({
    "Part No": "first",
    "T": "first",
    "PartName": "first",
    "PartName_Simplified": "first",
    "Routing": "first",
    "L": "first",
    "W": "first",
    "Qty": "sum",
    "Faces": "first",
    "Color": "first"                              # (O 'Edge' según tu nombre actual)
})

# 2. Generar 'Unique' inteligente:
# - Si el Part No es único en todo el catálogo -> Solo el número ("15", "28")
# - Si el Part No está repetido en el lote -> Clave compuesta ("423W", "429B")
# =========================================================================
# 2. GENERAR 'UNIQUE' INTELIGENTE (JERARQUÍA DE 3 ESCALONES POKA-YOKE)
# Escalón 1: 2 dígitos numéricos (Part No único en el catálogo)
# Escalón 2: 3 dígitos numéricos (Part No repetido, pero grosor determinante)
# Escalón 3: 4 caracteres con letra (Part No repetido Y mismo grosor)
# =========================================================================

# Frecuencia global del Part No
conteo_part_no = df["Part No"].value_counts()

# Construir clave numérica intermedia: Part No + último dígito del espesor (ej. 386, 389)
df["Thk_Digito"] = df["T"].apply(lambda t: str(int(t))[-1] if t > 0 else "")
df["Clave_Num"] = df["Part No"].astype(str) + df["Thk_Digito"]

# Frecuencia de la combinación Part No + Calibre
conteo_clave_num = df["Clave_Num"].value_counts()

# Conjunto de Part Nos originales para blindar colisiones cruzadas
part_nos_existentes = set(df["Part No"].astype(str))

def asignar_unique_jerarquico(r):
    p_no = str(r["Part No"])
    
    # Escalón 1: Si no tiene duplicados en todo el lote, número limpio
    if conteo_part_no[r["Part No"]] == 1:
        return p_no
    
    # Escalón 2: Si está duplicado pero el grosor es único para esa partida
    # (Poka-Yoke adicional: valida que la clave de 3 dígitos no coincida con otro Part No real)
    c_num = r["Clave_Num"]
    if conteo_clave_num[c_num] == 1 and c_num not in part_nos_existentes:
        return c_num
        
    # Escalón 3: El grosor no desempata (mismo Part No y mismo espesor) -> requiere letra
    return r["Clave_Compuesta"]

df["Unique"] = df.apply(asignar_unique_jerarquico, axis=1)

# Orden multinivel: 1° Calibre ascendente (mm), 2° Alfabético por nombre simplificado
df = df.sort_values(by=["T", "PartName"], ascending=[True, True])

# ==========================================
# 2. GENERACIÓN DEL ARCHIVO EXCEL DE PISO
# ==========================================
wb = Workbook()

font_piso = Font(name="Calibri", size=14, bold=False)
font_bold = Font(name="Calibri", size=14, bold=True)
font_header = Font(name="Calibri", size=11, bold=True, color="000000")
fill_header = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

# Borde estándar para datos
border_delgado = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

# Guía visual cada 5 filas (línea inferior gris más oscura)
border_cinco = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='medium', color='555555')
)

# Corte de sección por cambio de calibre (doble línea negra)
border_corte_grosor = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='double', color='000000')
)
# Borde normal estándar visible (cuadrícula clásica de Excel)
# Borde horizontal continuo para columnas ancla (sin paredes laterales)
border_normal_visible = Border(
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000')
)

# Relleno fijo para la columna ancla (Part No)
fill_ancla = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

# Cebra de bloques (+5% de contraste: #E5E5E5)
fill_cebra = PatternFill(start_color="E5E5E5", end_color="E5E5E5", fill_type="solid")
# Hoja 1: Catalogo_Lote
ws1 = wb.active
ws1.title = "Catalogo_Lote"

columnas_visibles = ["Unique", "PartName", "Part No","PartName_Simplified", "L", "W", "T","Faces","Color","Routing", "Qty"]
ws1.append(columnas_visibles)

# Salto de línea para compactar la columna B
idx_part_no = columnas_visibles.index("Part No") + 1
ws1.cell(row=1, column=idx_part_no, value="Part\nNo.")

idx_name_simp = columnas_visibles.index("PartName_Simplified") + 1
ws1.cell(row=1, column=idx_name_simp, value="PartName\nSimplified")

# Formato de cabecera con wrap_text=True
for col_idx in range(1, len(columnas_visibles) + 1):
    c = ws1.cell(row=1, column=col_idx)
    c.font = font_header
    c.fill = fill_header
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Altura calibrada a 32 pt para alojar los 2 renglones en 14 pt
ws1.row_dimensions[1].height = 32
# Repetir la fila 1 de encabezados en la parte superior de cada hoja física
# ==========================================
# CONFIGURACIÓN DE PÁGINA E IMPRESIÓN (PORTRAIT)
# ==========================================
# 1. Repetir la fila 1 de encabezados en cada hoja física
ws1.print_title_rows = '1:1'

# 2. Orientación vertical y tamaño de papel
ws1.page_setup.paperSize = ws1.PAPERSIZE_LETTER
ws1.page_setup.orientation = ws1.ORIENTATION_PORTRAIT

# 3. Márgenes estrechos (0.3 in laterales para maximizar el ancho útil)
ws1.page_margins.left = 0.25
ws1.page_margins.right = 0.25
ws1.page_margins.top = 0.4
ws1.page_margins.bottom = 0.25
ws1.page_margins.header = 0.2
ws1.page_margins.footer = 0

# 4. Forzar ajuste a 1 página de ancho (las filas fluyen libremente hacia abajo)
ws1.sheet_properties.pageSetUpPr.fitToPage = True
ws1.page_setup.fitToWidth = 1
ws1.page_setup.fitToHeight = 0
ws1.print_options.horizontalCentered = True

# 5. Encabezado de impresión nativo
ws1.oddHeader.left.text = f"&BOrder / Batch: {nombre_corrida}  |  Color: {color_lote}&B"
ws1.oddHeader.right.text = "Sheet &[Page] of &[Pages]"

# Altura ergonómica de encabezado
ws1.row_dimensions[1].height = 28

thickness_anterior = None
filas_matriz = df[columnas_visibles].values.tolist()
idx_t = columnas_visibles.index("T")
idx_name = columnas_visibles.index("PartName") + 1
idx_part_no = columnas_visibles.index("Part No") + 1
idx_qty = columnas_visibles.index("Qty") + 1

contador_bloque = 0

for r_idx, fila in enumerate(filas_matriz, start=2):
    thickness_actual = fila[idx_t]
    
    # 1. Fijar altura de 26 pt por fila para lectura descansada
    ws1.row_dimensions[r_idx].height = 26
    
    # 2. Si cambia el grosor, sobreescribir la fila anterior con borde doble y reiniciar cuenta
    if thickness_anterior is not None and thickness_actual != thickness_anterior:
        for c in range(1, len(columnas_visibles) + 1):
            ws1.cell(row=r_idx - 1, column=c).border = border_corte_grosor
        contador_bloque = 0
# Salto de página solicitado por Gloria:
        ws1.row_breaks.append(Break(id=r_idx - 1))
    ws1.append(fila)
    contador_bloque += 1
    
# 3. Asignar línea inferior más marcada cada 5 registros
    borde_actual = border_cinco if (contador_bloque % 5 == 0) else border_delgado
    es_bloque_gris = ((contador_bloque - 1) // 5) % 2 == 1
    es_centro = (contador_bloque % 5 == 3)
    es_fila_gris = (not es_bloque_gris) if es_centro else es_bloque_gris

    # Bucle único: aplica bordes, sombreados y alineación
    for c_idx in range(1, len(columnas_visibles) + 1):
        cell = ws1.cell(row=r_idx, column=c_idx)
        cell.font = font_piso
        cell.border = borde_actual

# 1. Bordes normales protegidos para Part No y Qty
        if c_idx in [idx_part_no, idx_qty]:
            cell.border = border_normal_visible
        else:
            cell.border = borde_actual
        
        # Sombreado: Gris fijo permanente en Part No (Columna 2), cebra en el resto
        if c_idx == 2:
            cell.fill = fill_ancla
# 2. Sombreado de fila corrida (incluye Part No y Qty)
        elif es_fila_gris:
            cell.fill = fill_cebra
        else:
            cell.fill = PatternFill(fill_type=None)
            
        # Alineación: texto a la izquierda en PartName, centrado en el resto
        if c_idx in [idx_name, idx_name_simp]:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="center", vertical="center")
    #for c_idx in range(1, len(columnas_visibles) + 1):
     #   cell = ws1.cell(row=r_idx, column=c_idx)
      #  cell.font = font_piso
       # cell.border = borde_actual
        #if c_idx == idx_name:
        #    cell.alignment = Alignment(horizontal="left", vertical="center")
        #else:
         #   cell.alignment = Alignment(horizontal="center", vertical="center")
            
    thickness_anterior = thickness_actual

 #Borde inferior de cierre para la última fila de datos
for c in range(1, len(columnas_visibles) + 1):
   if c not in [idx_part_no, idx_qty]:
       ws1.cell(row=len(filas_matriz) + 1, column=c).border = border_corte_grosor

ws1.print_area = f"C1:K{ws1.max_row}"
# ==========================================
# HOJA 2: Recut_List (Captura Reactiva Dinámica - 4 Secciones)
# ==========================================
ws2 = wb.create_sheet(title="Recut_List")

# ------------------------------------------
# CONFIGURACIÓN DE PÁGINA E IMPRESIÓN (LANDSCAPE)
# ------------------------------------------
ws2.page_setup.paperSize = ws2.PAPERSIZE_LETTER
ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE

ws2.page_margins.left = 0.25
ws2.page_margins.right = 0.25
ws2.page_margins.top = 0.50
ws2.page_margins.bottom = 0.25
ws2.page_margins.header = 0
ws2.page_margins.footer = 0

ws2.sheet_properties.pageSetUpPr.fitToPage = True
ws2.page_setup.fitToWidth = 1
ws2.page_setup.fitToHeight = 0
ws2.print_options.horizontalCentered = True

# Sin encabezados ni pies de página nativos
ws2.oddHeader.left.text = ""
ws2.oddHeader.right.text = ""
ws2.oddFooter.left.text = ""
ws2.oddFooter.right.text = ""

# Definición de las 4 secciones operativas de piso
secciones_recut = [
    {"nombre": "RECUTS",  "filas": 25, "auto_qty": False},
    {"nombre": "MISSING", "filas": 15, "auto_qty": True},
    {"nombre": "WIP",     "filas": 15, "auto_qty": True},
    {"nombre": "EXTRA",   "filas": 10, "auto_qty": False}
]

headers_recuts = [
    "Unique",
    "Qty",
    "Special\nFaces",
    "Part\nNo.",
    "PartName",
    "L",
    "W",
    "T",
    "Faces",
    "Color",
    "Routing",
    "Qty"
]

columnas_map = {
    4: 'C',  # Part No.          -> Catalogo_Lote Col C
    5: 'D',  # PartName_Simp     -> Catalogo_Lote Col D
    6: 'E',  # L                 -> Catalogo_Lote Col E
    7: 'F',  # W                 -> Catalogo_Lote Col F
    8: 'G',  # T                 -> Catalogo_Lote Col G
    9: 'H',  # Faces             -> Catalogo_Lote Col H
    10: 'I', # Color             -> Catalogo_Lote Col I
    11: 'J', # Routing           -> Catalogo_Lote Col J
}

current_row = 1

for idx_sec, sec in enumerate(secciones_recut):
    row_meta = current_row
    ws2.row_dimensions[row_meta].height = 28
    
    # 1. Metadatos de Cabecera (alineados a la zona de impresión D:L)
    ws2.merge_cells(f"D{row_meta}:E{row_meta}")
    c_title = ws2.cell(row=row_meta, column=4, value=sec["nombre"])
    c_title.font = Font(name="Calibri", size=16, bold=True)
    c_title.alignment = Alignment(horizontal="left", vertical="center")

    ws2.merge_cells(f"F{row_meta}:G{row_meta}")
    c_lbl = ws2.cell(row=row_meta, column=6, value="Date/Color:")
    c_lbl.font = font_bold
    c_lbl.alignment = Alignment(horizontal="right", vertical="center")

    ws2.merge_cells(f"H{row_meta}:L{row_meta}")
    c_val = ws2.cell(row=row_meta, column=8, value=f"{nombre_corrida} {color_lote}")
    c_val.font = font_piso
    c_val.alignment = Alignment(horizontal="left", vertical="center")

    # 2. Encabezados de Tabla (12 columnas)
    row_header = row_meta + 1
    ws2.row_dimensions[row_header].height = 26
    for col_idx, h in enumerate(headers_recuts, start=1):
        c = ws2.cell(row=row_header, column=col_idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.border = border_delgado
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 3. Filas de Datos
    start_data = row_header + 1
    end_data = start_data + sec["filas"] - 1

    for r in range(start_data, end_data + 1):
        ws2.row_dimensions[r].height = 24
        
        # Columna 1 (A): Unique
        cell_a = ws2.cell(row=r, column=1)
        cell_a.font = font_piso
        cell_a.border = border_delgado
        cell_a.alignment = Alignment(horizontal="center", vertical="center")
        
        # Columna 2 (B): Qty
        cell_b = ws2.cell(row=r, column=2)
        cell_b.font = font_piso
        cell_b.border = border_delgado
        cell_b.alignment = Alignment(horizontal="center", vertical="center")
        if sec["auto_qty"]:
            # MISSING y WIP: jala la cantidad total programada de corte
            cell_b.value = f'=IF(A{r}="","",_xlfn.XLOOKUP(TEXT(A{r}, "@"),Catalogo_Lote!A:A,Catalogo_Lote!K:K, ""))'
        else:
            # RECUTS y EXTRA: captura manual por el operador
            cell_b.value = None

        # Columna 3 (C): Special Faces (Manual opcional para acabados especiales)
        cell_c = ws2.cell(row=r, column=3)
        cell_c.font = font_piso
        cell_c.border = border_delgado
        cell_c.alignment = Alignment(horizontal="center", vertical="center")
        
        # Columnas 4 a 11 (D a K): Fórmulas XLOOKUP
        for col_idx, col_letra in columnas_map.items():
            cell = ws2.cell(row=r, column=col_idx)
            cell.value = f'=IF(A{r}="","",_xlfn.XLOOKUP(TEXT(A{r}, "@"),Catalogo_Lote!A:A,Catalogo_Lote!{col_letra}:{col_letra}, "NO ENCONTRADO"))'
            cell.font = font_piso
            cell.border = border_delgado
            if col_idx == 5:  # PartName alineado a la izquierda
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
        # Columna 12 (L): Qty Final Espejo
        cell_l = ws2.cell(row=r, column=12)
        cell_l.value = f'=IF(B{r}="","",B{r})'
        cell_l.font = font_piso
        cell_l.border = border_delgado
        cell_l.alignment = Alignment(horizontal="center", vertical="center")

    # 4. Fila de Cierre y Resumen (Totales)
    row_total = end_data + 1
    ws2.row_dimensions[row_total].height = 26
    for col_idx in range(1, 13):
        c_tot = ws2.cell(row=row_total, column=col_idx)
        c_tot.font = font_bold
        c_tot.border = border_corte_grosor
        c_tot.alignment = Alignment(horizontal="center", vertical="center")

    # Conteo de partidas en Columna D (Part No.)
    ws2.cell(row=row_total, column=4).value = f'=COUNT(D{start_data}:D{end_data})'
    
    # Etiqueta TOTAL en Columna E (PartName)
    c_tot_e = ws2.cell(row=row_total, column=5, value="TOTAL")
    c_tot_e.alignment = Alignment(horizontal="right", vertical="center")

    # Suma total de piezas en Columna L (Qty)
    ws2.cell(row=row_total, column=12).value = f'=SUM(L{start_data}:L{end_data})'

    # 5. Salto de página entre secciones (excepto la última)
    if idx_sec < len(secciones_recut) - 1:
        ws2.row_breaks.append(Break(id=row_total))
        current_row = row_total + 1
    else:
        current_row = row_total

# Área de impresión estricta: Columna D a L hasta la última fila de EXTRA
ws2.print_area = f"D1:L{current_row}"

# ------------------------------------------
# AJUSTE DE ANCHOS DINÁMICOS Y HERENCIA
# ------------------------------------------
for col in ws1.columns:
    col_letter = get_column_letter(col[0].column)
    header_val = str(ws1.cell(row=1, column=col[0].column).value or '')
    
    max_len = max(
        max(len(line) for line in str(cell.value or '').split('\n'))
        for cell in col
    )
    
    if any(k in header_val for k in ["Simplificado", "Simplified"]):
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.25), 28)
    elif header_val == "PartName":
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.15), 15)
    elif "Part" in header_val:
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.2) + 1, 7.5)
    elif "Qty" in header_val:
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.2) + 4, 9)
    elif header_val == "T":
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.1) + 2, 5.5)
    elif "Unique" in header_val:
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.1) + 1, 11)    
    elif "Routing" in header_val:
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.15), 18)
    elif "Color" in header_val:
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.2) + 3, 9)
    else:
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.1) + 3, 7.5)

# ws2 hereda anchos correspondientes desde ws1 respetando las 12 columnas
ws2.column_dimensions['A'].width = ws1.column_dimensions['A'].width # Unique
ws2.column_dimensions['B'].width = 8                                 # Qty 1
ws2.column_dimensions['C'].width = 16                                # Special Faces (Override)
ws2.column_dimensions['D'].width = ws1.column_dimensions['C'].width # Part No
ws2.column_dimensions['E'].width = max(int(ws1.column_dimensions['D'].width or 0), 32) # PartName_Simplified
ws2.column_dimensions['F'].width = ws1.column_dimensions['E'].width # L
ws2.column_dimensions['G'].width = ws1.column_dimensions['F'].width # W
ws2.column_dimensions['H'].width = ws1.column_dimensions['G'].width # T
ws2.column_dimensions['I'].width = ws1.column_dimensions['H'].width # Faces
ws2.column_dimensions['J'].width = ws1.column_dimensions['I'].width # Color
ws2.column_dimensions['K'].width = ws1.column_dimensions['J'].width # Routing
ws2.column_dimensions['L'].width = 8                                 # Qty 2 (espejo)
# HOJA 3: Omitidos_Auditoria (Para revisión con Gloria)
# ------------------------------------------
ws3 = wb.create_sheet(title="Omitidos_Auditoria")


ws3.merge_cells("A1:F1")
ws3["A1"] = "REVISIÓN DE COMPONENTES OMITIDOS / FILTRADOS DEL LOTE"
ws3["A1"].font = Font(name="Calibri", size=15, bold=True)
ws3["A1"].alignment = Alignment(horizontal="center")

# Bloque de validación y firma
ws3["A3"] = "Criterio de Exclusión:"
ws3["A3"].font = font_bold
ws3["B3"] = ", ".join(EXCLUDED_KEYWORDS)
ws3["B3"].font = font_piso

ws3["D3"] = "Revisado / Validado por:"
ws3["D3"].font = font_bold
ws3["E3"] = "Gloria / ____________________"
ws3["E3"].font = font_piso

headers_omitidos = ["Part No", "PartName Completo", "L", "W", "Total Qty", "Regla Activada"]
for idx, h in enumerate(headers_omitidos, start=1):
    c = ws3.cell(row=5, column=idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = Alignment(horizontal="center")

if not df_omitidos.empty:
    filas_omit_matriz = df_omitidos[["Part No", "PartName", "L", "W", "Qty", "Filtro"]].values.tolist()
    for r_idx, fila in enumerate(filas_omit_matriz, start=6):
        ws3.append(fila)
        for c_idx in range(1, 7):
            cell = ws3.cell(row=r_idx, column=c_idx)
            cell.font = font_piso
            cell.border = border_delgado
                        # Alineación: texto a la izquierda en PartName y PartName_Simplified, centrado en el resto
            if c_idx in [1, 3, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
# Ajuste de anchos para la hoja de auditoría
anchos_omitidos = {'A': 11, 'B': 50, 'C': 9, 'D': 9, 'E': 12, 'F': 18}
for col_let, ancho in anchos_omitidos.items():
    ws3.column_dimensions[col_let].width = ancho
wb.save(EXCEL_OUTPUT)
wb.save(ruta_record)
print(f"Éxito: {len(df)} partidas procesadas y guardadas en '{EXCEL_OUTPUT}'.")
generar_pdf_con_marca_de_agua(EXCEL_OUTPUT, nombre_corrida, color_lote)
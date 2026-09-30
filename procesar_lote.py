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
import win32com.client
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from pypdf import PdfReader, PdfWriter
def generar_pdf_con_marca_de_agua(excel_path, nombre_corrida, color_lote):
    """
    Exporta 'Catalogo_Lote' a PDF y estampa el nombre del color en
    blanco y negro a gran escala sobre filas blancas y cebra (over=True).
    """
    excel_abs = os.path.abspath(excel_path)
    temp_pdf = os.path.abspath("temp_catalogo_base.pdf")
    final_pdf = os.path.abspath(f"Catalogo_{nombre_corrida}_{color_lote}.pdf")
    
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
    can.setFont("Helvetica-Bold", 115)
    can.setFillColor(colors.Color(0.0, 0.0, 0.0, alpha=0.09))
    can.drawCentredString(0, -15, str(color_lote).upper())
    
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


PDF_INPUT = "Lote_Actual.pdf"
EXCEL_OUTPUT = "Lote_Produccion_Procesado.xlsx"

# Palabras clave para excluir departamentos ajenos
EXCLUDED_KEYWORDS = ["25mm","Light Shield Molding","TKM","DOOR", "DRAWER", "TOEKICK", "Rip", "DWEP Panel_FF", "Shelf", "Overlay","FLTCROWN","Finished End","Universal Filler","Valance","Skin", "Baffle","P4896", "P34.5X96_1/4_FF"]
ALLOWED_EXCEPTIONS = ["WALL SIDE", "BASE SIDE", "BASE BOTTOM", "WALL BOTTOM", "OW BOTTOM", "DIVIDER","DWEP Universal Filler"]
def extraer_caras(linea):
    """Extrae el acabado con diagonal o el nombre tras la medida 4x8 / 5x8 del tablero."""
    # Descartar líneas técnicas de ARDIS y encabezados de tabla
    if any(k in linea.upper() for k in ["OPTIMIZER", "PAGE", "CUTTING", "DATA\\", "PART NO"]):
        return None
        
    # Descartar piezas de corte (líneas que inician con número de partida)
    if re.match(r'^\d+\s+[A-Za-z]', linea):
        return None

    # 1. Acabados con diagonal (ej. White/White, Fossil Grey/Fossil Grey)
    match_slash = re.search(r'([A-Za-z][A-Za-z0-9_]*(?:\s+[A-Za-z0-9_]+)?\s*/\s*[A-Za-z][A-Za-z0-9_]*(?:\s+[A-Za-z0-9_]+)?)', linea)
    if match_slash:
        return match_slash.group(1).strip()
    
    # 2. Acabados directos tras medida de tablero de 1 dígito en pies (ej. 4 X 8 o 5 X 8 Somerset)
    match_dim = re.search(r'\b[45]\s*[Xx]\s*[89]\s+([A-Za-z][A-Za-z0-9_ -]+?)(?:\s+\d+)?$', linea)
    if match_dim:
        val = match_dim.group(1).strip()
        return re.sub(r'\b(RIP|- Shelf)\b', '', val, flags=re.IGNORECASE).strip()
    return None


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

def simplificar_partname(nombre):
    """Genera la versión compacta para etiquetas amarillas de recut."""
    if not isinstance(nombre, str):
        return ""
    texto = nombre
    # Abreviar prefijo
    texto = re.sub(r'\bFEP_FOIL\b', 'F.Foil', texto, flags=re.IGNORECASE)
    # Remover sustratos, acabados redundantes y ruido
    palabras_ruido = [
        r'\bPLY\b', r'\bPB\b', r'\bMDF\b', r'\bPERFECT_WHITE\b', r'\bHONEY_MAPLE_STAINED\b',
        r'\bHONEY_MAPLE\b', r'\bFOSSIL_GREY\b', r'\bSOMERSET\b', r'\bWHITE_121\b', r'\bMAPLE_111\b',
        r'\bW/W\b', r'\bStr\b', r'\bStretch\b', r'Quarter\s+Bac\w*', r'Half\s+Depth\s+Shelf',
        r'\b19mm\b', r'\b18mm\b', r'\b16mm\b', r'\b13mm\b', r'\b6mm\b',
        r'\bDado\b', r'\bPB_EP\b', r'\bSW_UM\b'
    ]
    for r in palabras_ruido:
        texto = re.sub(r, '', texto, flags=re.IGNORECASE)
    return re.sub(r'\s+', ' ', texto).strip()

# ==========================================
# 1. PARSEO DIRECTO LÍNEA POR LÍNEA
# ==========================================
filas_extraidas = []
filas_omitidas = []
caras_actual = "SIN_ESPECIFICAR"
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
                caras_actual = formatear_faces(acabado)
                break

            # Al terminar el parseo del PDF, calculamos el color correspondiente
        color_lote = obtener_color_por_fecha(nombre_corrida)
                
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
                
                # 1. Tomar el último dígito del espesor (ej. 19 -> '9', 16 -> '6', 13 -> '3')
                thk_digito = str(thickness)[-1] if thickness > 0 else ""
                
                # 2. Tomar la primera letra mayúscula del nombre de la pieza
                mayusculas = re.findall(r'[A-Z]', part_name)
                primera_letra = mayusculas[0] if mayusculas else "X"
                
                # 3. Clave ultra compacta (ej. 40 + 9 + R = '409R')
                unique_key = f"{part_no}{thk_digito}{primera_letra}"
                
                filas_extraidas.append({
                    "Unique": unique_key,
                    "Part No": part_no,
                    "T": thickness,
                    "PartName": part_name,
                    #"PartName_Simplified": simplificar_partname(part_name),
                    "L": length,
                    "W": width,
                    "Qty": qty,
                    "Faces": caras_actual
                })

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
df = df.groupby("Unique", as_index=False).agg({
    "Part No": "first",
    "T": "first",
    "PartName": "first",
    #"PartName_Simplified": "first",
    "L": "first",
    "W": "first",
    "Qty": "sum",
    "Faces": "first"
})

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

columnas_visibles = ["Unique","T", "Part No", "PartName", "Faces", "L", "W", "Qty"]
ws1.append(columnas_visibles)

# Salto de línea para compactar la columna B
idx_part_no = columnas_visibles.index("Part No") + 1
ws1.cell(row=1, column=idx_part_no, value="Part\nNo.")

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
        elif es_bloque_gris:
            cell.fill = fill_cebra

# 2. Sombreado de fila corrida (incluye Part No y Qty)
        if es_fila_gris:
            cell.fill = fill_cebra
        else:
            cell.fill = PatternFill(fill_type=None)
            
        # Alineación: texto a la izquierda en PartName, centrado en el resto
        if c_idx == idx_name:
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

ws1.print_area = f"C1:H{ws1.max_row}"
## Hoja 2: Recut_List
ws2 = wb.create_sheet(title="Recut_List")


meta = [
    ("A3", "Fecha / Color Lote:", "B3", nombre_corrida),
    ("D3", "Hora Entrega:", "E3", ""),
    ("A4", "Entregado a:", "B4", ""),
    ("D4", "Auditado por:", "E4", "Edwin Zarate")
]
for p1, t1, p2, val in meta:
    ws2[p1] = t1
    ws2[p1].font = font_bold
    ws2[p2] = val
    ws2[p2].font = font_piso

def render_seccion(ws, start_row, titulo):
    ws.cell(row=start_row, column=1, value=titulo).font = font_bold
    headers = ["Unique", "Part No", "PartName Simplificado", "Thk", "Solicitados", "Físicos", "Status / Notas"]
    for idx, h in enumerate(headers, start=1):
        c = ws.cell(row=start_row + 1, column=idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center")
    for r in range(start_row + 2, start_row + 6):
        for c in range(1, 8):
            cell = ws.cell(row=r, column=c)
            cell.border = border_delgado
            cell.font = font_piso
    return start_row + 7

# Llamar a las 3 secciones para dibujarlas en Recut_List:
r_sig = render_seccion(ws2, 6, "1. RECUTS SOLICITADOS (Faltante Físico vs. Teórico)")
r_sig = render_seccion(ws2, r_sig, "2. PENDIENTES (En cola de maquinado / CNC)")
r_sig = render_seccion(ws2, r_sig, "3. 'WANTED' (Pallets no localizados en piso)")

# ------------------------------------------
# AJUSTE DE ANCHOS CALIBRADO PARA FUENTE 14 PT
# ------------------------------------------
# ------------------------------------------
# AJUSTE DE ANCHOS CALIBRADO PARA FUENTE 14 PT
# ------------------------------------------
# Hoja 1: Catalogo_Lote
#for col in ws1.columns:
#    max_len = max(len(str(cell.value or '')) for cell in col)
#    col_letter = get_column_letter(col[0].column)
    
#    if col_letter == 'C':  # PartName (Columna 3): factor 1.35x para evitar cortes
#        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.35) + 4, 55)
#    else:
#        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.2) + 3, 10)
# ------------------------------------------
# AJUSTE DE ANCHOS ULTRA COMPACTO (MEDICIÓN DINÁMICA)
# ------------------------------------------
# Hoja 1: Catalogo_Lote (Ajuste dinámico calibrado para 14 pt)
for col in ws1.columns:
    col_letter = get_column_letter(col[0].column)
    header_val = str(ws1.cell(row=1, column=col[0].column).value or '')
    
    # Medir la longitud máxima de línea por celda (ignora el acumulado de saltos de línea)
    max_len = max(
        max(len(line) for line in str(cell.value or '').split('\n'))
        for cell in col
    )
    
    if "PartName" in header_val:
        # Dinamismo puro: escala proporcional a 14 pt sin cortes
        ws1.column_dimensions[col_letter].width = int(max_len * 1.30) 
    elif "Part" in header_val:
        # Columna compacta gracias al salto Part / No.
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.2) + 3, 7.5)
    elif "Qty" in header_val:
        # Espacio suficiente para 3 dígitos con márgenes laterales limpios
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.2) + 4, 9)
    elif header_val == "T":
        # Calibre compacto
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.1) + 2, 5.5)
    elif "Unique" in header_val:
        # Espacio holgado para claves del tipo '4019-RFB' en 14 pt
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.1) + 1, 11)    
    else:
        # Columnas estándar (Unique, L, W, Faces)
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.1) + 3, 7.5)
for col in ws2.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

## Ajuste automático de anchos de columna
#for hoja in [ws1, ws2]:
#    for col in hoja.columns:
#        max_len = max(len(str(cell.value or '')) for cell in col)
 #       col_letter = get_column_letter(col[0].column)
  #      hoja.column_dimensions[col_letter].width = max(max_len + 3, 12)
# ------------------------------------------
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
            if c_idx in [1, 3, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

# Ajuste de anchos para la hoja de auditoría
anchos_omitidos = {'A': 11, 'B': 50, 'C': 9, 'D': 9, 'E': 12, 'F': 18}
for col_let, ancho in anchos_omitidos.items():
    ws3.column_dimensions[col_let].width = ancho
wb.save(EXCEL_OUTPUT)
print(f"Éxito: {len(df)} partidas procesadas y guardadas en '{EXCEL_OUTPUT}'.")
generar_pdf_con_marca_de_agua(EXCEL_OUTPUT, nombre_corrida, color_lote)
import pdfplumber
import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PDF_INPUT = "Lote_Actual.pdf"
EXCEL_OUTPUT = "Lote_Produccion_Procesado.xlsx"

# Palabras clave para excluir departamentos ajenos
EXCLUDED_KEYWORDS = ["DOOR", "DRAWER", "TOEKICK", "Rip", "DWEP Panel_FF", "Shelf", "Overlay","FLTCROWN","Finished End","Universal Filler","Valance","Skin", "Baffle"]
ALLOWED_EXCEPTIONS = ["WALL SIDE", "BASE SIDE", "BASE BOTTOM", "WALL BOTTOM", "OW BOTTOM", "DIVIDER","DWEP Universal Filler"]

def extraer_caras(linea):
    """Extrae el acabado con diagonal o el nombre de acabado tras la medida del tablero."""
    match_slash = re.search(r'([A-Za-z0-9_]+(?:\s+[A-Za-z0-9_]+)?\s*/\s*[A-Za-z0-9_]+(?:\s+[A-Za-z0-9_]+)?)', linea)
    if match_slash:
        return match_slash.group(1).strip()
    
    match_dim = re.search(r'\d+\s*[Xx]\s*\d+\s+([A-Za-z0-9_ -]+?)(?:\s+\d+)?$', linea)
    if match_dim:
        val = match_dim.group(1).strip()
        return re.sub(r'\b(RIP|- Shelf)\b', '', val, flags=re.IGNORECASE).strip()
    return None

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
    texto = re.sub(r'\bFEP_FOIL\b', 'F.Foil', texto, flags=re.IGNORECASE)
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

patron_partida = re.compile(
    r'^(\d+)\s+(.+?)\s+(?:([A-Za-z](?:\s+[A-Za-z])?)\s+)?(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+)$'
)

with pdfplumber.open(PDF_INPUT) as pdf:
    for page in pdf.pages:
        texto = page.extract_text() or ""
        lineas = [l.strip() for l in texto.split("\n") if l.strip()]
        
        if not any("PARTS_REVISED_NEW" in l.upper() for l in lineas[:3]):
            continue
            
        for l in lineas[:4]:
            acabado = extraer_caras(l)
            if acabado and "CUTTING" not in acabado.upper() and "PAGE" not in acabado.upper():
                caras_actual = acabado
                break
                
        for l in lineas:
            match = patron_partida.match(l)
            if match:
                part_no = match.group(1)
                part_name = match.group(2).strip()
                length = match.group(4)
                width = match.group(5)
                qty = int(match.group(6))
                
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
                                "Filtro": palabra_filtro.upper()
                            })
                            continue
                
                thickness = extraer_grosor(part_name)
                unique_key = f"{part_no}{thickness if thickness > 0 else ''}"
                
                filas_extraidas.append({
                    "Unique": unique_key,
                    "Part No": part_no,
                    "T": thickness,
                    "PartName": part_name,
                    "L": length,
                    "W": width,
                    "Qty": qty,
                    "Caras": caras_actual
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

df = df.groupby("Unique", as_index=False).agg({
    "Part No": "first",
    "T": "first",
    "PartName": "first",
    "L": "first",
    "W": "first",
    "Qty": "sum",
    "Caras": "first"
})

df = df.sort_values(by=["T", "PartName"], ascending=[True, True])

# ==========================================
# 2. GENERACIÓN DEL ARCHIVO EXCEL DE PISO
# ==========================================
wb = Workbook()

font_piso = Font(name="Calibri", size=14, bold=False)
font_bold = Font(name="Calibri", size=14, bold=True)

# Encabezado actualizado: Gris oscuro y letras negras para ahorro de tinta
font_header = Font(name="Calibri", size=14, bold=True, color="000000")
fill_header = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

border_delgado = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

border_cinco = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='medium', color='555555')
)

border_corte_grosor = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='double', color='000000')
)
# Borde normal estándar visible (tipo "Todos los bordes" de Excel)
border_normal_visible = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000')
)
fill_ancla = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
fill_cebra = PatternFill(start_color="E5E5E5", end_color="E5E5E5", fill_type="solid")

# Hoja 1: Catalogo_Lote
ws1 = wb.active
ws1.title = "Catalogo_Lote"

columnas_visibles = ["Unique", "Part No", "PartName", "L", "W", "T", "Qty"]
ws1.append(columnas_visibles)

for col_idx in range(1, len(columnas_visibles) + 1):
    c = ws1.cell(row=1, column=col_idx)
    c.font = font_header
    c.fill = fill_header
    c.alignment = Alignment(horizontal="center", vertical="center")

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
    
    ws1.row_dimensions[r_idx].height = 26
    
    if thickness_anterior is not None and thickness_actual != thickness_anterior:
        for c in range(1, len(columnas_visibles) + 1):
            if c not in [idx_part_no, idx_qty]:
                ws1.cell(row=r_idx - 1, column=c).border = border_corte_grosor
        contador_bloque = 0

    ws1.append(fila)
    contador_bloque += 1
    
    borde_base = border_cinco if (contador_bloque % 5 == 0) else border_delgado
    es_bloque_gris = ((contador_bloque - 1) // 5) % 2 == 1

    for c_idx in range(1, len(columnas_visibles) + 1):
        cell = ws1.cell(row=r_idx, column=c_idx)
        cell.font = font_piso
        
        # Part No y Qty conservan bordes normales en todas sus celdas
        if c_idx in [idx_part_no, idx_qty]:
            cell.border = border_normal_visible
        else:
            cell.border = borde_base
        
        if c_idx == 2:
            cell.fill = fill_ancla
        elif es_bloque_gris:
            cell.fill = fill_cebra
            
        if c_idx == idx_name:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
    thickness_anterior = thickness_actual

for c in range(1, len(columnas_visibles) + 1):
   if c not in [idx_part_no,idx_qty]:
       ws1.cell(row=len(filas_matriz) + 1, column=c).border = border_corte_grosor

# Hoja 2: Recut_List
ws2 = wb.create_sheet(title="Recut_List")

meta = [
    ("A3", "Fecha / Color Lote:", "B3", "9-15 Pink"),
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

r_sig = render_seccion(ws2, 6, "1. RECUTS SOLICITADOS (Faltante Físico vs. Teórico)")
r_sig = render_seccion(ws2, r_sig, "2. PENDIENTES (En cola de maquinado / CNC)")
r_sig = render_seccion(ws2, r_sig, "3. 'WANTED' (Pallets no localizados en piso)")

# Hoja 3: Omitidos_Auditoria
ws3 = wb.create_sheet(title="Omitidos_Auditoria")

ws3.merge_cells("A1:F1")
ws3["A1"] = "REVISIÓN DE COMPONENTES OMITIDOS / FILTRADOS DEL LOTE"
ws3["A1"].font = Font(name="Calibri", size=15, bold=True)
ws3["A1"].alignment = Alignment(horizontal="center")

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

# Ajuste de anchos
for col in ws1.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    if col_letter == 'C':
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.35) + 4, 55)
    else:
        ws1.column_dimensions[col_letter].width = max(int(max_len * 1.2) + 3, 10)

for col in ws2.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

anchos_omitidos = {'A': 11, 'B': 50, 'C': 9, 'D': 9, 'E': 12, 'F': 18}
for col_let, ancho in anchos_omitidos.items():
    ws3.column_dimensions[col_let].width = ancho

wb.save(EXCEL_OUTPUT)
print(f"Éxito: {len(df)} partidas procesadas y guardadas en '{EXCEL_OUTPUT}'.")
import pdfplumber

PDF_INPUT = "Lote_Actual.pdf"

with pdfplumber.open(PDF_INPUT) as pdf:
    print(f"--- MAPA DE PÁGINAS (Total: {len(pdf.pages)}) ---")
    for i, page in enumerate(pdf.pages, start=1):
        texto = page.extract_text() or ""
        lineas = [l.strip() for l in texto.split("\n") if l.strip()]
        
        # Tomar la línea del título (usualmente la línea 2)
        titulo = lineas[1] if len(lineas) > 1 else "SIN TÍTULO"
        print(f"Pág {i:02d}: {titulo}")
        
        # Si la página contiene 'PARTS' o 'Revised', mostrar las primeras 4 líneas
        if any(k in titulo.upper() for k in ["PARTS", "REVISED"]) and "TOTAL" not in titulo.upper():
            print("   -> Muestra de datos:")
            for l in lineas[2:6]:
                print(f"      {l}")
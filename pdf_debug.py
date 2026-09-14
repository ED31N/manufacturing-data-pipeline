import pdfplumber

with pdfplumber.open("Lote_Actual.pdf") as pdf:
    print(f"Total de páginas detectadas: {len(pdf.pages)}")
    
    for i, page in enumerate(pdf.pages[:2]): # Inspeccionar solo las 2 primeras páginas
        print(f"\n--- PÁGINA {i+1} ---")
        texto = page.extract_text() or ""
        lineas = [l for l in texto.split("\n") if l.strip()]
        print("Primeras 5 líneas de texto:")
        for l in lineas[:5]:
            print(f"  [TXT] {l}")
        
        tablas = page.extract_tables()
        print(f"Tablas detectadas con extract_tables(): {len(tablas)}")
        if tablas:
            print(f"Filas en tabla 1: {len(tablas[0])}")
            print(f"Muestra fila 1: {tablas[0][0] if len(tablas[0]) > 0 else 'Vacía'}")
            print(f"Muestra fila 2: {tablas[0][1] if len(tablas[0]) > 1 else 'Vacía'}")
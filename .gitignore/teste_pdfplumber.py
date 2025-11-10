import pdfplumber

arquivo_pdf = "PONTO GERAL OUTUBRO.pdf"

with pdfplumber.open(arquivo_pdf) as pdf:
    print(f"Numero de paginas encontradas: {len(pdf.pages)}")
    page = pdf.pages[3]
    texto = page.extract_text()

    print("=== Texto extraido da primeira pagina ===")
    print(texto)
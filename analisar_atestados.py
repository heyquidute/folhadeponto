from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

def analisar_atestados(caminho_arquivo):
    # Carrega o workbook existente
    wb = load_workbook(caminho_arquivo)

    # Remove a aba "ATESTADOS" se já existir
    if "ATESTADOS" in wb.sheetnames:
        del wb["ATESTADOS"]

    # Cria nova aba ATESTADOS no início
    aba_atestados = wb.create_sheet("ATESTADOS", 0)
    aba_atestados.sheet_view.showGridLines = False

    # Estilos usados
    alinhamento_centro = Alignment(horizontal="center", vertical="center")
    fonte_negrito = Font(bold=True)
    borda_inferior = Border(bottom=Side(style="thin", color="000000"))
    preenchimento_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    # Cabeçalho da aba ATESTADOS
    cabecalho = ["Funcionário", "Data", "Detalhe"]
    aba_atestados.append(cabecalho)
    for cel in aba_atestados[1]:
        cel.font = fonte_negrito
        cel.alignment = alinhamento_centro
        cel.border = borda_inferior

    # Percorre cada aba de funcionário
    for nome_aba in wb.sheetnames:
        if nome_aba == "ATESTADOS":
            continue

        aba = wb[nome_aba]

        # Pega cabeçalhos da aba
        cabecalhos = [celula.value for celula in next(aba.iter_rows(min_row=1, max_row=1))]
        if "Data" not in cabecalhos or "Ocorrencia" not in cabecalhos:
            continue

        idx_data = cabecalhos.index("Data") + 1
        idx_ocorrencia = cabecalhos.index("Ocorrencia") + 1

        # Percorre as linhas da aba do funcionário
        for linha_celulas in aba.iter_rows(min_row=2):
            data = linha_celulas[idx_data - 1].value
            ocorrencia_raw = linha_celulas[idx_ocorrencia - 1].value
            if not ocorrencia_raw:
                continue

            ocorrencia = str(ocorrencia_raw).strip().upper()

            # Se for atestado
            if ocorrencia.startswith("007") or ocorrencia.startswith("ATESTADO"):
                # Adiciona à aba ATESTADOS
                aba_atestados.append([nome_aba, data, ocorrencia_raw])

                # Pinta a linha inteira de verde na aba do funcionário
                for cel in linha_celulas:
                    cel.fill = preenchimento_verde

    # Ajusta formatação da aba ATESTADOS
    for coluna in aba_atestados.columns:
        coluna_letra = coluna[0].column_letter
        max_len = max(len(str(cel.value)) if cel.value else 0 for cel in coluna)
        aba_atestados.column_dimensions[coluna_letra].width = max_len + 2

    for linha in aba_atestados.iter_rows(min_row=2):
        for cel in linha:
            cel.alignment = alinhamento_centro

    wb.save(caminho_arquivo)
    print(f"Arquivo salvo com sucesso:\n{caminho_arquivo}")

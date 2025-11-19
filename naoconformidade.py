from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

from str_to_time import str_para_tempo

def analisar_conformidade(caminho_arquivo):
    # Carrega o workbook existente
    wb = load_workbook(caminho_arquivo)

    # Remove a aba "OCORRÊNCIAS" se já existir
    if "OCORRÊNCIAS" in wb.sheetnames:
        del wb["OCORRÊNCIAS"]

    # Cria nova aba OCORRÊNCIAS no início
    aba_ocorrencias = wb.create_sheet("OCORRÊNCIAS", 0)
    aba_ocorrencias.sheet_view.showGridLines = False

    # Estilos
    center_align = Alignment(horizontal="center", vertical="center")
    bold_font = Font(bold=True)
    borda_inferior = Border(bottom=Side(style="thin", color="000000"))
    preenchimento_amarelo = PatternFill(start_color="FFF467", end_color="FFF467", fill_type="solid")
    preenchimento_azul = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")

    # Cabeçalho da aba OCORRÊNCIAS
    cabecalho = ["Funcionário", "Data", "Ocorrência","Informações da folha de ponto"]
    aba_ocorrencias.append(cabecalho)
    for cel in aba_ocorrencias[1]:
        cel.font = bold_font
        cel.alignment = center_align
        cel.border = borda_inferior

    # Percorre cada aba de funcionário
    for nome_aba in wb.sheetnames.copy():
        if nome_aba == "OCORRÊNCIAS":
            continue

        aba = wb[nome_aba]

        # Pega cabeçalhos da aba
        cabecalhos = [celula.value for celula in next(aba.iter_rows(min_row=1, max_row=1))]
        if "Data" not in cabecalhos or "Ocorrencia" not in cabecalhos or "Total" not in cabecalhos:
            continue

        idx_data = cabecalhos.index("Data") + 1
        idx_hrtot = cabecalhos.index("Hr Tot T")+1

        max_col = aba.max_column
        col_turno_manha = max_col - 3
        col_t_almoco = max_col - 2
        col_turno_tarde = max_col - 1
        col_total = max_col

        # Percorre as linhas da aba do funcionário
        for linha_celulas in aba.iter_rows(min_row=2):
            data = linha_celulas[idx_data - 1].value
            t_manha = str_para_tempo(linha_celulas[col_turno_manha - 1].value)
            t_almoco = str_para_tempo(linha_celulas[col_t_almoco - 1].value)
            t_tarde = str_para_tempo(linha_celulas[col_turno_tarde - 1].value)
            t_total = str_para_tempo(linha_celulas[col_total - 1].value)
            valor_hrtot = linha_celulas[idx_hrtot-1].value

            # TEMPO DE ALMOÇO MENOR QUE 1 HORA
            if t_almoco and t_almoco < timedelta(hours=1):
                for cell in linha_celulas:
                    cell.fill = preenchimento_amarelo
                aba_ocorrencias.append([nome_aba, data, "Almoço < 1h", f"Tempo de almoço: {t_almoco}"])
            
            # TEMPO DE ALMOÇO MAIOR QUE 1 HORA E 20 MINUTOS
            elif t_almoco and t_almoco > timedelta(hours=1, minutes=20):
                for cell in linha_celulas:
                    cell.fill = preenchimento_amarelo
                aba_ocorrencias.append([nome_aba, data, "Almoço > 1h20min", f"Tempo de almoço: {t_almoco}"])

            # TURNO COM MAIS DE 6 HORAS SEM INTERVALO
            if t_manha and t_manha > timedelta(hours=6):
                for cell in linha_celulas:
                    cell.fill = preenchimento_amarelo
                aba_ocorrencias.append([nome_aba, data, "Turno da manhã > 6h", f"Duração do turno: {t_manha}"])

            if t_tarde and t_tarde > timedelta(hours=6):
                for cell in linha_celulas:
                    cell.fill = preenchimento_amarelo
                aba_ocorrencias.append([nome_aba, data, "Turno da tarde > 6h", f"Duração do turno: {t_tarde}"])
            
            # JORNADA TOTAL ACIMA DE 10 HORAS
            if t_total and t_total > timedelta(hours=10):
                for cell in linha_celulas:
                    cell.fill = preenchimento_amarelo
                aba_ocorrencias.append([nome_aba, data, "Jornada > 10h", f"Jornada total: {t_total}"])

            #ERRO NA BATIDA
            try:
                negativo = "-" in str(valor_hrtot) or valor_hrtot.startswith("−")
            except:
                negativo = False

            if negativo:
                for cel in linha_celulas:
                    cel.fill = preenchimento_azul

                aba_ocorrencias.append([nome_aba,data,"Diferente de 4 batidas",""])

    # Ajusta formatação da aba OCORRÊNCIAS
    for coluna in aba_ocorrencias.columns:
        coluna_letra = coluna[0].column_letter
        max_len = max(len(str(cel.value)) if cel.value else 0 for cel in coluna)
        aba_ocorrencias.column_dimensions[coluna_letra].width = max_len + 2

    for linha in aba_ocorrencias.iter_rows(min_row=2):
        for cel in linha:
            cel.alignment = center_align

    wb.save(caminho_arquivo)
    print(f"Arquivo salvo com sucesso:\n{caminho_arquivo}")

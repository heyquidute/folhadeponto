import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from datetime import datetime, timedelta
import re
import os

# Configurações de estilo
fill_amarelo = PatternFill(start_color="F9E700", end_color="F9E700", fill_type="solid")
fill_laranja = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
fill_vermelho = PatternFill(start_color="FF7F7F", end_color="FF7F7F", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center")
bold_font = Font(bold=True)

def str_para_tempo(valor):
    # Converte string "HH:MM" para objeto datetime.time
    if not valor:
        return None
    try:
        h, m = map(int, re.findall(r'\d+', valor))
        return timedelta(hours=h, minutes=m)
    except:
        return None   
    
def analisar_folha(caminho_excel):
    print(f"Analisando arquivo: {os.path.basename(caminho_excel)}")

    wb = openpyxl.load_workbook(caminho_excel)

    if "RESUMO" in wb.sheetnames:
        del wb["RESUMO"]
    
    resumo = wb.create_sheet("RESUMO")
    wb.move_sheet(resumo, offset=-(len(wb.sheetnames)-1)) # mover o RESUMO para o começo

    resumo.append(["Funcionário", "Data", "Tipo de Erro", "Detalhes"])
    for col in resumo.iter_cols(min_row=1, max_row=1):
        for cell in col:
            cell.font = bold_font
            cell.alignment = center_align

    
    # Central da análise
    for nome_aba in wb.sheetnames:
        if nome_aba == "RESUMO":
            continue

        ws = wb[nome_aba]

        for row in ws.iter_rows(min_row=2):
            try:
                data = row[0].value
                hr_ent_m = row[2].value
                hr_sai_m = row[3].value
                hr_ent_t = row[4].value
                hr_sai_t = row[5].value
                hr_tot_t = row[6].value

                t_ent_m = str_para_tempo(hr_ent_m)
                t_sai_m = str_para_tempo(hr_sai_m)
                t_ent_t = str_para_tempo(hr_ent_t)
                t_sai_t = str_para_tempo(hr_sai_t)
                t_tot_t = str_para_tempo(hr_tot_t)

                # Cálculo de jornada total
                if t_tot_t and t_tot_t > timedelta(hours=10):
                    for cell in row:
                        cell.fill = fill_vermelho
                    resumo.append([nome_aba, data, "Jornada > 10h", f"Duração total: {t_tot_t}"])
                
                # Almoço menor que 1 hora
                if t_sai_m and t_ent_t:
                    almoco = t_ent_t - t_sai_m
                    if almoco < timedelta(hours=1):
                        for cell in row:
                            cell.fill = fill_amarelo
                        resumo.append([nome_aba, data, "Almoço < 1h", f"Duração do almoço: {almoco}"])

                # Turno > 6h sem intervalo
                if t_ent_m and t_sai_m and not t_ent_t and not t_sai_t:
                    turno = t_sai_m - t_ent_m
                    if turno > timedelta(hours=6):
                        for cell in row:
                            cell.fill = fill_laranja
                        resumo.append([nome_aba, data, "Turno > 6h sem intervalo", f"Duração do turno: {turno}"])   

            except Exception as e:
                print(f"Erro ao analisar linha {row[0].row} na aba {nome_aba}: {e}")
                continue
        
    # Centraliza todo o conteúdo
    for row in ws.iter_rows():
            for cell in row:
                cell.alignment = center_align
    
    # ajusta a largura das colunas
    for coluna in resumo.columns:
        max_length = 0
        coluna_letra = coluna[0].column_letter
        for cell in coluna:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ajuste = (max_length + 2)
        resumo.column_dimensions[coluna_letra].width = ajuste


    wb.save(caminho_excel)
    print("Analise concluida e arquivo salvo.")
import tabula
import pandas as pd

#nome do arquivo pdf
arquivo_pdf = "PONTO GERAL OUTUBRO.pdf"

# teste: ler o pdf e ver e retornar as tabelas encontradas
dfs = tabula.read_pdf(
    arquivo_pdf,
    pages=1,
    lattice=True,
    multiple_tables=True
)
print(f"Numero de TABELAS encontradas: {len(dfs)}")

if dfs:
    for i, df in enumerate(dfs):
        print(f"------Tabela {1+i}-----")

tabula.convert_into(
    arquivo_pdf,
    "debug.csv",
    pages=1,
    lattice=True
)
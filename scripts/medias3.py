import pandas as pd
import numpy as np

# ===== arquivo de entrada =====
arquivo = "output/medias2.xlsx"
xls = pd.ExcelFile(arquivo)

# intervalos fixos
intervalos = ["dez-fev", "jan-mar", "mai-jul", "jun-ago"]

# acumulador: intervalo -> lista de Series (uma por ano)
acumulado = {intervalo: [] for intervalo in intervalos}

nomes_variaveis = None
unidades_variaveis = None

# ===== percorrer todas as abas (anos) =====
for aba in xls.sheet_names:
    df = pd.read_excel(
        arquivo,
        sheet_name=aba,
        header=None
    )

    # linha 0 -> nomes dos dados meteorológicos
    # linha 1 -> unidades
    nomes = df.iloc[0, 1:].copy()
    unidades = df.iloc[1, 1:].copy()

    # guardar referência UMA ÚNICA VEZ
    if nomes_variaveis is None:
        nomes_variaveis = nomes
        unidades_variaveis = unidades

    # dados começam na linha 2
    dados = df.iloc[2:, 1:]
    dados.index = df.iloc[2:, 0]  # dez-fev, jan-mar, etc.

    # acumular valores por intervalo
    for intervalo in intervalos:
        if intervalo in dados.index:
            acumulado[intervalo].append(dados.loc[intervalo])

# ===== calcular média final entre todos os anos =====
resultado = {}

for intervalo, listas in acumulado.items():
    df_concat = pd.concat(listas, axis=1).T
    resultado[intervalo] = df_concat.mean(skipna=True).round(2)

df_final = pd.DataFrame.from_dict(resultado, orient="index")

# ===== escrever Excel final com nomes + unidades =====
with pd.ExcelWriter("output/medias3.xlsx", engine="openpyxl") as writer:
    df_final.to_excel(
        writer,
        sheet_name="media3",
        startrow=2,
        header=False
    )

    ws = writer.sheets["media3"]

    # coluna A
    ws.cell(row=1, column=1, value="Intervalo")

    # linha 1 -> nomes dos dados meteorológicos (linha 1 do medias2_filtrado)
    for col_idx, nome in enumerate(nomes_variaveis, start=2):
        ws.cell(row=1, column=col_idx, value=nome)

    # linha 2 -> unidades correspondentes (linha 2 do medias2_filtrado)
    for col_idx, unidade in enumerate(unidades_variaveis, start=2):
        ws.cell(row=2, column=col_idx, value=unidade)

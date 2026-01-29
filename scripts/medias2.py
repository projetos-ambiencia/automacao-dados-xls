# medias2.py

import numpy as np
import pandas as pd
from leitura import ler_planilha
from config import PADRAO_ARQUIVO, ANOS_VALIDOS

# Variáveis oficiais (ordem e nome fixos)
VARIAVEIS = [
    "TMAX",
    "TMIN",
    "UR MAX",
    "UR MIN",
    "Vento",
]

INTERVALOS = {
    "dez-fev": ("dez", "jan", "fev"),
    "jan-mar": ("jan", "fev", "mar"),
    "mai-jul": ("mai", "jun", "jul"),
    "jun-ago": ("jun", "jul", "ago"),
}


def media_valores(valores):
    valores = [v for v in valores if not np.isnan(v)]
    if len(valores) == 0:
        return np.nan
    return round(float(np.mean(valores)), 2)


with pd.ExcelWriter("output/medias2.xlsx", engine="openpyxl") as writer:

    for ano in ANOS_VALIDOS:
        # carregar ano atual e anterior
        col_atual, unidades_atual, dados_atual = ler_planilha(
            PADRAO_ARQUIVO.format(ano=ano)
        )
        col_ant, unidades_ant, dados_ant = ler_planilha(
            PADRAO_ARQUIVO.format(ano=ano - 1)
        )

        resultado = {}

        for intervalo, meses in INTERVALOS.items():
            linha = {}

            for var in VARIAVEIS:
                valores = []

                for mes in meses:
                    if mes == "dez":
                        fonte = dados_ant
                    else:
                        fonte = dados_atual

                    if mes in fonte and var in fonte[mes]:
                        v = fonte[mes][var]
                        if not np.isnan(v):
                            valores.append(v)

                linha[var] = media_valores(valores)

            resultado[intervalo] = linha

        # DataFrame FINAL — colunas FIXAS e SEM confusão
        df = pd.DataFrame.from_dict(resultado, orient="index")
        df = df.reindex(columns=VARIAVEIS)

        sheet_name = str(ano)

        df.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=2,
            header=False
        )

        ws = writer.sheets[sheet_name]

        # Linha 1 — nomes das variáveis
        ws.cell(row=1, column=1, value="Intervalo")
        for col_idx, nome in enumerate(VARIAVEIS, start=2):
            ws.cell(row=1, column=col_idx, value=nome)

        # Linha 2 — unidades (do ano atual, se existir)
        unidades_dict = dict(zip(col_atual, unidades_atual))
        for col_idx, nome in enumerate(VARIAVEIS, start=2):
            ws.cell(row=2, column=col_idx, value=unidades_dict.get(nome, ""))

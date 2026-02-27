# automacao-dados-xls
este projeto calcula medias anuais e dos ultimos 30 anos da estacao meteorologica de Piracicaba (ESALQ), a analise e calculos sao feitos especificamente para o formato de dados da estação, quaisquer mudancas requerem adequacao do codigo ou da modelagem dos dados


passo a passo de execucao:

1 - python scripts/medias2.py
2 - python scripts/media3.py

3 - usar planilha de direca_vento.xlsx para calcular a media da direcao do vento em graus por intervalo de meses (os anos depois de 2018 nao possuem dados, portanto o intervalo foi de 1998 a 2018)
4 - usar planilha de radiacao_global.xlsx para calcular a media da radiacao global com conversao de (cal/cm².d) para (W/m²) usando periodo de insolacao em (h/dia) por intervalo de meses (de 1998 a 2025)

5 - adicionar os campos de direcao e radiacao no arquivo medias3.xlsx de acordo com seus respectivos intervalos de meses
6 - usar medias3.xlsx para ver as medias finais durante todo o periodo selecionado


# IMPORTANDO BIBLIOTECAS
import yfinance as yf
import pandas as pd
import os
from ferramentas import calcula_operacoes, manipulando_e_tratando, enviando_email

# DEFININDO TOKENS
tokens = ['BTC-USD', 'ETH-USD', 'SOL-USD', 'AAVE-USD']

# LAÇO DOS TOKENS
for token in tokens:

    # IMPORTANDO TOKEN & RETIRANDO DADOS NULOS
    token_dados = yf.download(token)
    token_dados = token_dados.dropna()
    token_dados = manipulando_e_tratando(token_dados)

    # CALCULANDO SITUAÇÕES DE COMPRA & VENDA
    data_compra, data_venda = calcula_operacoes(token_dados)

    # FILTRANDO DADOS & ENVIANDO PARA O EMAIL
    enviando_email(token, token_dados, data_compra, data_venda)

    # EXCLUINDO ANEXOS
    os.remove(f'{token}.png')
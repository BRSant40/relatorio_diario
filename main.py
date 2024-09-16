# IMPORTANDO BIBLIOTECAS
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplcyberpunk
import win32com.client as win32
import os
from datetime import datetime, timedelta
from ferramentas import calcula_operacoes

# DEFININDO TOKENS
tokens = ['BTC-USD', 'ETH-USD', 'SOL-USD', 'AAVE-USD']

# LAÇO DOS TOKENS
for token in tokens:

    # IMPORTANDO TOKEN & RETIRANDO DADOS NULOS
    token_dados = yf.download(token)
    token_dados = token_dados.dropna()


    # CALCULANDO RETORNOS
    token_dados['retornos'] = token_dados['Adj Close'].pct_change().dropna()


    # CALCULANDO RETORNOS POSITIVOS & NEGATIVOS
    token_dados['retornos_postivos'] = token_dados['retornos'].apply(lambda x: x if x > 0 else 0)
    token_dados['retornos_negativos'] = token_dados['retornos'].apply(lambda x: abs(x) if x < 0 else 0)


    # CALCULANDO MÉDIA DOS RETORNOS POSITIVOS & NEGATIVOS
    token_dados['media_retornos_positivos'] = token_dados['retornos_postivos'].rolling(window=14).mean()
    token_dados['media_retornos_negativos'] = token_dados['retornos_negativos'].rolling(window=14).mean()
    token_dados = token_dados.dropna()


    # CALCULANDO RSI
    token_dados['RSI'] = (100 - 100 / (1 + token_dados['media_retornos_positivos'] / token_dados['media_retornos_negativos']))


    # CALCULANDO MÉDIA 14 SEMANAL
    token_dados['EMA_14'] = token_dados['Close'].ewm(span=98, adjust=False, min_periods=0).mean()


    # CALCULANDO SITUAÇÕES DE COMPRA
    token_dados.loc[token_dados['Close'] > token_dados['EMA_14'], 'compra'] = 'sim'  # Se a cotação for maior que a EMA 14, COMPRA!!
    token_dados.loc[token_dados['RSI'] >= 55, 'compra'] = 'sim'  # Se o RSI for maior ou igual a 55, COMPRA!!
    token_dados.loc[token_dados['RSI'] < 55, 'compra'] = 'nao'  # Se o RSI for menor que 55, NÃO COMPRA!!
    token_dados.loc[token_dados['Close'] < token_dados['EMA_14'], 'compra'] = 'nao'  # Se a cotação for menor que o EMA 14, NÃO COMPRA!!
    token_dados.loc[token_dados['RSI'] > 70, 'compra'] = 'nao'  # Se o RSI for maior que 70, NÃO COMPRA!!


    # CALCULANDO SITUAÇÕES DE COMPRA & VENDA
    data_compra, data_venda = calcula_operacoes(token_dados)


    # FILTRANDO DADOS & ENVIANDO PARA O EMAIL
    cotacao_dia = token_dados.iloc[-1]
    retorno = cotacao_dia["retornos"]

    outlook = win32.Dispatch("outlook.application")

    email = outlook.CreateItem(0)

    email.To = "bruninho_123vini@hotmail.com"
    email.Subject = "Relatório de Mercado"
    email.Body = f"""Segue o Relatório das Criptos:

    * {token}
    - Cotação do dia: {cotacao_dia['Close']:.2f}
    - Variação do Dia: {retorno * 100:.2f}%
    - RSI: {cotacao_dia['RSI']:.0f}
    - Média Móvel: {cotacao_dia['EMA_14']:.0f}
    - Compra: {cotacao_dia['compra']}
    - Última Compra: {data_compra[-1]}
    - Última Venda: {data_venda[-1]}
    """


    ########################## PLOTANDO GRÁFICOS ################################

    plt.figure(figsize=(12, 5))
    plt.title(f"{token}")
    plt.scatter(token_dados.loc[data_compra].index, token_dados.loc[data_compra]['Adj Close'], marker='^',
                c='g')
    plt.scatter(token_dados.loc[data_venda].index, token_dados.loc[data_venda]['Adj Close'], marker='^',
                c='r')
    plt.plot(token_dados['Adj Close'], alpha=0.7)

    plt.savefig(f"{token}.png")


    ########################### ANEXANDO GRÁFICOS ##############################

    anexo = fr"C:\Users\brsan\PycharmProjects\Relatorio_diario\{token}.png"

    email.Attachments.Add(anexo)
    email.Send()

    ########################### EXCLUINDO ANEXOS ################################
    os.remove(f'{token}.png')
import pandas as pd
from datetime import datetime, timedelta

def calcula_operacoes(token_dados):
    ordem_aberta = False  # Use booleano para clareza
    data_compra = []
    data_venda = []

    for i in range(len(token_dados) - 1):  # Evita erro de índice no último dia
        # COMPRAS
        if "sim" in str(token_dados['compra'].iloc[i]).lower() and not ordem_aberta:  # Converte para string para evitar erros de tipo
            data_compra.append(token_dados.index[i + 1])
            ordem_aberta = True

        # VENDAS
        if token_dados['RSI'].iloc[i] < 40 and token_dados['Close'].iloc[i] < token_dados['EMA_14'].iloc[i] and ordem_aberta:
            data_venda.append(token_dados.index[i + 1])
            ordem_aberta = False

    # Trata o caso de haver uma compra no último dia e não ter venda
    if token_dados['RSI'].iloc[-1] < 40 and token_dados['Close'].iloc[-1] < token_dados['EMA_14'].iloc[-1] and ordem_aberta:
        data_amanha = datetime.now() + timedelta(days=1)
        data_venda.append(data_amanha)

    return data_compra, data_venda



import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplcyberpunk
import win32com.client as win32

btc = "BTC-USD"
eth = "ETH-USD"
sol = "SOL-USD"

btc_dados = yf.download(btc, period = "6mo")
eth_dados = yf.download(eth, period = "6mo")
sol_dados = yf.download(sol, period = "6mo")

# EXCLUINDO DADOS NULOS
btc_dados = btc_dados.dropna()
eth_dados = eth_dados.dropna()
sol_dados = sol_dados.dropna()

# CALCULANDO RETORNOS
btc_dados['retornos'] = btc_dados['Adj Close'].pct_change().dropna()
eth_dados['retornos'] = eth_dados['Adj Close'].pct_change().dropna()
sol_dados['retornos'] = sol_dados['Adj Close'].pct_change().dropna()

# CALCULANDO RETORNOS POSITIVOS & NEGATIVOS
btc_dados['retornos_postivos'] = btc_dados['retornos'].apply(lambda x: x if x > 0 else 0)
btc_dados['retornos_negativos'] = btc_dados['retornos'].apply(lambda x: abs(x) if x < 0 else 0)

eth_dados['retornos_postivos'] = eth_dados['retornos'].apply(lambda x: x if x > 0 else 0)
eth_dados['retornos_negativos'] = eth_dados['retornos'].apply(lambda x: abs(x) if x < 0 else 0)

sol_dados['retornos_postivos'] = sol_dados['retornos'].apply(lambda x: x if x > 0 else 0)
sol_dados['retornos_negativos'] = sol_dados['retornos'].apply(lambda x: abs(x) if x < 0 else 0)

# CALCULANDO MÉDIA DOS RETORNOS POSITIVOS & NEGATIVOS
btc_dados['media_retornos_positivos'] = btc_dados['retornos_postivos'].rolling(window = 14).mean()
btc_dados['media_retornos_negativos'] = btc_dados['retornos_negativos'].rolling(window = 14).mean()
btc_dados = btc_dados.dropna()

eth_dados['media_retornos_positivos'] = eth_dados['retornos_postivos'].rolling(window = 14).mean()
eth_dados['media_retornos_negativos'] = eth_dados['retornos_negativos'].rolling(window = 14).mean()
eth_dados = eth_dados.dropna()

sol_dados['media_retornos_positivos'] = sol_dados['retornos_postivos'].rolling(window = 14).mean()
sol_dados['media_retornos_negativos'] = sol_dados['retornos_negativos'].rolling(window = 14).mean()
sol_dados = sol_dados.dropna()

# CALCULANDO RSI
btc_dados['RSI'] = (100 - 100/ (1 + btc_dados['media_retornos_positivos']/btc_dados['media_retornos_negativos']))

eth_dados['RSI'] = (100 - 100/ (1 + eth_dados['media_retornos_positivos']/eth_dados['media_retornos_negativos']))

sol_dados['RSI'] = (100 - 100/ (1 + sol_dados['media_retornos_positivos']/sol_dados['media_retornos_negativos']))

# CALCULANDO MÉDIA 14 SEMANAL
btc_dados['EMA_14'] = btc_dados['Close'].ewm(span=98, adjust=False, min_periods=0).mean()
eth_dados['EMA_14'] = eth_dados['Close'].ewm(span=98, adjust=False, min_periods=0).mean()
sol_dados['EMA_14'] = sol_dados['Close'].ewm(span=98, adjust=False, min_periods=0).mean()

############################## CALCULANDO SITUAÇÕES DE COMPRA ##############################

# BITCOIN
btc_dados.loc[btc_dados['Close'] > btc_dados['EMA_14'], 'compra'] = 'sim' # Se a cotação for maior que a EMA 8, COMPRA!!
btc_dados.loc[btc_dados['RSI'] >= 55, 'compra'] = 'sim' # Se o RSI for maior ou igual a 55, COMPRA!!
btc_dados.loc[btc_dados['RSI'] < 55, 'compra'] = 'nao' # Se o RSI for menor que 55, NÃO COMPRA!!
btc_dados.loc[btc_dados['Close'] < btc_dados['EMA_14'], 'compra'] = 'nao' # Se a cotação for menor que o EMA 8, NÃO COMPRA!!
btc_dados.loc[btc_dados['RSI'] > 70, 'compra'] = 'nao' # Se o RSI for maior que 70, NÃO COMPRA!!

# ETHEREUM
eth_dados.loc[eth_dados['Close'] > eth_dados['EMA_14'], 'compra'] = 'sim' # Se a cotação for maior que a EMA 8, COMPRA!!
eth_dados.loc[eth_dados['RSI'] >= 55, 'compra'] = 'sim' # Se o RSI for maior ou igual a 55, COMPRA!!
eth_dados.loc[eth_dados['RSI'] < 55, 'compra'] = 'nao' # Se o RSI for menor que 55, NÃO COMPRA!!
eth_dados.loc[eth_dados['Close'] < eth_dados['EMA_14'], 'compra'] = 'nao' # Se a cotação for menor que o EMA 8, NÃO COMPRA!!
eth_dados.loc[eth_dados['RSI'] > 70, 'compra'] = 'nao' # Se o RSI for maior que 70, NÃO COMPRA!!

# SOLANA
sol_dados.loc[sol_dados['Close'] > sol_dados['EMA_14'], 'compra'] = 'sim' # Se a cotação for maior que a EMA 8, COMPRA!!
sol_dados.loc[sol_dados['RSI'] >= 55, 'compra'] = 'sim' # Se o RSI for maior ou igual a 55, COMPRA!!
sol_dados.loc[sol_dados['RSI'] < 55, 'compra'] = 'nao' # Se o RSI for menor que 55, NÃO COMPRA!!
sol_dados.loc[sol_dados['Close'] < sol_dados['EMA_14'], 'compra'] = 'nao' # Se a cotação for menor que o EMA 8, NÃO COMPRA!!
sol_dados.loc[sol_dados['RSI'] > 70, 'compra'] = 'nao' # Se o RSI for maior que 70, NÃO COMPRA!!

# CALCULANDO SITUAÇÕES DE COMPRA & VENDA


############################## BITCOIN ##############################
ordem_aberta = 0
data_compra_btc = []
data_compra_amanha_btc = ''
data_venda_btc = []
data_venda_amanha_btc = ''

from datetime import datetime, timedelta
for i in range(len(btc_dados)):
    ####### COMPRAS #######
    if "sim" in btc_dados['compra'].iloc[i]:
        if ordem_aberta == 0:
          try:
            data_compra_btc.append(btc_dados.iloc[i+1].name) # +1 porque a gente compra no preço de abertura do dia seguinte
          except:
            data_amanha_btc = datetime.now() + timedelta(days=1)
            data_compra_amanha_btc = data_amanha_btc

        ordem_aberta = 1

    ####### VENDAS #######
    try:
      if btc_dados['RSI'].iloc[i] < 40 and btc_dados['Close'].iloc[i] < btc_dados['EMA_14'].iloc[i]: #Vender se o RSI for menor que 40 e a cotação estiver menor que o EMA
        if ordem_aberta == 1:
          try:
            data_venda_btc.append(btc_dados.iloc[i + 1].name) #vende no dia seguinte q bater 40
            ordem_aberta = 0
          except:
            data_amanha_btc = datetime.now() + timedelta(days=1)
            data_venda_amanha_btc = data_amanha_btc
            break

    except:
      continue




############################## ETHEREUM ##############################
ordem_aberta = 0
data_compra_eth = []
data_compra_amanha_eth = ''
data_venda_eth = []
data_venda_amanha_eth = ''

from datetime import datetime, timedelta
for i in range(len(eth_dados)):
    ####### COMPRAS #######
    if "sim" in eth_dados['compra'].iloc[i]:
        if ordem_aberta == 0:
          try:
            data_compra_eth.append(eth_dados.iloc[i+1].name) # +1 porque a gente compra no preço de abertura do dia seguinte
          except:
            data_amanha_eth = datetime.now() + timedelta(days=1)
            data_compra_amanha_eth = data_amanha_eth

        ordem_aberta = 1

    ####### VENDAS #######
    try:
      if eth_dados['RSI'].iloc[i] < 40 and eth_dados['Close'].iloc[i] < eth_dados['EMA_14'].iloc[i]: #Vender se o RSI for menor que 40 e a cotação estiver menor que o EMA
        if ordem_aberta == 1:
          try:
            data_venda_eth.append(eth_dados.iloc[i + 1].name) #vende no dia seguinte q bater 40
            ordem_aberta = 0
          except:
            data_amanha_eth = datetime.now() + timedelta(days=1)
            data_venda_amanha_eth = data_amanha_eth
            break

    except:
      continue



############################## SOLANA ##############################
ordem_aberta = 0
data_compra_sol = []
data_compra_amanha_sol = ''
data_venda_sol = []
data_venda_amanha_sol = ''

from datetime import datetime, timedelta
for i in range(len(sol_dados)):
    ####### COMPRAS #######
    if "sim" in sol_dados['compra'].iloc[i]:
        if ordem_aberta == 0:
          try:
            data_compra_sol.append(sol_dados.iloc[i+1].name) # +1 porque a gente compra no preço de abertura do dia seguinte
          except:
            data_amanha_sol = datetime.now() + timedelta(days=1)
            data_compra_amanha_sol = data_amanha_sol

        ordem_aberta = 1

    ####### VENDAS #######
    try:
      if sol_dados['RSI'].iloc[i] < 40 and sol_dados['Close'].iloc[i] < sol_dados['EMA_14'].iloc[i]: #Vender se o RSI for menor que 40 e a cotação estiver menor que o EMA
        if ordem_aberta == 1:
          try:
            data_venda_sol.append(sol_dados.iloc[i + 1].name) #vende no dia seguinte q bater 40
            ordem_aberta = 0
          except:
            data_amanha_sol = datetime.now() + timedelta(days=1)
            data_venda_amanha_sol = data_amanha_sol
            break

    except:
      continue

cotacao_dia_btc = btc_dados.iloc[-1]
retorno_btc = cotacao_dia_btc["retornos"]

cotacao_dia_eth = eth_dados.iloc[-1]
retorno_eth = cotacao_dia_eth["retornos"]

cotacao_dia_sol = sol_dados.iloc[-1]
retorno_sol = cotacao_dia_sol["retornos"]


outlook = win32.Dispatch("outlook.application")

email = outlook.CreateItem(0)



email.To = "bruninho_123vini@hotmail.com"
email.Subject = "Relatório de Mercado"
email.Body = f"""Segue o Relatório das Criptos:

* BITCOIN
- Cotação do dia: {cotacao_dia_btc['Close']:.2f}
- Variação do Dia: {retorno_btc*100:.2f}%
- RSI: {cotacao_dia_btc['RSI']:.0f}
- Média Móvel: {cotacao_dia_btc['EMA_14']:.0f}
- Compra: {cotacao_dia_btc['compra']}
- Última Compra: {data_compra_btc[-1]}
- Última Venda: {data_venda_btc[-1]}
- Compra Amanha: {data_compra_amanha_btc}
- Venda Amanha: {data_venda_amanha_btc}

* ETHEREUM
- Cotação do Dia: {cotacao_dia_eth['Close']:.2f}
- Variação do Dia: {retorno_eth*100:.2f}%
- RSI: {cotacao_dia_eth['RSI']:.0f}
- Média Móvel: {cotacao_dia_eth['EMA_14']:.0f}
- Compra: {cotacao_dia_eth['compra']}
- Última Compra: {data_compra_eth[-1]}
- Última Venda: {data_venda_eth[-1]}
- Compra Amanha: {data_compra_amanha_eth}
- Venda Amanha: {data_venda_amanha_eth}

* SOLANA
- Cotação do Dia: {cotacao_dia_sol['Close']:.2f}
- Variação do Dia: {retorno_sol*100:.2f}%
- RSI: {cotacao_dia_sol['RSI']:.0f}
- Média Móvel: {cotacao_dia_sol['EMA_14']:.0f}
- Compra: {cotacao_dia_sol['compra']}
- Última Compra: {data_compra_sol[-1]}
- Última Venda: {data_venda_sol[-1]}
- Compra Amanha: {data_compra_amanha_sol}
- Venda Amanha: {data_venda_amanha_sol}

Segue em anexo o gráfico dos ativos nos últimos 6 meses."""

########################## PLOTANDO GRÁFICOS ################################
# BITCOIN
plt.figure(figsize = (12, 5))
plt.title("BITCOIN")
plt.scatter(btc_dados.loc[data_compra_btc].index, btc_dados.loc[data_compra_btc]['Adj Close'], marker = '^',
            c = 'g')
plt.scatter(btc_dados.loc[data_venda_btc].index, btc_dados.loc[data_venda_btc]['Adj Close'], marker = '^',
            c = 'r')
plt.plot(btc_dados['Adj Close'], alpha = 0.7)

plt.savefig("bitcoin.png")

# ETHEREUM
plt.figure(figsize = (12, 5))
plt.title("ETHEREUM")
plt.scatter(eth_dados.loc[data_compra_eth].index, eth_dados.loc[data_compra_eth]['Adj Close'], marker = '^',
            c = 'g')
plt.scatter(eth_dados.loc[data_venda_eth].index, eth_dados.loc[data_venda_eth]['Adj Close'], marker = '^',
            c = 'r')
plt.plot(eth_dados['Adj Close'], alpha = 0.7)

plt.savefig("ethereum.png")

# SOLANA
plt.figure(figsize = (12, 5))
plt.title("SOLANA")
plt.scatter(sol_dados.loc[data_compra_sol].index, sol_dados.loc[data_compra_sol]['Adj Close'], marker = '^',
            c = 'g')
plt.scatter(sol_dados.loc[data_venda_sol].index, sol_dados.loc[data_venda_sol]['Adj Close'], marker = '^',
            c = 'r')
plt.plot(sol_dados['Adj Close'], alpha = 0.7)

plt.savefig("solana.png")

############################################################################

anexo_bitcoin = r"C:\Users\brsan\Downloads\bitcoin.png"
anexo_ethereum = r"C:\Users\brsan\Downloads\ethereum.png"
anexo_solana = r"C:\Users\brsan\Downloads\solana.png"

email.Attachments.Add(anexo_bitcoin)
email.Attachments.Add(anexo_ethereum)
email.Attachments.Add(anexo_solana)

email.Send()




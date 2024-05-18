import pandas as pd
import os
import datetime

# Obtém a data e hora actual
data = datetime.datetime.now()

# Define as colunas do DataFrame consolidado
colunas = [
    'Tipo',
    'Cidade',
    'Produto',
    'Qtde de Unidades Vendidas',
    'Preço Unitário',
    'Valor Total',
    'Desconto',
    'Valor Total c/ Desconto',
    'Custo Total',
    'Lucro',
    'Data',
    'Mês',
    'Ano'
]

# Cria um DataFrame vazio com as colunas definidas acima
consolidado = pd.DataFrame(columns=colunas)

# Lista todos os arquivos no diretório "planilhas"
arquivos = os.listdir("planilhas")

# Itera sobre cada arquivo na lista de arquivos
for arquivo in arquivos:
    # Divide o nome do arquivo pelo caractere '-' para extrair tipo e cidade
    dados_arquivos = arquivo.split('-')
    tipo = dados_arquivos[0]
    cidade = dados_arquivos[1].replace('.xlsx', '')

    # Lê o arquivo Excel
    df = pd.read_excel(f'planilhas\\{arquivo}')
    # Insere colunas "Tipo" e "Cidade" no DataFrame
    df.insert(0, 'Tipo', tipo)
    df.insert(1, 'Cidade', cidade)

    # Concatena o DataFrame lido com o DataFrame consolidado
    consolidado = pd.concat([consolidado, df])

# Salva o DataFrame consolidado em um arquivo Excel, nomeado com o mês e ano actuais
consolidado.to_excel(f"Relatorio_{data.strftime('%m-%Y')}.xlsx", 
                    index=False, 
                    sheet_name="Relatorio_vendas")

import pandas as pd

# 1. Lendo dados de uma planilha Excel
# O Pandas lê o arquivo e o transforma em um DataFrame.
df = pd.read_excel('Vendas.xlsx')

# 2. Visualizando os dados
print("### Dados Originais ###")
print(df)
print("\n" + "="*30 + "\n")

# 3. Selecionando e Filtrando Dados
# Selecionar uma coluna específica (retorna uma Series)
vendedores = df['Vendedor']
# print("### Coluna de Vendedores ###")
# print(vendedores)

# Filtrar linhas com base em uma condição
# Ex: Vendas realizadas pela Ana
vendas_ana = df[df['Vendedor'] == 'Ana']
print("### Apenas Vendas da Ana ###")
print(vendas_ana)
print("\n" + "="*30 + "\n")

# 4. Manipulando Dados - Criando uma nova coluna
# Criar uma coluna 'Total Venda' multiplicando Quantidade pelo Preço Unitário
df['Total Venda'] = df['Quantidade'] * df['Preço Unitário']
print("### DataFrame com a nova coluna 'Total Venda' ###")
print(df)
print("\n" + "="*30 + "\n")

# 5. Agrupando Dados
# Calcular o total vendido por cada vendedor
total_por_vendedor = df.groupby('Vendedor')['Total Venda'].sum()
print("### Total de Vendas por Vendedor ###")
print(total_por_vendedor)
print("\n" + "="*30 + "\n")

# 6. Salvando o resultado em uma nova planilha
# O 'index=False' evita que o índice do DataFrame seja salvo como uma coluna no Excel.
df.to_excel('Vendas_Com_Total.xlsx', sheet_name='Relatorio', index=False)

print("Arquivo 'Vendas_Com_Total.xlsx' salvo com sucesso!")
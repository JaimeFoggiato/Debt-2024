import pandas as pd
import math

# Leia a planilha do Excel
df = pd.read_excel('deals-7428260-2348.xlsx', usecols=['Negócio - Título', 'Negócio - Pasta', 'Pessoa - Telefone'])

df[['Negócio ', 'Título']] = df['Negócio - Título'].str.split('- ', n=1, expand=True)

# Remova hífens, parênteses e espaços dos números de telefone
df['Pessoa - Telefone'] = df['Pessoa - Telefone'].str.replace('-', '').str.replace('(', '').str.replace(')', '').str.replace(' ', '')

# Separe os números de telefone que estão separados por vírgula e torne cada um deles uma linha
df['Pessoa - Telefone'] = df['Pessoa - Telefone'].str.split(',')
df = df.explode('Pessoa - Telefone')

# Adicione '55' na frente dos números de telefone
df['Pessoa - Telefone'] = '55' + df['Pessoa - Telefone']

# Edite a coluna 'Título' para deixar somente a primeira letra maiúscula e manter o primeiro e último nome
df['Título'] = df['Título'].str.title()
#df['Título'] = df['Título'].str.split(' ')
#df['Título'] = df['Título'].str[0] + ' ' + df['Título'].str[-1]


# Use a função melt para transformar as colunas de telefone em linhas
df = df.melt(id_vars=['Título', 'Negócio - Pasta', 'Pessoa - Telefone'], var_name='Telefone', value_name='Número')

# Exlcui as linhas em branco
df = df.dropna(subset=['Título', 'Negócio - Pasta', 'Pessoa - Telefone'])

df = df.drop(['Telefone', 'Número'], axis=1)

df = df.drop_duplicates()

# Determine the number of rows in the DataFrame
num_rows = len(df)


# Salve o resultado em um novo arquivo ou faça o que precisar com os dados
df.to_excel('Planilha Discador2.xlsx', index=False)

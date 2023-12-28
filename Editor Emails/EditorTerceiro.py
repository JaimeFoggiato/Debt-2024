import pandas as pd
import math

# Leia a planilha do Excel
df = pd.read_excel('deals-7428260-2375.xlsx', usecols=['Negócio - Pasta', 'Negócio - ID', 'Pessoa - E-mail', 'Negócio - Título', 'Negócio - Data do Acidente', 'Negócio - Título', 'Negócio - Proprietário', 'Pessoa - CPF/CNPJ'])

df[['Negócio ', 'Título']] = df['Negócio - Título'].str.split('- ', n=1, expand=True)

# Edite a coluna 'Título' para deixar somente a primeira letra maiúscula e manter o primeiro e último nome
df['Título'] = df['Título'].str.title()
df['Título'] = df['Título'] #.str.split(' ')
df['Título'] = df['Título'] #.str[0] + ' ' + df['Título'].str[-1]

df['Nome completo'] = df['Título']


# Formate a coluna 'Negócio - Data do Acidente' para "dd/mm/aaaa"
df['Negócio - Data do Acidente'] = pd.to_datetime(df['Negócio - Data do Acidente'], format='%Y-%m-%d').dt.strftime('%d/%m/%Y')

# Substitua hífens por pontos na coluna 'Negócio - Data do Acidente'
df['Negócio - Data do Acidente'] = df['Negócio - Data do Acidente'].str.replace('/', '.')

df['Email CC'] = 'debt-5188cf' + df['Negócio - ID'].astype(str) + '@pipedrivemail.com'

df['Pessoa - E-mail'] = df['Pessoa - E-mail'].str.replace(',', ';')

# Use a função melt para transformar as colunas de telefone em linhas
df = df.melt(id_vars=['Negócio - Pasta', 'Negócio - ID', 'Pessoa - E-mail', 'Título', 'Negócio - Data do Acidente', 'Negócio - Proprietário', 'Pessoa - CPF/CNPJ', 'Email CC'], var_name='Telefone', value_name='Número')

# Exlcui as linhas em branco
df = df.dropna(subset=['Título', 'Negócio - Pasta', 'Negócio - ID', 'Pessoa - E-mail', 'Título', 'Negócio - Data do Acidente', 'Negócio - Proprietário', 'Pessoa - CPF/CNPJ', 'Email CC'])

df = df.drop(['Telefone', 'Número'], axis=1)

df = df.drop_duplicates()

# Determine the number of rows in the DataFrame
num_rows = len(df)

# Agrupe o DataFrame com base na coluna 'Título'
grupos = df.groupby('Negócio - Proprietário')

# Iterar sobre os grupos e salvar cada grupo em uma planilha Excel separada
for titulo, grupo in grupos:
    nome_arquivo = f'{titulo}_dados.xlsx'  # Nome do arquivo de saída
    grupo.to_excel(nome_arquivo, index=False)
    print(f'Planilha para {titulo} salva como {nome_arquivo}')
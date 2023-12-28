import pandas as pd

# Leia a planilha do Excel
df = pd.read_excel('deals-7428260-2353.xlsx', usecols=['Negócio - Título', 'Negócio - Data do Acidente', 'Negócio - Veículo Segurado', 'Pessoa - Telefone', 'Negócio - Pasta'])

# Remova hífens, parênteses e espaços dos números de telefone
df['Pessoa - Telefone'] = df['Pessoa - Telefone'].str.replace('-', '').str.replace('(', '').str.replace(')', '').str.replace(' ', '')

# Separe os números de telefone que estão separados por vírgula e torne cada um deles uma linha
df['Pessoa - Telefone'] = df['Pessoa - Telefone'].str.split(',')
df = df.explode('Pessoa - Telefone')

# Adicione '55' na frente dos números de telefone
df['Pessoa - Telefone'] = '55' + df['Pessoa - Telefone']

# Edite a coluna 'Negócio - Título' para deixar somente a primeira letra maiúscula e manter o primeiro e último nome
df['Negócio - Título'] = df['Negócio - Título'].str.title()

cortar_string = 28
df['Negócio - Veículo Segurado'] = df['Negócio - Veículo Segurado'].str[:-cortar_string]

# Formate a coluna 'Negócio - Data do Acidente' para "dd/mm/aaaa"
df['Negócio - Data do Acidente'] = pd.to_datetime(df['Negócio - Data do Acidente'], format='%Y-%m-%d').dt.strftime('%d/%m/%Y')

# Substitua hífens por pontos na coluna 'Negócio - Data do Acidente'
df['Negócio - Data do Acidente'] = df['Negócio - Data do Acidente'].str.replace('/', '.')

# Use a função melt para transformar as colunas de telefone em linhas
df = df.melt(id_vars=['Negócio - Título', 'Negócio - Data do Acidente', 'Negócio - Veículo Segurado', 'Pessoa - Telefone'])

# Exclui as linhas em branco
df = df.dropna(subset=['Negócio - Título', 'Negócio - Data do Acidente', 'Negócio - Veículo Segurado', 'Pessoa - Telefone'])

#df = df.drop(['Telefone','Número'], axis=1)

df = df.drop_duplicates()

# Determine the number of rows in the DataFrame
num_rows = len(df)

# Salve o resultado em um novo arquivo ou faça o que precisar com os dados
df.to_excel('JetSender Seguradoteste.xlsx', index=False)

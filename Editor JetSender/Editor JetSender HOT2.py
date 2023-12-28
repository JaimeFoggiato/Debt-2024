import pandas as pd
import math

# Leia a planilha do Excel
df = pd.read_excel('PlanilhaJetSender.xlsx', usecols=['Negócio - Título', 'Negócio - Data do Acidente', 'Negócio - Veículo 3º', 'Pessoa - Telefone HOT 2'])

df[['Negócio ', 'Título']] = df['Negócio - Título'].str.split('- ', n=1, expand=True)

# Remova hífens, parênteses e espaços dos números de telefone
df['Pessoa - Telefone HOT 2'] = df['Pessoa - Telefone HOT 2'].str.replace('-', '').str.replace('(', '').str.replace(')', '').str.replace(' ', '')

# Separe os números de telefone que estão separados por vírgula e torne cada um deles uma linha
df['Pessoa - Telefone HOT 2'] = df['Pessoa - Telefone HOT 2'].str.split(',')
df = df.explode('Pessoa - Telefone HOT 2')

# Adicione '55' na frente dos números de telefone
df['Pessoa - Telefone HOT 2'] = '55' + df['Pessoa - Telefone HOT 2']

df['Pessoa - Telefone'] = df['Pessoa - Telefone'].str[:13]

# Exclua as linhas com números de telefone com menos de 13 caracteres / funciona como um if. Se tem a condição, continua, se não, ele cancela, não precisa estar exposto o else 
df = df[df['Pessoa - Telefone'].str.len() == 13]

# Edite a coluna 'Título' para deixar somente a primeira letra maiúscula e manter o primeiro e último nome
df['Título'] = df['Título'].str.title()
df['Título'] = df['Título'].str.split(' ')
df['Título'] = df['Título'].str[0] + ' ' + df['Título'].str[-1]

#df['Pessoa - Telefone'] = df['Pessoa - Telefone'].str.strip()   //Por enquanto posso tirar, a função acima já supre o que eu quero

# Formate a coluna 'Negócio - Data do Acidente' para "dd/mm/aaaa"
df['Negócio - Data do Acidente'] = pd.to_datetime(df['Negócio - Data do Acidente'], format='%Y-%m-%d').dt.strftime('%d/%m/%Y')

# Substitua hífens por pontos na coluna 'Negócio - Data do Acidente'
df['Negócio - Data do Acidente'] = df['Negócio - Data do Acidente'].str.replace('/', '.')


# Use a função melt para transformar as colunas de telefone em linhas
df = df.melt(id_vars=['Título', 'Negócio - Data do Acidente', 'Negócio - Veículo 3º', 'Pessoa - Telefone HOT 2'], var_name='Telefone', value_name='Número')

# Exlcui as linhas em branco
df = df.dropna(subset=['Título', 'Negócio - Data do Acidente', 'Negócio - Veículo 3º', 'Pessoa - Telefone HOT 2'])

df = df.drop(['Telefone', 'Número'], axis=1)

df = df.drop_duplicates()

# Determine the number of rows in the DataFrame
num_rows = len(df)

# Salve o resultado em um novo arquivo ou faça o que precisar com os dados
df.to_excel('Resultado Pessoa - Telefone HOT 2.xlsx', index=False)
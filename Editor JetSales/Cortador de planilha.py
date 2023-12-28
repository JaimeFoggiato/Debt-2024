import pandas as pd

# Carregue sua planilha
df = pd.read_excel('Planilha JetGO.xlsx')

# Determine o número de linhas por arquivo
linhas_por_arquivo = 5000

# Calcule o número total de arquivos necessários
num_arquivos = len(df) // linhas_por_arquivo + 1

# Divida o DataFrame em partes
partes = [df.iloc[i*linhas_por_arquivo : (i+1)*linhas_por_arquivo] for i in range(num_arquivos)]

# Salve cada parte como um arquivo Excel separado
for i, parte in enumerate(partes):
    parte.to_excel(f'Resultado Jetgo{i+1}.xlsx', index=False)

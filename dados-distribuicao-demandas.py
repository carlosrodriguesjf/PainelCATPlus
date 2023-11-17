import pandas as pd


def separa_data(data_protocolo):
    return data_protocolo.split(' ')[0]



# Especifique o caminho para o arquivo Excel
excel_file_path = r'E:\CentralDeAtendimento\Projetos\PainelCATPlus\arquivos_apoio\Distribuição de Demandas - Out2023.xlsx'

# Use a função read_excel para ler o arquivo Excel
xls = pd.ExcelFile(excel_file_path)

# Obtenha a lista de nomes de planilhas no arquivo
sheet_names = xls.sheet_names

# Itere pelas planilhas e carregue cada uma em um DataFrame do pandas
data = {}  # Dicionário para armazenar DataFrames de cada planilha

for sheet_name in sheet_names:
    data[sheet_name] = xls.parse(sheet_name)

# Agora você tem cada planilha em um DataFrame separado no dicionário 'data'
# Você pode acessar cada planilha pelo seu nome, por exemplo: data['NomeDaPlanilha']


df_aprov_estudos = data['Aprov. de estudos']
# Transforma a primeira linha em cabeçalho
df_aprov_estudos.columns = df_aprov_estudos.iloc[0]

# Remove a primeira linha (opcional)
df_aprov_estudos= df_aprov_estudos[1:]

df_aprov_estudos = df_aprov_estudos[['ATRIBUÍDO PARA','STATUS','SETOR','DATA PROTOCOLO']]

df_aprov_estudos.rename(columns={'ATRIBUÍDO PARA':'nome_operador','STATUS':'status','SETOR':'setor','DATA PROTOCOLO':'data_protocolo'},inplace=True)

df_aprov_estudos['data_protocolo'] = df_aprov_estudos['data_protocolo'].astype(str)

df_aprov_estudos.data_protocolo = df_aprov_estudos.data_protocolo.apply(separa_data)

df_aprov_estudos['data_protocolo'].unique()

df_aprov_estudos.dropna(inplace=True)




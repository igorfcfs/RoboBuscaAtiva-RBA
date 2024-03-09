import os
import pandas as pd
import subprocess

diretorio_base_usuario = os.path.expanduser("~")

# Pasta onde estão as planilhas
pasta_downloads = os.path.join(diretorio_base_usuario, "Downloads")
caminho_desempenho = 'BasesDeDados/DESEMPENHO POR ESTUDANTE.xlsx'

# Função para capitalizar a primeira letra de cada palavra
def capitalize_names(name):
    return ' '.join(word.capitalize() for word in name.split())

# Lista para armazenar os DataFrames
dfs = []

# Primeiro, processamos a planilha 'DESEMPENHO POR ESTUDANTE' para obter os RAs
if os.path.exists(caminho_desempenho):
    desempenho_df = pd.read_excel(caminho_desempenho)
    ra_dict = dict(zip(desempenho_df['ALUNO'], desempenho_df['RA']))
else:
    raise FileNotFoundError(f"O arquivo 'DESEMPENHO POR ESTUDANTE.xlsx' não foi encontrado em {caminho_desempenho}.")

# Percorre todos os arquivos na pasta de downloads
for arquivo in os.listdir(pasta_downloads):
    if arquivo.endswith('.xlsx') and not arquivo.startswith('~$') and arquivo != 'DESEMPENHO POR ESTUDANTE.xlsx':
        arquivo_path = os.path.join(pasta_downloads, arquivo)

        # Verifica o nome da coluna na primeira linha da planilha
        df = pd.read_excel(arquivo_path)
        if 'ALUNO' in df.columns:
            # Se a coluna se chama 'ALUNO', renomeia para 'Nome'
            df.rename(columns={'ALUNO': 'Nome'}, inplace=True)
        elif 'Nome' not in df.columns:
            raise ValueError(f"O arquivo '{arquivo}' não possui uma coluna válida.")
        
        # Filtra os alunos cuja primeira coluna tem o valor 1
        df = df[df.iloc[:, 0] == 1]

        # Capitaliza a primeira letra de cada palavra na coluna 'Nome'
        df['Nome'] = df['Nome'].apply(capitalize_names)

        # Adiciona a coluna 'RA' com base no mapeamento
        df['RA'] = df['Nome'].map(ra_dict)

        # Adiciona a coluna 'Série' do nome do arquivo
        df['Série'] = arquivo.split(' - ')[0]

        # Adiciona o DataFrame à lista
        dfs.append(df)

# Concatena todos os DataFrames em um único DataFrame
df_completo = pd.concat(dfs, ignore_index=True)

try:
    # Caminho para o arquivo de destino
    caminho_arquivo = 'ArquivosGerados/dados_filtrados_com_RA_e_Nome_e_Série.xlsx'

    # Extrair o diretório do caminho do arquivo
    diretorio_destino = os.path.dirname(caminho_arquivo)

    # Verificar se o diretório de destino não existe e, se não existir, criar-o
    if not os.path.exists(diretorio_destino):
        os.makedirs(diretorio_destino)

    # Agora você pode salvar o arquivo no diretório de destino
    df_completo.to_excel(caminho_arquivo, index=False)

    print(f"Arquivo '{caminho_arquivo}' foi gerado com sucesso.")
except Exception as e:
    print(f"Erro ao gerar o arquivo Excel: {str(e)}")


def executar_script_gerar_novo_modelo_convocacao():
    # Obter o diretório do arquivo em execução
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))

    # Nome do arquivo a ser executado (neste caso, na mesma pasta)
    caminho_segundo_script = "gerar_novo_modelo_convocacao.py"

    # Caminho completo para o segundo script
    caminho_completo = os.path.join(diretorio_atual, caminho_segundo_script)

    # Executar o segundo script
    subprocess.call(["python", caminho_completo])


executar_script_gerar_novo_modelo_convocacao()

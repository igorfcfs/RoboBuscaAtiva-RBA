from docxtpl import DocxTemplate, RichText
from pathlib import Path
import pandas as pd
import subprocess
import os

# Carregue a planilha 'dados_filtrados_agregados'
dados = pd.read_excel('ArquivosGerados/dados_filtrados_com_RA_e_Nome_e_Série.xlsx')

# Obtém o diretório do script
script_path = Path(__file__).resolve().parent

# Obtém o caminho absoluto para o arquivo do modelo
modelo_path = script_path.parent / "ArquivosGerados" / "modelo_convocacao.docx"

# Pasta para salvar os documentos gerados
output_folder = script_path.parent / "ArquivosGerados" / "documentos_individuais"
output_folder.mkdir(exist_ok=True)

# Itere sobre os dados da planilha e crie um documento para cada aluno
for _, linha in dados.iterrows():
    nome = linha['Nome']
    numero = linha['RA']
    serie = linha['Série']

    # Crie um novo documento baseado no modelo
    doc = DocxTemplate(modelo_path)

    # Crie o contexto com os dados específicos do aluno
    context = {
        'ALUNO': nome,
        'RA': numero,
        'SERIE': serie
    }

    # Renderize o modelo com o contexto
    doc.render(context)

    # Salve o novo documento com um nome único para cada aluno
    output_path = output_folder / f"documento_{nome}_{numero}_{serie}.docx"
    doc.save(output_path)

print("Documentos individuais gerados para cada aluno.")

# Função para executar o segundo script
def executar_script_converter_convocacao_para_pdf():
    # Obter o diretório do arquivo em execução
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))

    # Nome do arquivo a ser executado (neste caso, na mesma pasta)
    caminho_segundo_script = "converter_convocacao_para_pdf.py"

    # Caminho completo para o segundo script
    caminho_completo = os.path.join(diretorio_atual, caminho_segundo_script)

    # Executar o segundo script
    subprocess.call(["python", caminho_completo])

# Execute o segundo script
executar_script_converter_convocacao_para_pdf()

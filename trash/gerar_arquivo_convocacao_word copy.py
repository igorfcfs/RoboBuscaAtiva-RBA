from docxtpl import DocxTemplate, RichText
from pathlib import Path
import pandas as pd
import subprocess
import os

# Carregue a planilha 'dados_filtrados_agregados'
dados = pd.read_excel('ArquivosGerados/dados_filtrados_com_RA_e_Nome_e_Série.xlsx')

# Abra o arquivo Word de modelo (documento_base)
script_path = Path(__file__).resolve().parent
document_path = script_path.parent / "ArquivosGerados" / "modelo_convocacao.docx"
doc = DocxTemplate(document_path)

# Salve o novo documento Word
document_path_arquivos_gerados = script_path.parent / "ArquivosGerados" / "documento_convocacao.docx"

# Lista para armazenar os dados de contexto de cada aluno
contextos = []

# Itere sobre os dados da planilha e armazene os contextos em uma lista
for _, linha in dados.iterrows():
    nome = linha['Nome']
    numero = linha['RA']
    serie = linha['Série']

    context = {
        'ALUNO': nome,
        'RA': numero,
        'SERIE': serie 
    }

    contextos.append(context)

# Adicione uma quebra de página após cada aluno, exceto o último
for i, contexto in enumerate(contextos):
    doc.render(contexto)
    if i < len(contextos) - 1:
        doc.add_page_break()

# Salve o novo documento Word com as páginas duplicadas e as substituições
doc.save(document_path_arquivos_gerados)

print("Arquivo Word gerado com páginas separadas para cada aluno e as substituições de dados mantendo o estilo original.")

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

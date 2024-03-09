from docxtpl import DocxTemplate
from pathlib import Path
import pandas as pd
from docx import Document
import os
import subprocess
from docx.shared import Pt

# Carregue a planilha 'dados_filtrados_agregados'
dados = pd.read_excel('ArquivosGerados/dados_filtrados_com_RA_e_Nome_e_Série.xlsx')

# Obtém o diretório do script
script_path = Path(__file__).resolve().parent

# Obtém o caminho absoluto para o arquivo do modelo
modelo_path = script_path.parent / "ArquivosGerados" / "modelo_convocacao.docx"

# Caminho para salvar o documento gerado
output_path = script_path.parent / "ArquivosGerados" / "documento_convocacao.docx"

# Lista para armazenar os contextos de cada aluno
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

# Lista para armazenar os elementos de parágrafo e run de cada documento temporário
elementos_temporarios = []

# Crie um novo documento para cada aluno e adicione os elementos à lista
for i, contexto in enumerate(contextos):
    # Crie um novo objeto DocxTemplate
    template = DocxTemplate(modelo_path)

    # Renderize as variáveis no modelo
    template.render(contexto)

    # Adicione os elementos do documento temporário à lista
    for element in template.element.body:
        elementos_temporarios.append(element)

    # Adicione uma quebra de página após cada aluno, exceto o último
    if i < len(contextos) - 1:
        # Adicione uma quebra de página diretamente
        quebra_pagina = Document()
        quebra_pagina.add_page_break()
        elementos_temporarios.append(quebra_pagina.element.body)

# Crie um novo documento final e adicione os elementos da lista
documento_final = Document()
for element in elementos_temporarios:
    documento_final.element.body.append(element)

# Salve o documento final
documento_final.save(output_path)

print("Arquivo Word gerado com uma página para cada aluno e as substituições de dados mantendo o estilo original.")

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
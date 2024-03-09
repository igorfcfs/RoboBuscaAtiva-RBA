from docxtpl import DocxTemplate, RichText
from pathlib import Path
import pandas as pd
import copy
import subprocess
import os

# Carregue a planilha 'dados_filtrados_agregados'
dados = pd.read_excel('ArquivosGerados/dados_filtrados_com_RA_e_Nome_e_Série.xlsx')

# Abra o arquivo Word de modelo (documento_base)

# Obtém o diretório do script
script_path = Path(__file__).resolve().parent

# Obtém o caminho absoluto para o arquivo
document_path = script_path.parent / "ArquivosGerados" / "modelo_convocacao.docx"

doc = DocxTemplate(document_path)

# Salve o novo documento Word
document_path_arquivos_gerados = script_path.parent / "ArquivosGerados" / "documento_convocacao.docx"

# Variável para rastrear o título anterior
titulo_anterior = None

# Itere sobre os dados da planilha
for i, linha in dados.iterrows():
    nome = linha['Nome']
    numero = linha['RA']
    serie = linha['Série']

    context = {
        'ALUNO': nome,
        'RA': numero,
        'SERIE': serie 
    }

    # Adicione uma quebra de página após cada aluno, exceto o último
    if i < len(dados) - 1:
        doc.render(context)
        doc.paragraphs[-1].runs[0].add_break()

# Função para adicionar quebras de página antes de cada ocorrência de "GOVERNO DO ESTADO DE SÃO PAULO"
def adicionar_quebras_de_pagina(doc):
    primeira_pagina = True  # Flag para controlar a primeira página
    for par in doc.paragraphs:
        if "GOVERNO DO ESTADO DE SÃO PAULO" in par.text:
            if primeira_pagina:
                primeira_pagina = False  # Desativa a flag na primeira ocorrência
            else:
                par.clear()
                run = par.add_run()
                run.add_break()

def remover_primeiro_paragrafo(doc):
    # Verifique se o documento tem parágrafos
    if len(doc.paragraphs) > 0:
        # Se houver parágrafos, exclua o primeiro
        primeiro_paragrafo = doc.paragraphs[0]
        doc.element.body.remove(primeiro_paragrafo._element)

# Salve o novo documento Word com as páginas duplicadas e as substituições
adicionar_quebras_de_pagina(doc)
remover_primeiro_paragrafo(doc)
doc.save(Path(__file__).parent / document_path_arquivos_gerados)

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

import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from docx import Document
from docx2pdf import convert
import os
import shutil

# Carregue a planilha 'RelatorioBuscaAtiva'
dados = pd.read_excel('ArquivosGerados/RelatorioBuscaAtiva.xlsx')

script_dir = Path(__file__).resolve().parent 

dir_modelo_convocacao = script_dir.parent / "ArquivosGerados" / "modelo_convocacao.docx"

doc = DocxTemplate(dir_modelo_convocacao)
alunos = []

# Crie uma pasta para armazenar os documentos individuais dos alunos
pasta_convocacoes_indiv = script_dir.parent / "ArquivosGerados" / "Documentos Individuais"
pasta_convocacoes_indiv.mkdir(exist_ok=True)  # Cria a pasta se ainda não existir

for _,linha in dados.iterrows():
    nome = linha['Nome']
    numero = linha['RA']
    serie = linha["Série"]

    aluno = {
        'ALUNO': nome.upper(),
        # 'RA': numero.upper(),
        'SERIE': serie.upper()
    }

    # Selecionar uma chave específica do dicionário (por exemplo, 'ALUNO') para compor o nome do arquivo
    nome_arquivo = f"documento_convocacao_{aluno['ALUNO']}.docx"
    
    # Renderizar e salvar o documento com o nome do aluno na pasta específica
    doc.render(aluno)
    caminho_arquivo = pasta_convocacoes_indiv / nome_arquivo
    doc.save(caminho_arquivo)

    alunos.append(aluno)

def unir_convocacoes():
    diretorio = Path(__file__).parent.parent / "ArquivosGerados/Documentos Individuais"
    arquivos = os.listdir(diretorio)
    arquivos_listados = [arquivo for arquivo in arquivos if arquivo.endswith('.docx')]

    doc_inicial = Path(diretorio / f'{arquivos_listados[0]}')
    doc_todas_conv = diretorio.parent / "ConvocacaoParaCompensarFaltas.docx"
    shutil.copyfile(doc_inicial, doc_todas_conv)

    for arquivo in arquivos_listados[1:]:
        arquivo_temp = Path(diretorio) / f'{arquivo}'
        master = Document(doc_todas_conv)
        master.add_page_break()
        composer = Composer(master)
        doc = Document(arquivo_temp)
        composer.append(doc)
        composer.save(doc_todas_conv)

unir_convocacoes()


def converter_para_pdf():
    convert(Path(__file__).parent.parent / "ArquivosGerados/ConvocacaoParaCompensarFaltas.docx", Path(__file__).parent.parent / "ArquivosGerados")
converter_para_pdf()

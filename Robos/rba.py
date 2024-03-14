import os
import subprocess
from docxtpl import DocxTemplate
from pathlib import Path
from datetime import datetime

diretorio_padrao = "../Planilias"

def alterar_diretorio_padrao():
    pass

def executar_script_renomeia_tabelas():
    diretorio_atual = Path(__file__).parent / "renomeia_tabelas.py"
    subprocess.call(["python", diretorio_atual])

def executar_script_gerar_planilia_alunos_convocados():
    diretorio_atual = Path(__file__).parent / "gerar_planilia_alunos_convocados.py"
    subprocess.call(["python", diretorio_atual])

def executar_script_gerar_arquivos_convocacao():
    diretorio_atual = Path(__file__).parent / "gerar_arquivos_convocacao.py"
    subprocess.call(["python", diretorio_atual])

def enviar_email():
    diretorio_atual = Path(__file__).parent / "enviar_email.py"
    subprocess.call(["python", diretorio_atual])

def executar_tudo():
    executar_script_renomeia_tabelas()
    executar_script_gerar_planilia_alunos_convocados()
    criar_modelo_convocacao()
    executar_script_gerar_arquivos_convocacao()
    enviar_email()

def criar_modelo_convocacao():
    ano_atual = datetime.today()

    def substituir_documento():
        # Obtém o caminho absoluto para o arquivo
        document_path_base_dados = Path(__file__).parent.parent / "BasesDeDados" / "CONVOCACAO_PARA_COMPENSAR_FALTAS.docx"

        doc = DocxTemplate(document_path_base_dados)
        
        with open('./BasesDeDados/dados.txt', 'r') as arquivo:
            linhas = arquivo.readlines()
            valores = [linha.strip() for linha in linhas]

        # Dicionário de padrões a serem substituídos
        padroes = {
            'REGIAO': valores[0].upper(),
            'DEPARTAMENTO': valores[1].upper(),
            'RUA': valores[2].capitalize(),
            'NUMERO': valores[3].capitalize(),
            'BAIRRO': valores[4].capitalize(),
            'CIDADE': valores[5].capitalize(),
            'ESTADO': valores[6].capitalize(),
            'CEP': valores[7],
            'TELEFONE': valores[8],
            'EMAIL': valores[9],
            'ANO': ano_atual.strftime("%Y"),
            "ALUNO": "{{ALUNO}}",
            "RA": "{{RA}}",
            "SERIE": "{{SERIE}}"
        }
        
        doc.render(padroes)

        # Salve o novo documento Word
        document_path_arquivos_gerados = Path(__file__).parent.parent / "ArquivosGerados" / "modelo_convocacao.docx"
        doc.save(Path(__file__).parent / document_path_arquivos_gerados)

        print(f'Documento gerado e salvo em {document_path_arquivos_gerados}')

    substituir_documento()

executar_tudo()
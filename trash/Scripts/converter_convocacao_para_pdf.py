import docx
from docx2pdf import convert
import subprocess
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, PageBreak, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO
import os

# Caminho para o arquivo Word recuperado
caminho_arquivo_word = 'ArquivosGerados/documento_convocacao.docx'

convert('ArquivosGerados/documento_convocacao.docx')

os.rename('ArquivosGerados/documento_convocacao.pdf', 'ArquivosGerados/ConvocacaoParaCompensarFaltas.pdf')

print("Arquivo PDF gerado com quebras de página, exceto na primeira página.")

def executar_script_enviar_email():
    # Obter o diretório do arquivo em execução
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))

    # Nome do arquivo a ser executado (neste caso, na mesma pasta)
    caminho_segundo_script = "enviar_email.py"

    # Caminho completo para o segundo script
    caminho_completo = os.path.join(diretorio_atual, caminho_segundo_script)

    # Executar o segundo script
    subprocess.call(["python", caminho_completo])

executar_script_enviar_email()
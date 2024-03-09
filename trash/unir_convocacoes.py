from docxcompose.composer import Composer
from docx import Document
from pathlib import Path
import os
import shutil
from docx import *

diretorio = Path(__file__).parent.parent / "ArquivosGerados/Documentos Individuais"
arquivos = os.listdir(diretorio)
arquivos_listados = [arquivo for arquivo in arquivos if arquivo.endswith('.docx')]

doc_inicial = Path(diretorio / f'{arquivos_listados[0]}')
doc_todas_conv = diretorio.parent / "Todas_convocações.docx"
shutil.copyfile(doc_inicial, doc_todas_conv)

for arquivo in arquivos_listados[1:]:
    arquivo_temp = Path(diretorio) / f'{arquivo}'
    master = Document(doc_todas_conv)
    master.add_page_break()
    composer = Composer(master)
    doc = Document(arquivo_temp)
    composer.append(doc)
    composer.save(doc_todas_conv)
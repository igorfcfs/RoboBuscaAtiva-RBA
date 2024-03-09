from pathlib import Path
from docxtpl import DocxTemplate
import datetime

ano_atual = datetime.datetime.today()

#Pedir formulário
regiao_escola = "SUL"
departamento_escola = "SAS"
rua_endereco = "ASASsss"
numero_endereco_escola = "ASSss"
jardim_endereco_escola = "Sss"
cidade_endereco_escola = "sssaa"
estado_endereco_escola = "skkok"
cep_endereco_escola = ""
telefone_escola = ""
email_escola = ""

#Informações do excel
nome_aluno = ""
ra_aluno = ""
serie_aluno = ""

#Abre o modelo de convocação word
modelo = Path(__file__).parent / "BasesDeDados/CONVOCACAO_PARA_COMPENSAR_FALTAS.docx"
doc = DocxTemplate(modelo)

#Aqui onde são substituidas os caracteres
context = {
    "REGIAO": regiao_escola,
    "DEPARTAMENTO": departamento_escola,
    "RUA": rua_endereco,
    "NUMERO": numero_endereco_escola,
    "JARDIM": jardim_endereco_escola,
    "CIDADE": cidade_endereco_escola,
    "ESTADO": estado_endereco_escola,
    "CEP": cep_endereco_escola,
    "TELEFONE": telefone_escola,
    "EMAIL": email_escola,
    "ANO": ano_atual.strftime("%Y")
}
#Substitui as palavras no documento
doc.render(context)
doc.save(Path(__file__).parent / "modelo_convocacao.docx")
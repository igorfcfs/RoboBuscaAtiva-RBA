import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

nome = "Diretor"

# Configurações de email
de_email = 'igorfcfs@gmail.com'
para_email = 'igorfcfs@gmail.com'
senha = 'lwayqzokxlfdatwf'

# Crie o objeto MIMEMultipart
msg = MIMEMultipart()
msg['From'] = de_email
msg['To'] = para_email
msg['Subject'] = 'Relatório Busca Ativa e Arquivos de Convocação gerados pelo RBA dos Alunos Faltosos'

# Corpo do email (texto HTML)
corpo_email = f"""
<p>Olá {nome},</p>
<p>Segue aqui o Relatório da Busca Ativa e o Arquivo de Convocação dos Alunos que Faltaram.</p>
"""

msg.attach(MIMEText(corpo_email, 'html'))

# Anexos
caminho_arquivo_relatorio = 'ArquivosGerados/RelatorioBuscaAtiva.xlsx'
caminho_arquivo_convocacao = 'ArquivosGerados/ConvocacaoParaCompensarFaltas.pdf'

# Anexo Relatório
anexo_relatorio = MIMEBase('application', 'octet-stream')
anexo_relatorio.set_payload(open(caminho_arquivo_relatorio, 'rb').read())
encoders.encode_base64(anexo_relatorio)
anexo_relatorio.add_header('Content-Disposition', f'attachment; filename={os.path.basename(caminho_arquivo_relatorio)}')
msg.attach(anexo_relatorio)

# Anexo Convocação
anexo_convocacao = MIMEBase('application', 'octet-stream')
anexo_convocacao.set_payload(open(caminho_arquivo_convocacao, 'rb').read())
encoders.encode_base64(anexo_convocacao)
anexo_convocacao.add_header('Content-Disposition', f'attachment; filename={os.path.basename(caminho_arquivo_convocacao)}')
msg.attach(anexo_convocacao)

# Conecte-se ao servidor SMTP (neste caso, Gmail)
servidor_smtp = smtplib.SMTP('smtp.gmail.com', 587)
servidor_smtp.starttls()

# Faça login na sua conta de email
servidor_smtp.login(de_email, senha)

# Envie o email
servidor_smtp.sendmail(de_email, para_email, msg.as_string())

# Encerre a conexão
servidor_smtp.quit()

print('Email com anexos enviados com sucesso.')

import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

nome = "Diretor" # Por hora o nome sera diretor, porem, depois o nome sera um selec da coluna NomeDoDiretor do nosso banco de dados

# Configurações de email
de_email = 'igorfcfs@gmail.com'
para_email = 'igorfcfs@gmail.com'
senha = 'lwayqzokxlfdatwf'

# Crie o objeto MIMEMultipart
msg = MIMEMultipart()
msg['From'] = de_email
msg['To'] = para_email
msg['Subject'] = 'Arquivos de Convocação gerados pelo RBA dos Alunos Faltosos'

# Corpo do email (texto HTML)
corpo_email = f"""
<p>Olá {nome},</p>
<p>Segue aqui o arquivo da convocação dos alunos que faltaram.</p>
"""

msg.attach(MIMEText(corpo_email, 'html'))

# Anexo
caminho_arquivo = 'ArquivosGerados/ConvocacaoParaCompensarFaltas.pdf'

# Crie o objeto do anexo
anexo = MIMEBase('application', 'octet-stream')
anexo.set_payload(open(caminho_arquivo, 'rb').read())
encoders.encode_base64(anexo)
anexo.add_header('Content-Disposition', f'attachment; filename={os.path.basename(caminho_arquivo)}')

# Adicione o anexo ao email
msg.attach(anexo)

# Conecte-se ao servidor SMTP (neste caso, Gmail)
servidor_smtp = smtplib.SMTP('smtp.gmail.com', 587)
servidor_smtp.starttls()

# Faça login na sua conta de email
servidor_smtp.login(de_email, senha)

# Envie o email
servidor_smtp.sendmail(de_email, para_email, msg.as_string())

# Encerre a conexão
servidor_smtp.quit()

print('Email com anexo enviado com sucesso.')


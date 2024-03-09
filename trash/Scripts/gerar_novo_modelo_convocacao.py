import docx
from docxtpl import DocxTemplate
from pathlib import Path
import sys
import os
import subprocess
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QVBoxLayout, QWidget
from PyQt5.QtCore import QObject, pyqtSignal

# Dicionário de padrões a serem substituídos
padroes = {
    'DIRETORIA': '',
    'NOME DA ESCOLA': '',
    'RUA': '',
    'NÚMERO DO ENDEREÇO': '',
    'BAIRRO': '',
    'CIDADE': '',
    'ESTADO': '',
    'CEP': '',
    'TELEFONE': '',
    'EMAIL': ''
}

# Classe para sinais
class Sinais(QObject):
    fechar_janela = pyqtSignal()

# Classe principal da interface gráfica
class Formulario(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Formulário de Dados")
        self.setGeometry(100, 100, 400, 400)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        # Crie campos de entrada de texto para os dados
        for padrao in padroes:
            label = QLabel(f"Digite o valor para {padrao}:")
            input_field = QLineEdit()
            self.layout.addWidget(label)
            self.layout.addWidget(input_field)

            # Conecte o evento de edição ao método para atualizar os dados
            input_field.textChanged.connect(self.atualizar_dados)

        # Botão para realizar a substituição nos documentos
        self.botao_substituir = QPushButton("Substituir Padrões no Documento")
        self.botao_substituir.clicked.connect(self.substituir_no_documento)
        self.layout.addWidget(self.botao_substituir)

        self.central_widget.setLayout(self.layout)

        # Configurar sinais para fechar a janela
        self.sinais = Sinais()
        self.sinais.fechar_janela.connect(self.fechar_janela_event)

    # Atualize o dicionário de padrões com os dados inseridos
    def atualizar_dados(self):
        for i, padrao in enumerate(padroes):
            padroes[padrao] = self.layout.itemAt(i * 2 + 1).widget().text()

    # Substitua os padrões no documento Word
    def substituir_no_documento(self):
        # Obtém o diretório do script
        script_path = Path(__file__).resolve().parent

        # Obtém o caminho absoluto para o arquivo
        document_path_base_dados = script_path.parent / "BasesDeDados" / "CONVOCACAO_PARA_COMPENSAR_FALTAS.docx"

        doc = DocxTemplate(document_path_base_dados)
        doc.render(padroes)

        # Salve o novo documento Word
        document_path_arquivos_gerados = script_path.parent / "ArquivosGerados" / "modelo_convocacao.docx"
        doc.save(Path(__file__).parent / document_path_arquivos_gerados)

        print(f'Documento gerado e salvo em ')

    def closeEvent(self, event):
        self.sinais.fechar_janela.emit()

    def fechar_janela_event(self):
        # Após fechar a janela, execute o último comando
        executar_script_gerar_arquivo_convocacao_word()

def executar_script_gerar_arquivo_convocacao_word():
    # Obter o diretório do arquivo em execução
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))

    # Nome do arquivo a ser executado (neste caso, na mesma pasta)
    caminho_segundo_script = "gerar_arquivo_convocacao_word.py"

    # Caminho completo para o segundo script
    caminho_completo = os.path.join(diretorio_atual, caminho_segundo_script)

    # Executar o segundo script
    subprocess.call(["python", caminho_completo])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = Formulario()
    janela.show()
    sys.exit(app.exec_())

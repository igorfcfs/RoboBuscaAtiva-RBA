import customtkinter as ctk
import os
import subprocess
from docxtpl import DocxTemplate
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import messagebox

deadline = datetime(2024, 12, 31, 23, 59, 59)

if datetime.today() <= deadline:
    diretorio_padrao = "../Planilias"

    def alterar_diretorio_padrao():
        pass
    def executar_script_renomeia_tabelas():
        diretorio_atual = Path(__file__).parent / "renomeia_tabelas.py"
        subprocess.call(["python", diretorio_atual])

    def executar_script_gerar_planilia_alunos_convocados():
        diretorio_atual = Path(__file__).parent / "gerar_planilia_alunos_convocados.py"
        subprocess.call(["python", diretorio_atual])

        # Obter o diretório do arquivo em execução
        diretorio_atual = os.path.dirname(os.path.abspath(__file__))
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
            # Dicionário de padrões a serem substituídos
            padroes = {
                'REGIAO': regiao_escola.upper(),
                'DEPARTAMENTO': departamento.upper(),
                'RUA': rua.capitalize(),
                'NUMERO': numero_endereco.capitalize(),
                'BAIRRO': bairro.capitalize(),
                'CIDADE': cidade.capitalize(),
                'ESTADO': estado.capitalize(),
                'CEP': cep,
                'TELEFONE': telefone,
                'EMAIL': email,
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
        
    def janela_Modelo_Conv():
        def salvar_informacoes():
            global departamento, regiao_escola, rua, numero_endereco, bairro, cidade, estado, cep, telefone, email
            regiao_escola = entry_regiao_escola.get()
            departamento = entry_departamento_escola.get()
            rua = entry_rua.get()
            numero_endereco = entry_numero_endereco.get()
            bairro = entry_bairro.get()
            cidade = entry_cidade.get()
            estado = entry_estado.get()
            cep = entry_cep.get()
            telefone = entry_telefone.get()
            email = entry_email.get()

            janela_modelo_convocacao.destroy()

        janela_modelo_convocacao = ctk.CTkToplevel()
        janela_modelo_convocacao.resizable(False, False)
        janela_modelo_convocacao.geometry("800x500")
        janela_modelo_convocacao.title("Modelo Convocação")
        
        # Define o ícone da janela
        # janela_modelo_convocacao.iconbitmap("caminho/do/arquivo.ico")

        label_modelo_convocacao = ctk.CTkLabel(janela_modelo_convocacao, text="Por favor, preencha as informações da sua Escola:")
        label_modelo_convocacao.pack(pady=10)

        entry_departamento_escola = ctk.CTkEntry(janela_modelo_convocacao, placeholder_text="Departamento da Escola")
        entry_departamento_escola.pack(padx=20, pady=5)

        entry_regiao_escola = ctk.CTkEntry(janela_modelo_convocacao, placeholder_text="Região da Escola")
        entry_regiao_escola.pack(pady=5)

        entry_rua = ctk.CTkEntry(janela_modelo_convocacao, placeholder_text="Rua")
        entry_rua.pack(pady=5)

        entry_numero_endereco = ctk.CTkEntry(janela_modelo_convocacao, placeholder_text="Número do Endereço")
        entry_numero_endereco.pack(pady=5)

        entry_bairro = ctk.CTkEntry(janela_modelo_convocacao, placeholder_text="Bairro")
        entry_bairro.pack(pady=5)

        entry_cidade = ctk.CTkEntry(janela_modelo_convocacao, placeholder_text="Cidade")
        entry_cidade.pack(pady=5)

        entry_estado = ctk.CTkEntry(janela_modelo_convocacao, placeholder_text="Estado")
        entry_estado.pack(pady=5)

        entry_cep = ctk.CTkEntry(janela_modelo_convocacao, placeholder_text="CEP")
        entry_cep.pack(pady=5)

        entry_telefone = ctk.CTkEntry(janela_modelo_convocacao, placeholder_text="Telefone")
        entry_telefone.pack(pady=5)

        entry_email = ctk.CTkEntry(janela_modelo_convocacao, placeholder_text="Email")
        entry_email.pack(pady=5)

        botao_salvar = ctk.CTkButton(janela_modelo_convocacao, text="Salvar", command=salvar_informacoes)
        botao_salvar.pack(pady=20)
        
        # O metodo grab_set() garante que a janela de modelo_convocacao apareça acima da janela principal, ja o lift() faz o oposto
        janela_modelo_convocacao.grab_set()

    janela_princ = ctk.CTk()
    janela_princ.resizable(False, False)
    janela_princ.geometry("800x500")
    janela_princ.minsize(width=800, height=500)
    janela_princ.maxsize(width=800, height=500)
    janela_princ.title("Robô Busca Ativa - RBA")
    
    # Define o ícone da janela
    # janela_modelo_convocacao.iconbitmap("caminho/do/arquivo.ico")

    label_busca_ativa = ctk.CTkLabel(janela_princ, text="Robô Busca Ativa")
    label_busca_ativa.pack(pady=10)

    botao_modelo_conv = ctk.CTkButton(janela_princ, text="Criar Modelo de Convocação", command=janela_Modelo_Conv)
    botao_modelo_conv.pack(pady=15)

    botao_rodar_robo = ctk.CTkButton(janela_princ, text="Rodar", command=executar_tudo)
    botao_rodar_robo.pack(pady=30)

    diretorio_padrao = ctk.CTk

    janela_princ.mainloop()
else:    
    def mostrar_mensagem():
        messagebox.showinfo('Aviso', 'A data de avaliação do Robô Busca Ativa (RBA) expirou.\n\nMuito obrigado por ter nos ajudado com as melhorias do projeto.\n\nCaso tenha interesse em ser informado sobre novas atualizações, entre em contato via email "rba_automation@gmail.com".')

    # Chamar a função para exibir a mensagem quando a janela é inicializada
    mostrar_mensagem()
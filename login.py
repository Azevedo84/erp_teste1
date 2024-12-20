import sys
from banco_dados.conexao import conectar_banco_nuvem
from forms.tela_login import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.telas import icone
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtGui import QPixmap
import os
import inspect
import traceback
import pyzipper

from erp import TelaMenu


class TelaLogin(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        versao = f"Versão 2.01.005"
        self.label_Versao.setText(versao)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        pixmap = QPixmap('arquivos/Logo_sem_fundo.png')
        self.label_12.setPixmap(pixmap)

        pixmap1 = QPixmap('arquivos/imagens tela/login.png')
        self.label_3.setPixmap(pixmap1)

        pixmap2 = QPixmap('arquivos/imagens tela/cadeado.png')
        self.label_4.setPixmap(pixmap2)

        icone(self, "menu_menu.png")

        self.btn_Login.clicked.connect(self.verifica_usuario_senha)

        self.menu_principal = []

        self.arquivo_zip = "dados_user.zip"

        self.senha_arquivo = "minha_senha_forte"
        login_usuario, senha_usuario = self.ler_arquivo_zip(self.arquivo_zip, self.senha_arquivo)

        if login_usuario and senha_usuario:
            self.check_Lembrar.setChecked(True)
            self.verifica_arquivo_login(login_usuario, senha_usuario)

    def trata_excecao(self, nome_funcao, mensagem, arquivo, excecao):
        try:
            tb = traceback.extract_tb(excecao)
            num_linha_erro = tb[-1][1]

            traceback.print_exc()
            print(f'Houve um problema no arquivo: {arquivo} na função: "{nome_funcao}"\n{mensagem} {num_linha_erro}')
            self.mensagem_alerta(f'Houve um problema no arquivo:\n\n{arquivo}\n\n'
                                 f'Comunique o desenvolvedor sobre o problema descrito abaixo:\n\n'
                                 f'{nome_funcao}: {mensagem}')

            grava_erro_banco(nome_funcao, mensagem, arquivo, num_linha_erro)

        except Exception as e:
            nome_funcao_trat = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            tb = traceback.extract_tb(exc_traceback)
            num_linha_erro = tb[-1][1]
            print(f'Houve um problema no arquivo: {self.nome_arquivo} na função: "{nome_funcao_trat}"\n'
                  f'{e} {num_linha_erro}')
            grava_erro_banco(nome_funcao_trat, e, self.nome_arquivo, num_linha_erro)

    def mensagem_alerta(self, mensagem):
        try:
            alert = QMessageBox()
            alert.setIcon(QMessageBox.Warning)
            alert.setText(mensagem)
            alert.setWindowTitle("Atenção")
            alert.setStandardButtons(QMessageBox.Ok)
            alert.exec_()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_usuario_senha(self):
        try:
            usuario = self.line_Usuario.text()
            senha = self.line_Senha.text()

            if not usuario:
                self.mensagem_alerta('O campo "Nome do Usuário" não pode estar vazio!')
            elif not senha:
                self.mensagem_alerta('O campo "Senha" não pode estar vazio!')
            else:
                self.verifica_sql_login()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_sql_login(self):
        conecta = conectar_banco_nuvem()
        try:
            usuario = self.line_Usuario.text()
            senha = self.line_Senha.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT login, senha_login "
                           f"FROM USUARIO "
                           f"where login = '{usuario}' "
                           f"and ativo = 'S';")
            dados_usuario = cursor.fetchall()
            
            if not dados_usuario:
                self.mensagem_alerta(f'O Usuário {usuario} não existe!')
            else:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, login, senha_login "
                               f"FROM USUARIO "
                               f"where login = '{usuario}' "
                               f"and senha_login = '{senha}' "
                               f"and ativo = 'S';")
                dados_usuario_senha = cursor.fetchall()
                print(dados_usuario_senha)

                if not dados_usuario_senha:
                    self.mensagem_alerta(f'A senha do Usuário {usuario} está incorreta!')
                else:
                    id_func = dados_usuario_senha[0][0]
                    self.abrir_menu_principal(usuario, senha, id_func)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def abrir_menu_principal(self, usuario, senha, id_func):
        try:
            if self.check_Lembrar.isChecked():
                self.criar_ou_atualizar_zip(self.arquivo_zip, usuario, senha, self.senha_arquivo)
            elif not self.check_Lembrar.isChecked():
                usuario_zerado = ""
                senha_zerada = ""
                self.criar_ou_atualizar_zip(self.arquivo_zip, usuario_zerado, senha_zerada, self.senha_arquivo)

            print("abrir tela menu pelo botão", id_func)
            self.menu_principal = TelaMenu(id_func)
            self.menu_principal.show()
            self.close()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def criar_ou_atualizar_zip(self, caminho_arquivo, login_user, senha_user, senha):
        try:
            conteudo = f"login:{login_user}\nsenha:{senha_user}"

            with pyzipper.AESZipFile(caminho_arquivo, 'w', compression=pyzipper.ZIP_LZMA) as zipf:
                zipf.setpassword(senha.encode())
                zipf.setencryption(pyzipper.WZ_AES)
                zipf.writestr("dados.txt", conteudo)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def ler_arquivo_zip(self, caminho_arquivo, senha):
        try:
            if not os.path.exists(caminho_arquivo):
                self.criar_ou_atualizar_zip(caminho_arquivo, "", "", senha)

            with pyzipper.AESZipFile(caminho_arquivo, 'r') as zipf:
                zipf.setpassword(senha.encode())
                conteudo = {}

                login_usuario = ""
                senha_usuario = ""

                for nome_arquivo in zipf.namelist():
                    conteudo[nome_arquivo] = zipf.read(nome_arquivo).decode()

                for arquivo, informacoes_arquivo in conteudo.items():
                    linhas = informacoes_arquivo.split('\n')

                    if 'login:' in linhas[0]:
                        login_usuario = linhas[0].split('login:')[1].strip()

                    if len(linhas) > 1 and 'senha:' in linhas[1]:
                        senha_usuario = linhas[1].split('senha:')[1].strip()

                print(login_usuario, senha_usuario)

                return login_usuario, senha_usuario

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_arquivo_login(self, usuario, senha):
        conecta = conectar_banco_nuvem()
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT login, senha_login FROM USUARIO where login = '{usuario}';")
            dados_usuario = cursor.fetchall()

            if dados_usuario:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, login, senha_login "
                               f"FROM USUARIO where login = '{usuario}' and senha_login = '{senha}';")
                dados_usuario_senha = cursor.fetchall()

                if dados_usuario_senha:
                    id_func = dados_usuario_senha[0][0]

                    print("abrir tela menu direto", id_func)

                    self.menu_principal = TelaMenu(id_func)
                    self.menu_principal.show()
                    sys.exit(qt.exec_())

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    telamenu = TelaLogin()
    telamenu.show()
    sys.exit(qt.exec_())

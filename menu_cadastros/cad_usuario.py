import sys
from banco_dados.conexao import conectar_banco_nuvem
from forms.tela_cad_usuario import *
from banco_dados.controle_erros import grava_erro_banco
from banco_dados.bc_consultas import proximo_registro_autoincremento
from comandos.tabelas import lanca_tabela, layout_cabec_tab, extrair_tabela
from comandos.telas import tamanho_aplicacao, icone
from comandos.lines import validador_so_numeros
from comandos.conversores import timestamp_brasileiro
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import inspect
import os
from datetime import date
import traceback


class TelaCadastroUsuario(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_cadastro.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Lista)

        self.lanca_numero_id()
        self.definir_data_emissao()
        self.manipula_dados_tabela()

        self.definir_combo_funcionario()

        validador_so_numeros(self.line_Num)
        self.line_Num.setReadOnly(True)

        self.table_Lista.viewport().installEventFilter(self)

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)
        self.btn_Limpar.clicked.connect(self.reiniciando_tela)
        self.btn_Excluir.clicked.connect(self.excluir_cadastro)

        self.btn_Consulta.clicked.connect(self.procura_palavra)
        self.line_Consulta.returnPressed.connect(lambda: self.procura_palavra())

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
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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

    def pergunta_confirmacao(self, mensagem):
        try:
            confirmacao = QMessageBox()
            confirmacao.setIcon(QMessageBox.Question)
            confirmacao.setText(mensagem)
            confirmacao.setWindowTitle("Confirmação")

            sim_button = confirmacao.addButton("Sim", QMessageBox.YesRole)
            nao_button = confirmacao.addButton("Não", QMessageBox.NoRole)

            confirmacao.setDefaultButton(nao_button)

            confirmacao.exec_()

            if confirmacao.clickedButton() == sim_button:
                return True
            else:
                return False

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_numero_id(self):
        try:
            num_id = proximo_registro_autoincremento("USUARIO")
            self.line_Num.setText(str(num_id))
            self.line_Login.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_data_emissao(self):
        try:
            data_hoje = date.today()
            self.date_Emissao.setDate(data_hoje)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_combo_funcionario(self):
        conecta = conectar_banco_nuvem()
        try:
            tabela = []

            self.combo_Funcionario.clear()
            tabela.append("")

            cur = conecta.cursor()
            cur.execute(f"SELECT id, nome FROM FUNCIONARIO where ativo = 'S' order by nome;")
            detalhes_func = cur.fetchall()

            for dadus in detalhes_func:
                ides, func = dadus
                msg = f"{ides} - {func}"
                tabela.append(msg)

            self.combo_Funcionario.addItems(tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def manipula_dados_tabela(self):
        conecta = conectar_banco_nuvem()
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"select user.id, user.criacao, "
                           f"user.login, user.email, func.nome, user.ativo "
                           f"from USUARIO as user "
                           f"INNER JOIN FUNCIONARIO as func ON user.funcionario_id = func.id "
                           f"order by user.login;")
            select_numero = cursor.fetchall()

            if select_numero:
                for i in select_numero:
                    num_id, data, login, email, func, ativo = i

                    data_formatada = timestamp_brasileiro(data)

                    dados = (num_id, data_formatada, login, email, func, ativo)
                    tabela_nova.append(dados)

            if tabela_nova:
                lanca_tabela(self.table_Lista, tabela_nova)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def reiniciando_tela(self):
        try:
            self.line_Login.clear()
            self.line_Email.clear()
            self.check_Ativo.setChecked(False)
            self.line_Consulta.clear()

            self.combo_Funcionario.setCurrentText("")

            self.definir_combo_funcionario()

            self.lanca_numero_id()
            self.manipula_dados_tabela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def procura_palavra(self):
        conecta = conectar_banco_nuvem()
        try:
            tabela_nova = []

            palavra_consulta = self.line_Consulta.text()

            if not palavra_consulta:
                self.mensagem_alerta(f'O Campo "Consulta Login" não pode estar vazio!')
                self.line_Consulta.clear()
            else:
                palavra_maiuscula = palavra_consulta.upper()

                cursor = conecta.cursor()
                cursor.execute(f"select DISTINCT user.id, user.criacao, "
                               f"user.login, user.email, func.nome, user.ativo "
                               f"from USUARIO as user "
                               f"INNER JOIN FUNCIONARIO as func ON user.funcionario_id = func.id "
                               f"WHERE user.login LIKE '%{palavra_maiuscula}%' "
                               f"order by user.login;")
                palavra = cursor.fetchall()

                if not palavra:
                    self.mensagem_alerta(f'Não foi encontrado nenhum Usuário com o Nome:\n  "{palavra_consulta}"!')
                    self.line_Consulta.clear()
                else:
                    for i in palavra:
                        num_id, data, login, email, func, ativo = i

                        data_formatada = timestamp_brasileiro(data)

                        dados = (num_id, data_formatada, login, email, func, ativo)
                        tabela_nova.append(dados)

                if tabela_nova:
                    lanca_tabela(self.table_Lista, tabela_nova)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def eventFilter(self, source, event):
        try:
            if (event.type() == QtCore.QEvent.MouseButtonDblClick and
                    event.buttons() == QtCore.Qt.LeftButton and
                    source is self.table_Lista.viewport()):
                item = self.table_Lista.currentItem()

                extrai_recomendados = extrair_tabela(self.table_Lista)
                item_selecionado = extrai_recomendados[item.row()]

                num_id, data, login, email, func, ativo = item_selecionado

                self.line_Num.setText(num_id)
                self.line_Login.setText(login)
                if email:
                    self.line_Email.setText(email)

                item_count = self.combo_Funcionario.count()
                for i in range(item_count):
                    item_text = self.combo_Funcionario.itemText(i)
                    if func in item_text:
                        self.combo_Funcionario.setCurrentText(item_text)

                if ativo == "S":
                    self.check_Ativo.setChecked(True)
                else:
                    self.check_Ativo.setChecked(False)

            return super(QMainWindow, self).eventFilter(source, event)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_cadastro(self):
        conecta = conectar_banco_nuvem()
        try:
            num_id = self.line_Num.text()
            login_user = self.line_Login.text()
            email = self.line_Email.text()

            if not num_id:
                self.mensagem_alerta('O campo "Código" não pode estar vazio!   ')
                self.line_Num.setFocus()
            elif num_id == "0":
                self.mensagem_alerta('O campo "Código" não pode ser "0"!   ')
                self.line_Num.clear()
                self.line_Num.setFocus()
            elif not login_user:
                self.mensagem_alerta('O campo "Login" não pode estar vazio!   ')
                self.line_Login.clear()
                self.line_Login.setFocus()
            elif login_user == "0":
                self.mensagem_alerta('O campo "Login" não pode ser "0"!   ')
                self.line_Login.clear()
                self.line_Login.setFocus()
            elif not email:
                self.mensagem_alerta('O campo "Email" não pode estar vazio!   ')
                self.line_Email.clear()
                self.line_Email.setFocus()
            elif email == "0":
                self.mensagem_alerta('O campo "Email" não pode ser "0"!   ')
                self.line_Email.clear()
                self.line_Email.setFocus()
            else:
                cursor = conecta.cursor()
                cursor.execute(f"select user.id, user.criacao, "
                               f"user.login, user.email, func.nome, user.ativo "
                               f"from USUARIO as user "
                               f"INNER JOIN FUNCIONARIO as func ON user.funcionario_id = func.id "
                               f"where user.id = {num_id};")
                select_usuario = cursor.fetchall()

                if not select_usuario:
                    self.mensagem_alerta(f'O cadastro do Usuário Nº {num_id} não existe!')
                else:
                    msg = f'Tem certeza que deseja excluir o Usuário {login_user}?'
                    if self.pergunta_confirmacao(msg):
                        cursor = conecta.cursor()
                        cursor.execute(f"DELETE FROM USUARIO WHERE id = {num_id};")

                        conecta.commit()

                        self.mensagem_alerta(f"Cadastro do Usuário {login_user} foi excluído com Sucesso!")

                        self.reiniciando_tela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def verifica_salvamento(self):
        try:
            num_id = self.line_Num.text()
            login_user = self.line_Login.text()
            email = self.line_Email.text()

            if not num_id:
                self.mensagem_alerta('O campo "Código" não pode estar vazio!   ')
                self.line_Num.setFocus()
            elif num_id == "0":
                self.mensagem_alerta('O campo "Código" não pode ser "0"!   ')
                self.line_Num.clear()
                self.line_Num.setFocus()
            elif not login_user:
                self.mensagem_alerta('O campo "Login" não pode estar vazio!   ')
                self.line_Login.clear()
                self.line_Login.setFocus()
            elif login_user == "0":
                self.mensagem_alerta('O campo "Login" não pode ser "0"!   ')
                self.line_Login.clear()
                self.line_Login.setFocus()
            elif not email:
                self.mensagem_alerta('O campo "Email" não pode estar vazio!   ')
                self.line_Email.clear()
                self.line_Email.setFocus()
            elif email == "0":
                self.mensagem_alerta('O campo "Email" não pode ser "0"!   ')
                self.line_Email.clear()
                self.line_Email.setFocus()
            else:
                self.salvar_dados()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_dados(self):
        conecta = conectar_banco_nuvem()
        try:
            num_id = self.line_Num.text()

            login_user = self.line_Login.text()
            login_user_maiuscula = login_user.upper()

            email = self.line_Email.text()

            if self.check_Ativo.isChecked():
                ativo = "S"
            else:
                ativo = "N"

            func = self.combo_Funcionario.currentText()
            functete = func.find(" - ")
            id_func = func[:functete]

            cursor = conecta.cursor()
            cursor.execute(f"select user.id, user.criacao, "
                           f"user.login, user.email, user.funcionario_id, user.ativo "
                           f"from USUARIO as user "
                           f"INNER JOIN FUNCIONARIO as func ON user.funcionario_id = func.id "
                           f"where user.id = {num_id};")
            select_usuario = cursor.fetchall()

            if select_usuario:
                num_id_b, data_b, login_b, email_b, id_func_b, ativo_b = select_usuario[0]

                campos_atualizados = []
                if login_user_maiuscula != login_b:
                    campos_atualizados.append(f"login = '{login_user_maiuscula}'")
                    campos_atualizados.append(f"senha_login = '{login_user_maiuscula}'")
                if email != email_b:
                    campos_atualizados.append(f"email = '{email}'")
                if id_func != id_func_b:
                    campos_atualizados.append(f"funcionario_id = {id_func}")
                if ativo != ativo_b:
                    campos_atualizados.append(f"ativo = '{ativo}'")

                if campos_atualizados:
                    msg = f'Deseja realmente atualizar o cadastro do Usuário?'
                    if self.pergunta_confirmacao(msg):
                        campos_update = ", ".join(campos_atualizados)

                        cursor.execute(f"UPDATE USUARIO SET {campos_update} "
                                       f"WHERE id = {num_id};")

                        conecta.commit()

                        msg = f'O cadastro do Usuário {login_user_maiuscula} foi atualizado com sucesso!'
                        self.mensagem_alerta(msg)

            else:
                msg = f'Deseja realmente cadastrar este Usuário?'
                if self.pergunta_confirmacao(msg):
                    cursor = conecta.cursor()
                    cursor.execute(f"Insert into USUARIO (login, senha_login, email, funcionario_id, ativo) "
                                   f"values ('{login_user_maiuscula}', '{login_user_maiuscula}', "
                                   f"'{email}', {id_func}, '{ativo}');")

                    conecta.commit()

                    msg = f'O cadastro do Usuário {login_user_maiuscula} foi criado com sucesso!'
                    self.mensagem_alerta(msg)

            self.reiniciando_tela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaCadastroUsuario()
    tela.show()
    qt.exec_()

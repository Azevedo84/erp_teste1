import sys
from banco_dados.conexao import conectar_banco_nuvem
from forms.tela_cad_local_estoque import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import lanca_tabela, layout_cabec_tab, extrair_tabela
from comandos.telas import tamanho_aplicacao, icone
from comandos.lines import validador_so_numeros
from comandos.conversores import timestamp_brasileiro
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import inspect
import os
from datetime import date
import traceback
from unidecode import unidecode


class TelaCadastroLocalEstoque(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_cadastro.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Lista)

        self.definir_data_emissao()
        self.manipula_dados_tabela()

        validador_so_numeros(self.line_Num)

        self.table_Lista.viewport().installEventFilter(self)

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)
        self.btn_Limpar.clicked.connect(self.reiniciando_tela)
        self.btn_Excluir.clicked.connect(self.excluir_cadastro)

        self.btn_Consulta.clicked.connect(self.procura_palavra)
        self.line_Consulta.returnPressed.connect(lambda: self.procura_palavra())

        self.line_Num.setFocus()

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

    def definir_data_emissao(self):
        try:
            data_hoje = date.today()
            self.date_Emissao.setDate(data_hoje)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela(self):
        conecta = conectar_banco_nuvem()
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"select id_siger, criacao, "
                           f"descricao, COALESCE(usar_estoque, ''), COALESCE(negativo, ''), COALESCE(acinplas, '') "
                           f"from ESTOQUE_LOCAL_EST "
                           f"order by descricao;")
            select_numero = cursor.fetchall()

            if select_numero:
                for i in select_numero:
                    id_func, data, nome, estoque, negativo, acinplas = i

                    data_formatada = timestamp_brasileiro(data)

                    dados = (id_func, data_formatada, nome, estoque, negativo, acinplas)
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
            self.line_Num.clear()
            self.line_Descricao.clear()
            self.check_Saldo_Estoque.setChecked(False)
            self.check_Negativo.setChecked(False)
            self.check_Acinplas.setChecked(False)
            self.line_Consulta.clear()

            self.manipula_dados_tabela()

            self.line_Num.setFocus()

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
                self.mensagem_alerta(f'O Campo "Consulta Nome" não pode estar vazio!')
                self.line_Consulta.clear()
            else:
                palavra_maiuscula = palavra_consulta.upper()

                cursor = conecta.cursor()
                cursor.execute(f"SELECT DISTINCT id_siger, criacao, "
                               f"descricao, COALESCE(usar_estoque, ''), COALESCE(negativo, ''), COALESCE(acinplas, '') "
                               f"FROM ESTOQUE_LOCAL_EST "
                               f"WHERE descricao LIKE '%{palavra_maiuscula}%' ORDER BY descricao;")
                palavra = cursor.fetchall()

                if not palavra:
                    self.mensagem_alerta(f'Não foi encontrado nenhum item com o nome:\n  "{palavra_consulta}"!')
                    self.line_Consulta.clear()
                else:
                    for i in palavra:
                        id_func, data, nome, estoque, negativo, acinplas = i

                        data_formatada = timestamp_brasileiro(data)

                        dados = (id_func, data_formatada, nome, estoque, negativo)
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

                id_func, criacao, nome, estoque, negativo, acinplas = item_selecionado

                self.line_Num.setText(id_func)
                self.line_Descricao.setText(nome)

                if estoque == "S":
                    self.check_Saldo_Estoque.setChecked(True)
                else:
                    self.check_Saldo_Estoque.setChecked(False)

                if negativo == "S":
                    self.check_Negativo.setChecked(True)
                else:
                    self.check_Negativo.setChecked(False)

                if acinplas == "S":
                    self.check_Acinplas.setChecked(True)
                else:
                    self.check_Acinplas.setChecked(False)

            return super(QMainWindow, self).eventFilter(source, event)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_cadastro(self):
        conecta = conectar_banco_nuvem()
        try:
            num_id = self.line_Num.text()
            nome = self.line_Descricao.text()

            if not num_id:
                self.mensagem_alerta('O campo "Código" não pode estar vazio!   ')
                self.line_Num.setFocus()
            elif num_id == "0":
                self.mensagem_alerta('O campo "Código" não pode ser "0"!   ')
                self.line_Num.clear()
                self.line_Num.setFocus()
            elif not nome:
                self.mensagem_alerta('O campo "Descrição" não pode estar vazio!   ')
                self.line_Descricao.clear()
                self.line_Descricao.setFocus()
            elif nome == "0":
                self.mensagem_alerta('O campo "Descrição" não pode ser "0"!   ')
                self.line_Descricao.clear()
                self.line_Descricao.setFocus()
            else:
                cursor = conecta.cursor()
                cursor.execute(f"select id_siger, criacao, descricao, COALESCE(usar_estoque, ''), "
                               f"COALESCE(negativo, ''), COALESCE(acinplas, '') "
                               f"from ESTOQUE_LOCAL_EST where id_siger = {num_id};")
                locais = cursor.fetchall()

                cursor = conecta.cursor()
                cursor.execute(f"SELECT COUNT(*) "
                               f"FROM MOVIMENTACAO "
                               f"WHERE local_estoque_id = {num_id};")
                movimentacao = cursor.fetchone()[0]

                if not locais:
                    self.mensagem_alerta(f'O cadastro do Local de Estoque Nº {num_id} não existe!')
                elif movimentacao > 0:
                    self.mensagem_alerta(f'O cadastro do Local de Estoque Nº {num_id} não pode ser excluído, '
                                         f'pois existem vínculos com a movimentação!')
                else:
                    msg = f'Tem certeza que deseja excluir o Local de Estoque {nome}?'
                    if self.pergunta_confirmacao(msg):
                        cursor = conecta.cursor()
                        cursor.execute(f"DELETE FROM ESTOQUE_LOCAL_EST WHERE id_siger = {num_id};")

                        conecta.commit()

                        self.mensagem_alerta(f"Cadastro do Local de Estoque {nome} foi excluído com Sucesso!")

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
            nome = self.line_Descricao.text()

            if not num_id:
                self.mensagem_alerta('O campo "Código" não pode estar vazio!')
                self.line_Num.setFocus()
            elif num_id == "0":
                self.mensagem_alerta('O campo "Código" não pode ser "0"!')
                self.line_Num.clear()
                self.line_Num.setFocus()
            elif not nome:
                self.mensagem_alerta('O campo "Descrição" não pode estar vazio!   ')
                self.line_Descricao.clear()
                self.line_Descricao.setFocus()
            elif nome == "0":
                self.mensagem_alerta('O campo "Descrição" não pode ser "0"!   ')
                self.line_Descricao.clear()
                self.line_Descricao.setFocus()
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

            descr = self.line_Descricao.text()
            descr_maiuscula = descr.upper()
            descr_sem_acentos = unidecode(descr_maiuscula)

            if self.check_Saldo_Estoque.isChecked():
                estoque = "S"
            else:
                estoque = "N"

            if self.check_Negativo.isChecked():
                negativo = "S"
            else:
                negativo = "N"

            if self.check_Acinplas.isChecked():
                acinplas = "S"
            else:
                acinplas = "N"

            cursor = conecta.cursor()
            cursor.execute(f"select id_siger, criacao, descricao, COALESCE(usar_estoque, ''), "
                           f"COALESCE(negativo, ''), COALESCE(acinplas, '') "
                           f"from ESTOQUE_LOCAL_EST where id_siger = {num_id};")
            locais = cursor.fetchall()

            if locais:
                num_id_b, criacao_b, nome_b, estoque_b, negativo_b, acinplas_b = locais[0]

                campos_atualizados = []
                if descr_sem_acentos != nome_b:
                    campos_atualizados.append(f"descricao = '{descr_sem_acentos}'")
                if estoque != estoque_b:
                    campos_atualizados.append(f"usar_estoque = '{estoque}'")
                if negativo != negativo_b:
                    campos_atualizados.append(f"negativo = '{negativo}'")
                if acinplas != acinplas_b:
                    campos_atualizados.append(f"acinplas = '{acinplas}'")

                if campos_atualizados:
                    msg = f'Deseja realmente atualizar o cadastro do Local de Estoque?'
                    if self.pergunta_confirmacao(msg):
                        campos_update = ", ".join(campos_atualizados)

                        cursor.execute(f"UPDATE ESTOQUE_LOCAL_EST SET {campos_update} "
                                       f"WHERE id_siger = {num_id_b};")

                        conecta.commit()

                        msg = f'O cadastro do Local de Estoque {descr_sem_acentos} foi atualizado com sucesso!'
                        self.mensagem_alerta(msg)

            else:
                msg = f'Deseja realmente cadastrar este Local de Estoque?'
                if self.pergunta_confirmacao(msg):
                    cursor = conecta.cursor()
                    cursor.execute(f"Insert into ESTOQUE_LOCAL_EST (id_siger, descricao, usar_estoque, "
                                   f"negativo, acinplas) "
                                   f"values ({num_id}, '{descr_sem_acentos}', '{estoque}', "
                                   f"'{negativo}', '{acinplas}');")

                    conecta.commit()

                    msg = f'O cadastro do Local de Estoque  {descr_sem_acentos} foi criado com sucesso!'
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
    tela = TelaCadastroLocalEstoque()
    tela.show()
    qt.exec_()

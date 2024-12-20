import sys
from banco_dados.conexao import conectar_banco_nuvem
from forms.tela_cad_servico_interno import *
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
from unidecode import unidecode


class TelaCadastroServicoInterno(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_cadastro.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Lista)
        layout_cabec_tab(self.table_Prioridade)

        self.lanca_numero_id()
        self.definir_data_emissao()
        self.manipula_dados_tabela()
        self.definir_combo_funcionario()
        self.definir_combo_prioridade()

        self.table_Lista.viewport().installEventFilter(self)

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)
        self.btn_Limpar.clicked.connect(self.reiniciando_tela)
        self.btn_Excluir.clicked.connect(self.excluir_cadastro)

        self.btn_Adicionar.clicked.connect(self.adicionar_func_prioridade)
        self.btn_Excluir_Func.clicked.connect(self.excluir_funcionario)

        self.btn_Consulta.clicked.connect(self.procura_palavra)
        self.line_Consulta.returnPressed.connect(lambda: self.procura_palavra())

        validador_so_numeros(self.line_Num)
        self.line_Num.setReadOnly(True)

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

    def manipula_dados_tabela(self):
        conecta = conectar_banco_nuvem()
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"select id, criacao, descricao "
                           f"from SERVICO_INTERNO "
                           f"order by descricao;")
            select_numero = cursor.fetchall()

            if select_numero:
                for i in select_numero:
                    id_conj, data, descricao = i

                    data_formatada = timestamp_brasileiro(data)

                    dados = (id_conj, data_formatada, descricao)
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

    def definir_combo_funcionario(self):
        conecta = conectar_banco_nuvem()
        try:
            tabela = []

            self.combo_Func.clear()
            tabela.append("")

            dados_tabela = extrair_tabela(self.table_Prioridade)
            funcionarios_prioridade = set()
            if dados_tabela:
                for dadus in dados_tabela:
                    funcionarios_prioridade.add(dadus[0])

            cur = conecta.cursor()
            cur.execute(f"SELECT id, nome FROM FUNCIONARIO where ativo = 'S' order by nome;")
            detalhes_func = cur.fetchall()

            for ides, func in detalhes_func:
                if func not in funcionarios_prioridade:
                    tabela.append(func)

            self.combo_Func.addItems(tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def definir_combo_prioridade(self):
        conecta = conectar_banco_nuvem()
        try:
            tabela = []

            self.combo_Prioridade.clear()
            tabela.append("")
            tabela.append("PRINCIPAL")

            dados_tabela = extrair_tabela(self.table_Prioridade)
            if dados_tabela:
                for dadus in dados_tabela:
                    funcionario = dadus[0]
                    msg = f"APÓS {funcionario}"
                    tabela.append(msg)

            self.combo_Prioridade.addItems(tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def adicionar_func_prioridade(self):
        try:
            descricao_servico = self.line_Descricao.text()
            funcionario = self.combo_Func.currentText()
            prioridade = self.combo_Prioridade.currentText()

            if descricao_servico:

                if funcionario and prioridade:
                    num_rows = self.table_Prioridade.rowCount()
                    lista_prioridade = []

                    for row in range(num_rows):
                        item = self.table_Prioridade.item(row, 0)
                        if item is not None:
                            lista_prioridade.append(item.text())

                    if prioridade == "PRINCIPAL":
                        lista_prioridade.insert(0, funcionario)
                    elif prioridade.startswith("APÓS"):
                        nome_anterior = prioridade.split(' ', 1)[1]

                        if nome_anterior not in lista_prioridade:
                            raise ValueError(f"O funcionário '{nome_anterior}' não está na lista de prioridades.")

                        idx = lista_prioridade.index(nome_anterior)
                        lista_prioridade.insert(idx + 1, funcionario)
                    else:
                        raise ValueError("Prioridade inválida selecionada.")

                    self.table_Prioridade.setRowCount(0)

                    for idx, nome_funcionario in enumerate(lista_prioridade):
                        self.table_Prioridade.insertRow(idx)
                        self.table_Prioridade.setItem(idx, 0, QtWidgets.QTableWidgetItem(nome_funcionario))

                    print("Funcionário adicionado com sucesso à prioridade.")

                    self.definir_combo_prioridade()
                    self.definir_combo_funcionario()

            else:
                self.mensagem_alerta("Defina a Descrição do Serviço Interno!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_funcionario(self):
        try:
            current_row = self.table_Prioridade.currentRow()
            if current_row == -1:
                raise ValueError("Nenhum funcionário selecionado para exclusão.")

            nome_funcionario = self.table_Prioridade.item(current_row, 0).text()

            self.table_Prioridade.removeRow(current_row)

            self.definir_combo_funcionario()
            self.definir_combo_prioridade()

            print(f"Funcionário '{nome_funcionario}' excluído com sucesso.")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_numero_id(self):
        try:
            num_id = proximo_registro_autoincremento("SERVICO_INTERNO")
            self.line_Num.setText(str(num_id))
            self.line_Descricao.setFocus()

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

    def reiniciando_tela(self):
        try:
            self.line_Descricao.clear()
            self.line_Consulta.clear()

            self.table_Prioridade.setRowCount(0)

            self.lanca_numero_id()
            self.definir_data_emissao()
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
                self.mensagem_alerta(f'O Campo "Consulta Descrição" não pode estar vazio!')
                self.line_Consulta.clear()
            else:
                palavra_maiuscula = palavra_consulta.upper()

                cursor = conecta.cursor()
                cursor.execute(f"SELECT DISTINCT id, criacao, descricao "
                               f"FROM SERVICO_INTERNO "
                               f"WHERE descricao LIKE '%{palavra_maiuscula}%' "
                               f"ORDER BY descricao;")
                palavra = cursor.fetchall()

                if not palavra:
                    self.mensagem_alerta(f'Não foi encontrado nenhum item com o nome:\n  "{palavra_consulta}"!')
                    self.line_Consulta.clear()
                else:
                    for i in palavra:
                        id_conjunto, data, descricao = i

                        data_formatada = timestamp_brasileiro(data)

                        dados = (id_conjunto, data_formatada, descricao)
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
        conecta = conectar_banco_nuvem()
        try:
            if (event.type() == QtCore.QEvent.MouseButtonDblClick and
                    event.buttons() == QtCore.Qt.LeftButton and
                    source is self.table_Lista.viewport()):
                item = self.table_Lista.currentItem()

                extrai_recomendados = extrair_tabela(self.table_Lista)
                item_selecionado = extrai_recomendados[item.row()]

                id_servico, criacao, descricao = item_selecionado

                self.table_Prioridade.setRowCount(0)

                self.line_Num.setText(id_servico)
                self.line_Descricao.setText(descricao)

                cursor = conecta.cursor()
                cursor.execute(f"select func.nome "
                               f"from FUNCIONARIO_SERVICO_INT as fun_ser "
                               f"INNER JOIN FUNCIONARIO as func ON fun_ser.funcionario_id = func.id "
                               f"where fun_ser.processo_interno_id = {id_servico} "
                               f"order by fun_ser.prioridade;")
                dados_banco = cursor.fetchall()

                if dados_banco:
                    lanca_tabela(self.table_Prioridade, dados_banco)

                self.definir_combo_funcionario()
                self.definir_combo_prioridade()

            return super(QMainWindow, self).eventFilter(source, event)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

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
                self.mensagem_alerta('O campo "Descrição:" não pode estar vazio!   ')
                self.line_Descricao.clear()
                self.line_Descricao.setFocus()
            elif nome == "0":
                self.mensagem_alerta('O campo "Descrição:" não pode ser "0"!   ')
                self.line_Descricao.clear()
                self.line_Descricao.setFocus()
            else:
                cursor = conecta.cursor()
                cursor.execute(f"select id, criacao, descricao "
                               f"from SERVICO_INTERNO where id = {num_id};")
                servicos = cursor.fetchall()

                cursor = conecta.cursor()
                cursor.execute(f"SELECT * "
                               f"FROM PRODUTO_SERVICO_INTERNO "
                               f"WHERE SERVICO_INTERNO_ID = {num_id};")
                produtos = cursor.fetchall()

                if not servicos:
                    self.mensagem_alerta(f'O cadastro do Serviço Interno Nº {num_id} não existe!')
                elif produtos:
                    self.mensagem_alerta(f'O cadastro do Serviço Interno Nº {num_id} não pode ser excluído, '
                                         f'pois existem vinculos com produtos!')
                else:
                    msg = f'Tem certeza que deseja excluir o Serviço Interno {nome}?'
                    if self.pergunta_confirmacao(msg):
                        cursor = conecta.cursor()
                        cursor.execute(f"DELETE FROM SERVICO_INTERNO WHERE id = {num_id};")

                        conecta.commit()

                        self.mensagem_alerta(f"Cadastro do Serviço Interno {nome} foi excluído com Sucesso!")

                        self.reiniciando_tela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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
                self.mensagem_alerta('O campo "Descrição:" não pode estar vazio!   ')
                self.line_Descricao.clear()
                self.line_Descricao.setFocus()
            elif nome == "0":
                self.mensagem_alerta('O campo "Descrição:" não pode ser "0"!   ')
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

            cursor = conecta.cursor()
            cursor.execute(f"select id, criacao, descricao "
                           f"from SERVICO_INTERNO where id = {num_id};")
            servicos = cursor.fetchall()

            if servicos:
                num_id_b, criacao_b, conj_b = servicos[0]

                campos_atualizados = []
                if descr_sem_acentos != conj_b:
                    campos_atualizados.append(f"descricao = '{descr_sem_acentos}'")

                if campos_atualizados:
                    msg = f'Deseja realmente atualizar o cadastro do Serviço Interno?'
                    if self.pergunta_confirmacao(msg):
                        campos_update = ", ".join(campos_atualizados)
                        """
                        cursor.execute(f"UPDATE SERVICO_INTERNO SET {campos_update} "
                                       f"WHERE id = {num_id_b};")
                        conecta.commit()
                        """

                        msg = f'O cadastro do Serviço Interno {descr_sem_acentos} foi atualizado com sucesso!'
                        self.mensagem_alerta(msg)

            else:
                msg = f'Deseja realmente cadastrar este Serviço Interno?'
                if self.pergunta_confirmacao(msg):
                    """
                    cursor = conecta.cursor()
                    cursor.execute(f"Insert into SERVICO_INTERNO (descricao) "
                                   f"values ('{descr_sem_acentos}');")

                    conecta.commit()
                    """

            self.salvar_servico_funcionario(num_id, descr_sem_acentos)
            self.reiniciando_tela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def salvar_servico_funcionario(self, id_servico_interno, descr_servico):
        conecta = conectar_banco_nuvem()
        try:

            dados_tabela = extrair_tabela(self.table_Prioridade)

            if dados_tabela:
                for indice, i in enumerate(dados_tabela, start=1):
                    nome_func = i[0]
                    print(nome_func)

                    cursor = conecta.cursor()
                    cursor.execute(f"select id, nome "
                                   f"from FUNCIONARIO "
                                   f"where nome = '{nome_func}';")
                    dados_funcionario = cursor.fetchall()

                    if dados_funcionario:
                        id_func = dados_funcionario[0][0]

                        cursor = conecta.cursor()
                        cursor.execute(f"select funcionario_id, processo_interno_id, prioridade "
                                       f"from FUNCIONARIO_SERVICO_INT "
                                       f"where funcionario_id = {id_func} "
                                       f"and processo_interno_id = {id_servico_interno};")
                        dados_banco = cursor.fetchall()
                        if dados_banco:
                            print(dados_banco)
                        else:
                            cursor = conecta.cursor()
                            cursor.execute(f"Insert into FUNCIONARIO_SERVICO_INT (funcionario_id, "
                                           f"processo_interno_id, prioridade) "
                                           f"values ({id_func}, {id_servico_interno}, {indice});")

                            conecta.commit()

            msg = f'O cadastro do Serviço Interno {descr_servico} foi criado com sucesso!'
            self.mensagem_alerta(msg)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaCadastroServicoInterno()
    tela.show()
    qt.exec_()

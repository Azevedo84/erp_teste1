import sys
from banco_dados.conexao import conectar_banco_nuvem
from forms.tela_pre_incluir import *
from banco_dados.controle_erros import grava_erro_banco
from banco_dados.bc_consultas import proximo_registro_autoincremento
from comandos.tabelas import lanca_tabela, layout_cabec_tab, extrair_tabela
from comandos.telas import tamanho_aplicacao, icone
from comandos.lines import validador_so_numeros
from comandos.conversores import data_brasileiro
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import inspect
import os
from datetime import date
import traceback
from unidecode import unidecode


class TelaPreIncluir(QMainWindow, Ui_MainWindow):
    def __init__(self, id_user, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        self.id_usuario = id_user

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_cadastro.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Produto)

        self.lanca_numero_id()
        self.definir_data_emissao()
        self.manipula_dados_tabela()

        self.table_Produto.viewport().installEventFilter(self)

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)
        self.btn_Limpar.clicked.connect(self.limpa_tela)
        self.btn_Excluir.clicked.connect(self.excluir_cadastro)

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
            tabela = []
            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, obs, descricao, complementar, referencia, um, ncm, "
                           f"fornecedor, criacao, produto_id "
                           f"FROM PRODUTO_PRELIMINAR where produto_id IS NULL "
                           f"order by obs;")
            dados_banco = cursor.fetchall()

            if dados_banco:
                for i in dados_banco:
                    id_pre, obs, descr, compl, ref, um, ncm, forn, emissao, codigo = i

                    data_formatada = data_brasileiro(emissao)

                    dados = (data_formatada, id_pre, obs, descr, compl, ref, um, ncm, forn)
                    tabela.append(dados)

            if tabela:
                lanca_tabela(self.table_Produto, tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def lanca_numero_id(self):
        try:
            num_id = proximo_registro_autoincremento("PRODUTO_PRELIMINAR")
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
            self.line_Consulta.clear()

            self.line_Descricao.clear()
            self.line_NCM.clear()
            self.combo_UM.setCurrentText("")

            self.line_Referencia.clear()
            self.line_DescrCompl.clear()
            self.line_Fornecedor.clear()

            self.lanca_numero_id()
            self.definir_data_emissao()
            self.manipula_dados_tabela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tela(self):
        try:
            self.reiniciando_tela()

            self.plain_Obs.clear()

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
                cursor.execute(f"SELECT DISTINCT id, obs, descricao, complementar, referencia, um, ncm, "
                               f"fornecedor, criacao, produto_id "
                               f"FROM PRODUTO_PRELIMINAR "
                               f"WHERE descricao LIKE '%{palavra_maiuscula}%' "
                               f"ORDER BY obs;")
                palavra = cursor.fetchall()

                if not palavra:
                    self.mensagem_alerta(f'Não foi encontrado nenhum item com o nome:\n  "{palavra_consulta}"!')
                    self.line_Consulta.clear()
                else:
                    for i in palavra:
                        id_pre, obs, descr, compl, ref, um, ncm, forn, emissao, codigo = i

                        data_formatada = data_brasileiro(emissao)

                        dados = (data_formatada, id_pre, obs, descr, compl, ref, um, ncm, forn)
                        tabela_nova.append(dados)

                if tabela_nova:
                    lanca_tabela(self.table_Produto, tabela_nova)

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
                    source is self.table_Produto.viewport()):
                item = self.table_Produto.currentItem()

                extrai_recomendados = extrair_tabela(self.table_Produto)
                item_selecionado = extrai_recomendados[item.row()]

                data, id_pre, obs, descr, compl, ref, um, ncm, forn = item_selecionado

                self.line_Num.setText(id_pre)

                self.line_Descricao.setText(descr)
                self.line_DescrCompl.setText(compl)
                self.line_Referencia.setText(ref)
                self.line_NCM.setText(ncm)
                self.line_Fornecedor.setText(forn)

                self.plain_Obs.setPlainText(obs)

                item_count = self.combo_UM.count()
                for i in range(item_count):
                    item_text = self.combo_UM.itemText(i)
                    if um in item_text:
                        self.combo_UM.setCurrentText(item_text)

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
                cursor.execute(f"select * "
                               f"from PRODUTO_PRELIMINAR where id = {num_id};")
                conjunto = cursor.fetchall()

                if not conjunto:
                    self.mensagem_alerta(f'O cadastro do Produto Preliminar Nº {num_id} não existe!')
                else:
                    msg = f'Tem certeza que deseja excluir o Produto {nome}?'
                    if self.pergunta_confirmacao(msg):
                        cursor = conecta.cursor()
                        cursor.execute(f"DELETE FROM PRODUTO_PRELIMINAR WHERE id = {num_id};")

                        conecta.commit()

                        self.mensagem_alerta(f"Cadastro da Produto Preliminar {nome} foi excluído com Sucesso!")

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

            obs = self.plain_Obs.toPlainText()
            if not obs:
                obs_maiusculo = ""
            else:
                obs_maiusculo = obs.upper()

            descr = self.line_Descricao.text()
            descr_maiuscula = descr.upper()
            descr_sem_acentos = unidecode(descr_maiuscula)

            compl = self.line_DescrCompl.text()
            compl_maiuscula = compl.upper()
            compl_sem_acentos = unidecode(compl_maiuscula)

            ref = self.line_Referencia.text()
            ref_maiuscula = ref.upper()
            ref_sem_acentos = unidecode(ref_maiuscula)

            um = self.combo_UM.currentText()

            ncm  = self.line_NCM.text()

            fornecedor = self.line_Fornecedor.text()

            cursor = conecta.cursor()
            cursor.execute(f"select id, obs, descricao, complementar, referencia, um, ncm, "
                           f"fornecedor "
                           f"from PRODUTO_PRELIMINAR where id = {num_id};")
            conjuntos = cursor.fetchall()

            if conjuntos:
                num_id_b, obs_b, descr_b, compl_b, ref_b, um_b, ncm_b, forn_b = conjuntos[0]

                campos_atualizados = []
                if obs_maiusculo != obs_b:
                    campos_atualizados.append(f"obs = '{obs_maiusculo}'")
                if descr_sem_acentos != descr_b:
                    campos_atualizados.append(f"descricao = '{descr_sem_acentos}'")
                if compl_sem_acentos != compl_b:
                    campos_atualizados.append(f"complementar = '{compl_sem_acentos}'")
                if ref_sem_acentos != ref_b:
                    campos_atualizados.append(f"referencia = '{ref_sem_acentos}'")
                if um != um_b:
                    campos_atualizados.append(f"um = '{um}'")
                if ncm != ncm_b:
                    campos_atualizados.append(f"ncm = '{ncm}'")
                if fornecedor != forn_b:
                    campos_atualizados.append(f"fornecedor = '{fornecedor}'")

                if campos_atualizados:
                    msg = f'Deseja realmente atualizar o cadastro da Produto Preliminar?'
                    if self.pergunta_confirmacao(msg):
                        campos_update = ", ".join(campos_atualizados)

                        cursor.execute(f"UPDATE PRODUTO_PRELIMINAR SET {campos_update} "
                                       f"WHERE id = {num_id_b};")

                        conecta.commit()

                        msg = f'O cadastro da Produto Preliminar {descr_sem_acentos} foi atualizado com sucesso!'
                        self.mensagem_alerta(msg)

            else:
                msg = f'Deseja realmente cadastrar este Produto?'
                if self.pergunta_confirmacao(msg):
                    cursor = conecta.cursor()
                    cursor.execute(f"Insert into PRODUTO_PRELIMINAR (OBS, DESCRICAO, COMPLEMENTAR, "
                                   f"REFERENCIA, UM, NCM, FORNECEDOR, ENTREGUE, USUARIO_ID) "
                                   f"values ('{obs_maiusculo}', '{descr_sem_acentos}', "
                                   f"'{compl_sem_acentos}', '{ref_sem_acentos}', "
                                   f"'{um}', '{ncm}', '{fornecedor}', 'N', {self.id_usuario});")

                    conecta.commit()

                    msg = f'O cadastro dO Produto Preliminar {descr_sem_acentos} foi criado com sucesso!'
                    self.mensagem_alerta(msg)

            self.reiniciando_tela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaPreIncluir("1")
    tela.show()
    qt.exec_()

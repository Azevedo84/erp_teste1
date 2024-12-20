import sys
from banco_dados.conexao import conecta
from forms.tela_estrut_versao import *
from banco_dados.controle_erros import grava_erro_banco
from banco_dados.bc_consultas import definir_proximo_registro
from comandos.tabelas import lanca_tabela, layout_cabec_tab, extrair_tabela
from comandos.telas import tamanho_aplicacao, icone
from comandos.lines import validador_so_numeros
from comandos.conversores import timestamp_brasileiro
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtCore import QDate
import inspect
import os
from datetime import date
import traceback
from unidecode import unidecode


class TelaCadastroVersoes(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_estrutura.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Lista)

        self.processando = False

        self.lanca_numero_id()
        self.definir_data_emissao()

        validador_so_numeros(self.line_Num)
        self.line_Num.setReadOnly(True)

        validador_so_numeros(self.line_Codigo)

        self.table_Lista.viewport().installEventFilter(self)

        self.btn_Limpar.clicked.connect(self.limpa_tudo)

        self.line_Codigo.editingFinished.connect(self.verifica_line_codigo_acabado)
        self.btn_Salvar.clicked.connect(self.verifica_salvamento)

        self.line_Codigo.setFocus()

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

    def definir_bloqueios(self):
        try:
            self.line_Descricao.setReadOnly(True)
            self.line_Referencia.setReadOnly(True)
            self.line_UM.setReadOnly(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_numero_id(self):
        try:
            definir_proximo_registro(self.line_Num, "id", "estrutura")

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

    def verifica_line_codigo_acabado(self):
        if not self.processando:
            try:
                self.processando = True

                self.limpa_dados_produto_estrutura()

                codigo_produto = self.line_Codigo.text()

                if not codigo_produto:
                    self.mensagem_alerta('O campo "Código" não pode estar vazio!')
                    self.limpa_dados_produto_estrutura()
                    self.limpa_tabela()
                elif int(codigo_produto) == 0:
                    self.mensagem_alerta('O campo "Código" não pode ser "0"!')
                    self.limpa_dados_produto_estrutura()
                    self.limpa_tabela()
                else:
                    self.verifica_sql_acabado()

                    self.table_Lista.setFocus()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_sql_acabado(self):
        try:
            codigo_produto = self.line_Codigo.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade, conjunto, quantidade "
                           f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cursor.fetchall()
            if not detalhes_produto:
                self.mensagem_alerta('Este código de produto não existe!')
                self.limpa_dados_produto_estrutura()
                self.limpa_tabela()
                self.line_Codigo.clear()
            else:
                conjunto = detalhes_produto[0][3]
                if conjunto == 10:
                    self.lanca_dados_acabado()
                else:
                    self.mensagem_alerta('Este produto não tem o conjunto classificado como "Produtos Acabados"!')
                    self.limpa_dados_produto_estrutura()
                    self.limpa_tabela()
                    self.line_Codigo.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_acabado(self):
        try:
            codigo_produto = self.line_Codigo.text()
            cur = conecta.cursor()
            cur.execute(f"SELECT prod.descricao, COALESCE(tip.tipomaterial, '') as tipus, "
                        f"COALESCE(prod.obs, '') as ref, prod.unidade, "
                        f"COALESCE(prod.ncm, '') as ncm, COALESCE(prod.obs2, '') as obs "
                        f"FROM produto as prod "
                        f"LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id "
                        f"where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            descr, tipo, ref, um, ncm, obs = detalhes_produto[0]

            self.line_Descricao.setText(descr)
            self.line_Referencia.setText(ref)
            self.line_UM.setText(um)

            self.lanca_versoes()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_versoes(self):
        try:
            nova_tabela = []
            codigo_produto = self.line_Codigo.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {codigo_produto};")
            select_prod = cursor.fetchall()
            id_pai, cod, id_versao = select_prod[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, num_versao, data_versao, obs, data_criacao "
                           f"from estrutura "
                           f"where id_produto = {id_pai} order by data_versao;")
            tabela_versoes = cursor.fetchall()

            if tabela_versoes:
                for i in tabela_versoes:
                    id_estrut, num_versao, data, obs, criacao = i

                    data_versao = data.strftime("%d/%m/%Y")

                    criacao_formatada = timestamp_brasileiro(criacao)

                    if id_versao == id_estrut:
                        status_txt = "ATIVO"
                    else:
                        status_txt = "INATIVO"

                    dados = (id_estrut, num_versao, data_versao, status_txt, obs, criacao_formatada)
                    nova_tabela.append(dados)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT MAX(num_versao) from estrutura where id_produto = {id_pai};")
            resultado = cursor.fetchone()

            if resultado[0] is not None:
                proximo_num_versao = resultado[0] + 1
            else:
                proximo_num_versao = 0

            self.line_Num_Versao.setText(str(proximo_num_versao))

            status_count = self.combo_Status.count()
            for status_ in range(status_count):
                status_text = self.combo_Status.itemText(status_)
                if status_text == "ATIVO":
                    self.combo_Status.setCurrentText(status_text)
                    break

            if self.line_Num_Versao.text() == "0":
                self.plain_Obs.setPlainText("PRIMEIRA VERSAO CRIADA")

            if nova_tabela:
                lanca_tabela(self.table_Lista, nova_tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tudo(self):
        self.limpa_tabela()
        self.limpa_dados_produto_estrutura()
        self.line_Num_Versao.clear()

        status_count = self.combo_Status.count()
        for status_ in range(status_count):
            status_text = self.combo_Status.itemText(status_)
            if status_text == "":
                self.combo_Status.setCurrentText(status_text)
                break

        self.plain_Obs.clear()

        self.lanca_numero_id()
        self.definir_data_emissao()

        self.line_Codigo.clear()
        self.line_Codigo.setFocus()

    def limpa_dados_produto_estrutura(self):
        try:
            self.line_Descricao.clear()
            self.line_Referencia.clear()
            self.line_UM.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tabela(self):
        try:
            self.table_Lista.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def eventFilter(self, source, event):
        try:
            if (event.type() == QtCore.QEvent.MouseButtonDblClick and
                    event.buttons() == QtCore.Qt.LeftButton and
                    source is self.table_Lista.viewport()):
                item = self.table_Lista.currentItem()

                extrai_recomendados = extrair_tabela(self.table_Lista)
                item_selecionado = extrai_recomendados[item.row()]

                id_estrut, num_versao, data, status, obs, criacao = item_selecionado

                self.line_Num.setText(id_estrut)

                self.line_Num_Versao.setText(num_versao)

                date_obj = QDate.fromString(data, "dd/MM/yyyy")
                self.date_Emissao.setDate(date_obj)

                if status:
                    status_count = self.combo_Status.count()
                    for status_ in range(status_count):
                        status_text = self.combo_Status.itemText(status_)
                        if status == status_text:
                            self.combo_Status.setCurrentText(status_text)

                self.plain_Obs.setPlainText(obs)

                self.definir_bloqueios()

            return super(QMainWindow, self).eventFilter(source, event)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_salvamento(self):
        try:
            num_id = self.line_Num.text()
            num_versao = self.line_Num_Versao.text()
            codigo_produto = self.line_Codigo.text()
            status = self.combo_Status.currentText()
            obs = self.plain_Obs.toPlainText()

            if not num_id:
                self.mensagem_alerta('O campo "Código" não pode estar vazio!')
                self.line_Num.setFocus()
            elif num_id == "0":
                self.mensagem_alerta('O campo "Código" não pode ser "0"!')
                self.line_Num.clear()
                self.line_Num.setFocus()
            elif not num_versao:
                self.mensagem_alerta('O campo "Nº Versão" não pode estar vazio!')
                self.line_Num_Versao.setFocus()
            elif not codigo_produto:
                self.mensagem_alerta('O campo "Código" do Produto não pode estar vazio!   ')
                self.line_Codigo.setFocus()
            elif codigo_produto == "0":
                self.mensagem_alerta('O campo "Código" do Produto não pode ser "0"!   ')
                self.line_Codigo.clear()
                self.line_Codigo.setFocus()
            elif not status:
                self.mensagem_alerta('O campo "Status" não pode estar vazio!   ')
                self.combo_Status.setFocus()
            elif not obs:
                self.mensagem_alerta('Utilize o campo "Observação:" para especificar as alterações do projeto, '
                                     'se houver.')
                self.plain_Obs.setFocus()
            else:
                self.salvar_dados()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_dados(self):
        try:
            num_id = self.line_Num.text()
            num_versao = self.line_Num_Versao.text()

            codigo_produto = self.line_Codigo.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {codigo_produto};")
            select_prod = cursor.fetchall()
            id_pai, cod, id_versao = select_prod[0]

            status = self.combo_Status.currentText()

            if int(num_id) == id_versao:
                if status == "INATIVO":
                    self.mensagem_alerta("Este é a versão ativa atualmente e não pode ser inativada!")
            else:
                cursor.execute(f"UPDATE produto SET id_versao = {num_id} "
                               f"WHERE id = {id_pai};")

                conecta.commit()

                emissao = self.date_Emissao.date()
                data_emi = emissao.toString("yyyy-MM-dd")

                obs = self.plain_Obs.toPlainText()
                obs_maiuscula = obs.upper()
                obs_sem_acentos = unidecode(obs_maiuscula)

                cursor = conecta.cursor()
                cursor.execute(f"select id, id_produto, num_versao, data_versao, obs, data_criacao "
                               f"from estrutura where id = {num_id};")
                estrutura = cursor.fetchall()

                if estrutura:
                    num_id_b, id_produto_b, num_versao_b, data_versao_b, obs_b, criacao_b = estrutura[0]

                    if id_pai == id_produto_b and int(num_versao) == num_versao_b:
                        campos_atualizados = []
                        if obs_sem_acentos != obs_b:
                            campos_atualizados.append(f"obs = '{obs_sem_acentos}'")
                        if data_emi != data_versao_b:
                            campos_atualizados.append(f"data_versao = '{data_emi}'")

                        if campos_atualizados:
                            msg = f'Deseja realmente atualizar os dados da Versão {num_versao} do ' \
                                  f'Produto {codigo_produto}?'
                            if self.pergunta_confirmacao(msg):
                                campos_update = ", ".join(campos_atualizados)

                                cursor.execute(f"UPDATE estrutura SET {campos_update} "
                                               f"WHERE id = {num_id_b};")

                                conecta.commit()

                                msg = f'O cadastro da Estrutura do Produto {codigo_produto} foi atualizado com sucesso!'
                                self.mensagem_alerta(msg)
                    else:
                        self.mensagem_alerta("O Código do produto e o número da versão não podem ser alterados"
                                             " no controle de versões!")

                else:
                    msg = f'Deseja realmente cadastrar a Versão {num_versao} do Produto {codigo_produto}?'
                    if self.pergunta_confirmacao(msg):
                        cursor = conecta.cursor()
                        cursor.execute(f"Insert into estrutura (id, id_produto, num_versao, data_versao, obs) "
                                       f"values (GEN_ID(GEN_ESTRUTURA_ID,1), {id_pai}, '{num_versao}', '{data_emi}', "
                                       f"'{obs_sem_acentos}');")

                        conecta.commit()

                        msg = f'O cadastro da Versão {num_versao} do Produto {codigo_produto} foi criado com sucesso!'
                        self.mensagem_alerta(msg)

            self.limpa_tudo()
            self.line_Codigo.clear()
            self.line_Codigo.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaCadastroVersoes()
    tela.show()
    qt.exec_()

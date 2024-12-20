import sys
from banco_dados.conexao import conecta
from forms.tela_op_consumir_v2 import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.cores import cor_cinza_claro, cor_vermelho, cor_branco
from comandos.telas import tamanho_aplicacao, icone
from PyQt5.QtWidgets import QApplication, QMainWindow, QShortcut, QMessageBox, QHeaderView
from PyQt5.QtGui import QColor, QFont, QKeySequence
from PyQt5.QtCore import Qt
from datetime import date, datetime
import inspect
import os
import traceback


class TelaOpConsumirV2(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        current_dir = os.path.dirname(__file__)
        icon_path = os.path.join(current_dir, '..', 'arquivos', 'imagens tela', 'op_consumir.png')
        self.label_imagem.setPixmap(QtGui.QPixmap(icon_path))

        icone(self, "menu_producao.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Estrutura)
        layout_cabec_tab(self.table_ConsumoOS)

        self.tab_shortcut = QShortcut(QKeySequence(Qt.Key_Tab), self)
        self.tab_shortcut.activated.connect(self.manipula_tab)

        self.line_Num_OP.returnPressed.connect(lambda: self.verifica_linenumero_os())

        self.btn_ConsumoTodos.clicked.connect(self.lanca_todos)

        self.line_Cod_Manu.returnPressed.connect(lambda: self.verifica_line_cod_manu())
        self.btn_Consome_Manu.clicked.connect(self.verifica_line_qtde_manu)
        self.line_Qtde_Manu.returnPressed.connect(lambda: self.verifica_line_qtde_manu())

        self.line_Cod_Subs.returnPressed.connect(lambda: self.verifica_line_cod_subs())
        self.btn_Consome_Subs.clicked.connect(self.verifica_line_qtde_subs)
        self.line_Qtde_Subs.returnPressed.connect(lambda: self.verifica_line_qtde_subs())

        self.btn_Define_Substituto.clicked.connect(self.exibe_substituto)

        self.btn_Excluir_Item.clicked.connect(self.excluir_item)

        self.btn_Salvar.clicked.connect(self.verifica_salvamento2)

        self.table_Estrutura.viewport().installEventFilter(self)
        self.line_Cod_Manu.setEnabled(False)
        self.line_Qtde_Manu.setEnabled(False)
        self.date_Consumo.setEnabled(False)
        self.btn_Excluir_Item.setEnabled(False)
        self.btn_Consome_Manu.setEnabled(False)
        self.btn_ConsumoTodos.setEnabled(False)
        self.btn_Salvar.setEnabled(False)

        validator = QtGui.QDoubleValidator(0, 9999999.000, 3, self.line_Qtde_Manu)
        locale = QtCore.QLocale("pt_BR")
        validator.setLocale(locale)
        self.line_Qtde_Manu.setValidator(validator)

        validator = QtGui.QRegExpValidator(QtCore.QRegExp(r'\d+'), self.line_Num_OP)
        self.line_Num_OP.setValidator(validator)

        validator = QtGui.QRegExpValidator(QtCore.QRegExp(r'\d+'), self.line_Cod_Manu)
        self.line_Cod_Manu.setValidator(validator)

        self.definir_bloqueios()

        self.line_Num_OP.setFocus()

        data_hoje = date.today()
        self.date_Consumo.setDate(data_hoje)

        self.remove_layout_substituto()
        self.btn_Define_Substituto.setHidden(True)

        self.qtde_vezes_select = 0

        self.op_encerra = []

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

    def definir_bloqueios(self):
        try:
            self.line_Codigo.setReadOnly(True)

            self.line_Descricao.setReadOnly(True)
            self.line_Referencia.setReadOnly(True)
            self.line_UM.setReadOnly(True)
            self.line_Compl.setReadOnly(True)

            self.line_Saldo.setReadOnly(True)
            self.line_Qtde.setReadOnly(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_tab(self):
        try:
            if self.line_Num_OP.hasFocus():
                self.verifica_linenumero_os()

            elif self.line_Cod_Manu.hasFocus():
                self.verifica_line_cod_manu()

            elif self.line_Qtde_Manu.hasFocus():
                self.verifica_line_qtde_manu()

            elif self.line_Cod_Subs.hasFocus():
                self.verifica_line_cod_subs()

            elif self.line_Qtde_Subs.hasFocus():
                self.verifica_line_qtde_subs()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def remove_layout_substituto(self):
        try:
            self.widget_Subs.setHidden(True)

            self.btn_Consome_Subs.setHidden(True)
            self.label_imagem.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def adiciona_layout_substituto(self):
        try:
            self.widget_Subs.setHidden(False)

            self.btn_Consome_Subs.setEnabled(False)
            self.btn_Consome_Manu.setHidden(True)
            self.line_Qtde_Subs.setEnabled(False)
            self.label_imagem.setHidden(False)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_linenumero_os(self):
        try:
            self.limpa_subs()
            self.line_Cod_Manu.clear()
            self.line_Cod_Manu.setFocus()

            self.line_Ref_Manu.clear()
            self.line_Saldo_Manu.clear()
            self.line_Descricao_Manu.clear()
            self.line_Compl_Manu.clear()
            self.line_Local_Manu.clear()
            self.line_UM_Manu.clear()
            self.line_Qtde_Manu.clear()
            numero_os_line = self.line_Num_OP.text()
            if len(numero_os_line) == 0:
                self.mensagem_alerta('O campo "Nº OP" não pode estar vazio')
                self.reiniciar()
            elif int(numero_os_line) == 0:
                self.mensagem_alerta('O campo "Nº OP" não pode ser "0"')
                self.reiniciar()
            else:
                self.verifica_sql_os()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_sql_os(self):
        try:
            numero_os_line = self.line_Num_OP.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT numero, datainicial, status, produto, quantidade "
                           f"FROM ordemservico where numero = {numero_os_line};")
            extrair_dados = cursor.fetchall()
            if not extrair_dados:
                self.mensagem_alerta('Este número de "OP" não existe!')
                self.reiniciar()
            else:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT numero, datainicial, status, produto, quantidade "
                               f"FROM ordemservico where numero = {numero_os_line} AND status = 'A';")
                select_status = cursor.fetchall()

                if not select_status:
                    self.mensagem_alerta('Esta Ordem de Produção está encerrada!')
                    self.reiniciar()
                else:
                    self.verifica_vinculo_materia()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_vinculo_materia(self):
        try:
            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut = self.dados_os()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT codigo, id_estrut_prod, qtde_estrut_prod "
                           f"FROM produtoos where numero = {numero_os};")
            itens_os = cursor.fetchall()
            verifica_cadastro = 0
            for item in itens_os:
                codigo, id_materia, qtde_materia = item
                if not id_materia and not qtde_materia:
                    verifica_cadastro = verifica_cadastro + 1

            if verifica_cadastro > 0:
                self.mensagem_alerta('O material consumido não está vinculado com a estrutura!')
                self.reiniciar()
            else:
                self.verifica_dados_os()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_dados_os(self):
        try:
            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut = self.dados_os()

            print(id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"((SELECT quantidade FROM ordemservico where numero = {numero_os}) * "
                           f"(estprod.quantidade)) AS Qtde, "
                           f"COALESCE(prod.localizacao, ''), prod.quantidade "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id_estrutura = {id_estrut} ORDER BY prod.descricao;")
            itens_select_estrut = cursor.fetchall()
            if not itens_select_estrut:
                self.mensagem_alerta('Este material não tem estrutura cadastrada!')
                self.reiniciar()
            elif status_os != "A":
                self.mensagem_alerta('Esta "OP" está encerrada!')
                self.reiniciar()
            elif data_emissao is None:
                self.mensagem_alerta('Esta "OP" está sem data de emissão!')
                self.reiniciar()
            elif produto_os is None:
                self.mensagem_alerta('Esta "OP" está sem código de produto!')
                self.reiniciar()
            elif qtde_os is None:
                self.mensagem_alerta('A quantidade da "OP" deve ser maior que "0"!')
                self.reiniciar()
            elif numero_os is None:
                self.mensagem_alerta('O número da "OP" deve ser maior que "0"!')
                self.reiniciar()
            else:
                self.lanca_dados_os()
                self.separar_dados_select()
                self.pintar_tabelas()
                self.btn_Salvar.setEnabled(True)
                self.line_Cod_Manu.setEnabled(True)
                self.line_Cod_Manu.setFocus()
                self.btn_ConsumoTodos.setEnabled(True)
                self.btn_Excluir_Item.setEnabled(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_os(self):
        try:
            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut = self.dados_os()
            cur = conecta.cursor()
            cur.execute(f"SELECT codigo, descricao, COALESCE(obs, ' ') as obs, "
                        f"COALESCE(descricaocomplementar, ' ') as compl, unidade, id_versao "
                        f"FROM produto where id = {produto_os};")
            detalhes_produtos = cur.fetchall()

            codigo_id, descricao_id, referencia_id, complementar_id, unidade_id, id_versao = detalhes_produtos[0]
            self.line_Codigo.setText(codigo_id)
            self.line_Descricao.setText(descricao_id)
            self.line_Referencia.setText(referencia_id)
            self.line_Compl.setText(complementar_id)
            self.line_UM.setText(unidade_id)
            numero = str(qtde_os).replace('.', ',')
            self.line_Qtde.setText(numero)

            data_em_texto = '{}/{}/{}'.format(data_emissao.day, data_emissao.month, data_emissao.year)
            self.label_Emissao.setText(data_em_texto)
            self.btn_Define_Substituto.setHidden(False)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, num_versao, data_versao, obs, data_criacao "
                           f"from estrutura "
                           f"where id = {id_versao};")
            tabela_versoes = cursor.fetchall()

            id_estrut, num_versao, data, obs, criacao = tabela_versoes[0]

            data_versao = data.strftime("%d/%m/%Y")

            if id_versao == id_estrut:
                status_txt = "ATIVO"
            else:
                status_txt = "INATIVO"

            msg = f"VERSÃO: {num_versao} - {data_versao} - {status_txt}"
            self.line_Versao.setText(msg)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def dados_os(self):
        try:
            numero_os_line = self.line_Num_OP.text()
            cur = conecta.cursor()
            cur.execute(f"SELECT id, numero, datainicial, status, produto, quantidade, id_estrutura "
                        f"FROM ordemservico where numero = {numero_os_line};")
            extrair_dados = cur.fetchall()
            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut = extrair_dados[0]

            return id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def select_mistura(self):
        try:
            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut = self.dados_os()

            dados_para_tabela = []
            campo_br = ""

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"((SELECT quantidade FROM ordemservico where numero = {numero_os}) * "
                           f"(estprod.quantidade)) AS Qtde, "
                           f"COALESCE(prod.localizacao, ''), prod.quantidade "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id_estrutura = {id_estrut} ORDER BY prod.descricao;")
            select_estrut = cursor.fetchall()

            for dados_estrut in select_estrut:
                id_mat_e, cod_e, descr_e, ref_e, um_e, qtde_e, local_e, saldo_e = dados_estrut

                cursor = conecta.cursor()
                cursor.execute(f"SELECT max(estprod.id), max(prod.codigo), max(prod.descricao), "
                               f"sum(prodser.QTDE_ESTRUT_PROD) as total "
                               f"FROM estrutura_produto as estprod "
                               f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                               f"INNER JOIN produtoos as prodser ON estprod.id = prodser.id_estrut_prod "
                               f"where prodser.numero = {numero_os} and estprod.id = {id_mat_e} "
                               f"group by prodser.id_estrut_prod;")
                select_os_resumo = cursor.fetchall()

                if not select_os_resumo:
                    dados0 = (id_mat_e, cod_e, descr_e, ref_e, um_e, qtde_e, local_e, saldo_e,
                              campo_br, campo_br, campo_br, campo_br, campo_br, campo_br)
                    dados_para_tabela.append(dados0)

                for dados_res in select_os_resumo:
                    id_mat_sum, cod_sum, descr_sum, qtde_sum = dados_res

                    sobras = qtde_e - qtde_sum
                    if sobras > 0:
                        dados1 = (id_mat_e, cod_e, descr_e, ref_e, um_e, sobras, local_e, saldo_e,
                                  campo_br, campo_br, campo_br, campo_br, campo_br, campo_br)
                        dados_para_tabela.append(dados1)

                    cursor = conecta.cursor()
                    cursor.execute(f"select prodser.id_estrut_prod, "
                                   f"COALESCE((extract(day from prodser.data)||'/'||"
                                   f"extract(month from prodser.data)||'/'||"
                                   f"extract(year from prodser.data)), '') AS DATA, prod.codigo, prod.descricao, "
                                   f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                                   f"prodser.quantidade, prodser.qtde_estrut_prod "
                                   f"from produtoos as prodser "
                                   f"INNER JOIN produto as prod ON prodser.produto = prod.id "
                                   f"where prodser.numero = {numero_os} and prodser.id_estrut_prod = {id_mat_e};")
                    select_os = cursor.fetchall()

                    for dados_os in select_os:
                        id_mat_os, data_os, cod_os, descr_os, ref_os, um_os, qtde_os, qtde_mat_os = dados_os

                        dados2 = (id_mat_e, cod_e, descr_e, ref_e, um_e, qtde_mat_os, local_e, saldo_e,
                                  data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                        dados_para_tabela.append(dados2)

            return dados_para_tabela

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def separar_dados_select(self):
        try:
            itens_manipula_total = self.select_mistura()

            self.qtde_vezes_select = self.qtde_vezes_select + 1

            tabela_estrutura = []
            tabela_consumo_os = []

            for itens in itens_manipula_total:
                id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo, data_os, cod_os, descr_os, \
                ref_os, um_os, qtde_os = itens

                qtde_est_str = str(qtde_est)
                qtde_est_float = float(qtde_est_str)
                qtde_est_red = "%.3f" % qtde_est_float

                if saldo == "":
                    saldo_red = saldo
                else:
                    saldo_str = str(saldo)
                    saldo_float = float(saldo_str)
                    saldo_red = "%.3f" % saldo_float

                if qtde_os == "":
                    qtde_os_red = qtde_os
                else:
                    qtde_os_str = str(qtde_os)
                    qtde_os_float = float(qtde_os_str)
                    qtde_os_red = "%.3f" % qtde_os_float

                lista_est = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est_red, local, saldo_red)
                tabela_estrutura.append(lista_est)

                lista_os = (id_mat, data_os, cod_os, descr_os, ref_os, um_os, qtde_os_red)
                tabela_consumo_os.append(lista_os)

            lanca_tabela(self.table_ConsumoOS, tabela_consumo_os)
            lanca_tabela(self.table_Estrutura, tabela_estrutura)
            self.table_ConsumoOS.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
            self.table_Estrutura.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def jutando_tabelas_extraidas(self):
        try:
            estrutura = extrair_tabela(self.table_Estrutura)
            consumo_os = extrair_tabela(self.table_ConsumoOS)

            linhas_est = len(estrutura)
            extrai_total = []

            for linha_est in range(linhas_est):
                id_mat_est, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo = estrutura[linha_est]

                id_mat_os, data_os, cod_os, descr_os, ref_os, um_os, qtde_os = consumo_os[linha_est]

                extrai_todos = (id_mat_est, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo,
                                data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                extrai_total.append(extrai_todos)
            return extrai_total

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pintar_tabelas(self):
        try:
            extrai_total = self.jutando_tabelas_extraidas()

            testinho = 0

            for itens in extrai_total:
                id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo, \
                data_os, cod_os, descr_os, ref_os, um_os, qtde_os = itens
                qtde_est_float = float(qtde_est)
                saldo_float = float(saldo)

                if not cod_os and saldo_float < qtde_est_float:
                    testinho = testinho + 1
                    testinho2 = testinho - 1
                    self.table_Estrutura.item(testinho2, 0).setBackground(QColor(cor_vermelho))
                    self.table_Estrutura.item(testinho2, 1).setBackground(QColor(cor_vermelho))
                    self.table_Estrutura.item(testinho2, 2).setBackground(QColor(cor_vermelho))
                    self.table_Estrutura.item(testinho2, 3).setBackground(QColor(cor_vermelho))
                    self.table_Estrutura.item(testinho2, 4).setBackground(QColor(cor_vermelho))
                    self.table_Estrutura.item(testinho2, 5).setBackground(QColor(cor_vermelho))
                    self.table_Estrutura.item(testinho2, 6).setBackground(QColor(cor_vermelho))
                    self.table_Estrutura.item(testinho2, 7).setBackground(QColor(cor_vermelho))

                    self.table_Estrutura.item(testinho2, 0).setForeground(QColor(cor_branco))

                    font = QFont()
                    font.setBold(True)
                    self.table_Estrutura.item(testinho2, 0).setFont(font)
                    self.table_Estrutura.item(testinho2, 0).setForeground(QColor(cor_branco))
                    self.table_Estrutura.item(testinho2, 1).setFont(font)
                    self.table_Estrutura.item(testinho2, 1).setForeground(QColor(cor_branco))
                    self.table_Estrutura.item(testinho2, 2).setFont(font)
                    self.table_Estrutura.item(testinho2, 2).setForeground(QColor(cor_branco))
                    self.table_Estrutura.item(testinho2, 3).setFont(font)
                    self.table_Estrutura.item(testinho2, 3).setForeground(QColor(cor_branco))
                    self.table_Estrutura.item(testinho2, 4).setFont(font)
                    self.table_Estrutura.item(testinho2, 4).setForeground(QColor(cor_branco))
                    self.table_Estrutura.item(testinho2, 5).setFont(font)
                    self.table_Estrutura.item(testinho2, 5).setForeground(QColor(cor_branco))
                    self.table_Estrutura.item(testinho2, 6).setFont(font)
                    self.table_Estrutura.item(testinho2, 6).setForeground(QColor(cor_branco))
                    self.table_Estrutura.item(testinho2, 7).setFont(font)
                    self.table_Estrutura.item(testinho2, 7).setForeground(QColor(cor_branco))

                elif not cod_os:
                    testinho = testinho + 1

                else:
                    testinho = testinho + 1
                    testinho2 = testinho - 1
                    self.table_Estrutura.item(testinho2, 0).setBackground(QColor(cor_cinza_claro))
                    self.table_Estrutura.item(testinho2, 1).setBackground(QColor(cor_cinza_claro))
                    self.table_Estrutura.item(testinho2, 2).setBackground(QColor(cor_cinza_claro))
                    self.table_Estrutura.item(testinho2, 3).setBackground(QColor(cor_cinza_claro))
                    self.table_Estrutura.item(testinho2, 4).setBackground(QColor(cor_cinza_claro))
                    self.table_Estrutura.item(testinho2, 5).setBackground(QColor(cor_cinza_claro))
                    self.table_Estrutura.item(testinho2, 6).setBackground(QColor(cor_cinza_claro))
                    self.table_Estrutura.item(testinho2, 7).setBackground(QColor(cor_cinza_claro))

                    self.table_ConsumoOS.item(testinho2, 0).setBackground(QColor(cor_cinza_claro))
                    self.table_ConsumoOS.item(testinho2, 1).setBackground(QColor(cor_cinza_claro))
                    self.table_ConsumoOS.item(testinho2, 2).setBackground(QColor(cor_cinza_claro))
                    self.table_ConsumoOS.item(testinho2, 3).setBackground(QColor(cor_cinza_claro))
                    self.table_ConsumoOS.item(testinho2, 4).setBackground(QColor(cor_cinza_claro))
                    self.table_ConsumoOS.item(testinho2, 5).setBackground(QColor(cor_cinza_claro))
                    self.table_ConsumoOS.item(testinho2, 6).setBackground(QColor(cor_cinza_claro))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consumirtodos(self):
        try:
            data_hoje = date.today()
            data_string = data_hoje.strftime("%d/%m/%Y")

            extrai_total = self.jutando_tabelas_extraidas()

            tabela_estrutura = []
            tabela_consumo_os = []

            for itens in extrai_total:
                id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo, \
                data_os, cod_os, descr_os, ref_os, um_os, qtde_os = itens
                qtde_est_float = float(qtde_est)
                saldo_float = float(saldo)

                if not cod_os and saldo_float < qtde_est_float:
                    lista_est = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo)
                    tabela_estrutura.append(lista_est)

                    lista_os = (id_mat, data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                    tabela_consumo_os.append(lista_os)
                elif not cod_os:
                    lista_est = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo)
                    tabela_estrutura.append(lista_est)

                    lista_os = (id_mat, data_string, cod_est, descr_est, ref_est, um_est, qtde_est)
                    tabela_consumo_os.append(lista_os)
                else:
                    lista_est = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo)
                    tabela_estrutura.append(lista_est)

                    lista_os = (id_mat, data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                    tabela_consumo_os.append(lista_os)
            return tabela_estrutura, tabela_consumo_os

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_todos(self):
        try:
            tabela_estrutura, tabela_consumo_os = self.consumirtodos()

            lanca_tabela(self.table_ConsumoOS, tabela_consumo_os)
            lanca_tabela(self.table_Estrutura, tabela_estrutura)
            self.pintar_tabelas()
            self.btn_Salvar.setEnabled(True)
            self.table_ConsumoOS.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
            self.table_Estrutura.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def eventFilter(self, source, event):
        try:
            if (event.type() == QtCore.QEvent.MouseButtonDblClick and
                    event.buttons() == QtCore.Qt.LeftButton and
                    source is self.table_Estrutura.viewport()):
                item = self.table_Estrutura.currentItem()

                if item is None:
                    return super(QMainWindow, self).eventFilter(source, event)

                data_hoje = date.today()
                data_string = data_hoje.strftime("%d/%m/%Y")

                extrai_total = self.jutando_tabelas_extraidas()

                lista_id_mat = [ide[0] for ide in extrai_total]
                id_mat_selecao = lista_id_mat[item.row()]

                tabela_estrutura = []
                tabela_consumo_os = []

                item_sem_saldo = 0
                item_ja_lancado = 0

                for itens in extrai_total:
                    id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo, \
                    data_os, cod_os, descr_os, ref_os, um_os, qtde_os = itens
                    qtde_est_float = float(qtde_est)
                    saldo_float = float(saldo)

                    if not cod_os and saldo_float < qtde_est_float and id_mat_selecao == id_mat:
                        item_sem_saldo += 1
                        lista_est = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo)
                        tabela_estrutura.append(lista_est)

                        lista_os = (id_mat, data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                        tabela_consumo_os.append(lista_os)

                    elif cod_os != "" and id_mat_selecao == id_mat:
                        item_ja_lancado += 1
                        lista_est = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo)
                        tabela_estrutura.append(lista_est)

                        lista_os = (id_mat, data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                        tabela_consumo_os.append(lista_os)

                    else:
                        if id_mat_selecao == id_mat:
                            lista_est = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo)
                            tabela_estrutura.append(lista_est)

                            lista_os = (id_mat, data_string, cod_est, descr_est, ref_est, um_est, qtde_est)
                            tabela_consumo_os.append(lista_os)
                        else:
                            lista_est = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo)
                            tabela_estrutura.append(lista_est)

                            lista_os = (id_mat, data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                            tabela_consumo_os.append(lista_os)

                if item_sem_saldo > 0:
                    self.mensagem_alerta(f'Este material não tem saldo suficiente!')

                lanca_tabela(self.table_ConsumoOS, tabela_consumo_os)
                lanca_tabela(self.table_Estrutura, tabela_estrutura)
                self.pintar_tabelas()
                self.btn_Salvar.setEnabled(True)
                self.table_ConsumoOS.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                self.table_Estrutura.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

            return super(QMainWindow, self).eventFilter(source, event)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_cod_manu(self):
        try:
            self.remove_layout_substituto()
            self.btn_Consome_Manu.setHidden(False)
            self.line_Qtde_Subs.setEnabled(True)

            codigo_produto = self.line_Cod_Manu.text()
            if len(codigo_produto) == 0:
                self.mensagem_alerta('O campo "Código" não pode estar vazio')
                self.line_Cod_Manu.clear()
            elif int(codigo_produto) == 0:
                self.mensagem_alerta('O campo "Código" não pode ser "0"')
                self.line_Cod_Manu.clear()
            else:
                self.verifica_sql_produto_manu()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_sql_produto_manu(self):
        try:
            codigo_produto = self.line_Cod_Manu.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT descricao, COALESCE(obs, '') as obs, unidade, "
                           f"COALESCE(localizacao, ''), quantidade "
                           f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cursor.fetchall()
            if not detalhes_produto:
                self.mensagem_alerta('Este código de produto não existe!')
                self.line_Cod_Manu.clear()
            else:
                self.lanca_dados_produtomanu()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_produtomanu(self):
        try:
            self.date_Consumo.setEnabled(True)
            codigo_produto = self.line_Cod_Manu.text()
            cur = conecta.cursor()
            cur.execute(f"SELECT descricao, COALESCE(obs, '') as obs, unidade, COALESCE(localizacao, ''), quantidade "
                        f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            descricao_id, referencia_id, unidade_id, local_id, quantidade_id = detalhes_produto[0]

            quantidade_id_float = float(quantidade_id)
            numero = str(quantidade_id).replace('.', ',')

            if quantidade_id_float < 0:
                self.mensagem_alerta(f'Este produto está com saldo negativo!\n'
                                     f'Saldo Total = {quantidade_id_float}')
                self.line_Cod_Manu.clear()
            else:
                self.verifica_estrutura_produto_manu()
                self.line_Descricao_Manu.setText(descricao_id)
                self.line_Ref_Manu.setText(referencia_id)
                self.line_UM_Manu.setText(unidade_id)
                self.line_Local_Manu.setText(local_id)
                self.line_Saldo_Manu.setText(numero)
                self.line_Qtde_Manu.setEnabled(True)
                self.line_Qtde_Manu.setFocus()
                self.btn_Consome_Manu.setEnabled(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_estrutura_produto_manu(self):
        try:
            codigo_produto = self.line_Cod_Manu.text()
            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut = self.dados_os()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"((SELECT quantidade FROM ordemservico where numero = {numero_os}) * "
                           f"(estprod.quantidade)) AS Qtde, "
                           f"COALESCE(prod.localizacao, ''), prod.quantidade "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id_estrutura = {id_estrut} AND prod.codigo = {codigo_produto} "
                           f"ORDER BY prod.descricao;")
            item_estrutura = cursor.fetchall()
            if not item_estrutura:
                self.adiciona_layout_substituto()
                self.line_Qtde_Manu.setFocus()
            else:
                self.line_Qtde_Manu.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_qtde_manu(self):
        try:
            qtdezinha = self.line_Qtde_Manu.text()
            if not qtdezinha:
                self.mensagem_alerta('O campo "Qtde:" não pode estar vazio')
                self.line_Qtde_Manu.clear()
                self.line_Qtde_Manu.setFocus()
            else:
                if "," in qtdezinha:
                    qtdezinha_com_ponto = qtdezinha.replace(',', '.')
                    qtdezinha_float = float(qtdezinha_com_ponto)
                else:
                    qtdezinha_float = float(qtdezinha)

                if qtdezinha_float == 0:
                    self.mensagem_alerta('O campo "Qtde:" não pode ser "0"')
                    self.line_Qtde_Manu.clear()
                    self.line_Qtde_Manu.setFocus()
                else:
                    self.verifica_saldo_produtomanual()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_saldo_produtomanual(self):
        try:
            unidadezinha = self.line_UM_Manu.text()
            qtdezinha = self.line_Qtde_Manu.text()
            saldozinho = self.line_Saldo_Manu.text()
            if "," in qtdezinha:
                qtdezinha_com_ponto = qtdezinha.replace(',', '.')
                qtdezinha_float = float(qtdezinha_com_ponto)
            else:
                qtdezinha_float = float(qtdezinha)
            if "," in saldozinho:
                saldozinho_com_ponto = saldozinho.replace(',', '.')
                saldozinho_float = float(saldozinho_com_ponto)
            else:
                saldozinho_float = float(saldozinho)
            if qtdezinha_float > saldozinho_float:
                diferenca = round((qtdezinha_float - saldozinho_float), 2)
                self.mensagem_alerta(f'Saldo deste produto é insuficiente!\n '
                                     f'Falta {diferenca} {unidadezinha} para consumir este produto')
                self.line_Qtde_Manu.clear()
            else:
                self.verifica_qtde_estrutura_manu()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_qtde_estrutura_manu(self):
        try:
            codigo_produto = self.line_Cod_Manu.text()
            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut = self.dados_os()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"((SELECT quantidade FROM ordemservico where numero = {numero_os}) * "
                           f"(estprod.quantidade)) AS Qtde, "
                           f"COALESCE(prod.localizacao, ''), prod.quantidade "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id_estrutura = {id_estrut} AND prod.codigo = {codigo_produto} "
                           f"ORDER BY prod.descricao;")
            item_estrutura = cursor.fetchall()
            if not item_estrutura:
                self.line_Cod_Subs.setFocus()
            else:
                if self.widget_Subs.isHidden() is False:
                    self.line_Cod_Subs.setFocus()
                else:
                    self.manipulando_dados_manu()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipulando_dados_manu(self):
        try:
            codigo_manu = self.line_Cod_Manu.text()

            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut = self.dados_os()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"((SELECT quantidade FROM ordemservico where numero = {numero_os}) * "
                           f"(estprod.quantidade)) AS Qtde, "
                           f"COALESCE(prod.localizacao, ''), prod.quantidade "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id_estrutura = {id_estrut} AND prod.codigo = {codigo_manu} "
                           f"ORDER BY prod.descricao;")
            item_estrutura = cursor.fetchall()

            id_mat_selecao, cod_prod, descr_prod, ref_prod, um_prod, qtde_prod, local_prod, saldo_prod = \
                item_estrutura[0]
            id_mat_sel_str = str(id_mat_selecao)
            qtde_prod_float = float(qtde_prod)

            qtde_manu = self.line_Qtde_Manu.text()
            if "," in qtde_manu:
                qtde_manu_com_ponto = qtde_manu.replace(',', '.')
                qtde_manu_float = float(qtde_manu_com_ponto)
            else:
                qtde_manu_float = float(qtde_manu)

            qtde_manu_red = "%.3f" % qtde_manu_float

            entrega = self.date_Consumo.text()
            entrega_mov = datetime.strptime(entrega, '%d/%m/%Y').date()
            data_string = '{}/{}/{}'.format(entrega_mov.day, entrega_mov.month, entrega_mov.year)

            ano_text = entrega_mov.strftime("%Y")
            mes_text = entrega_mov.strftime("%m")
            dia_text = entrega_mov.strftime("%d")

            data_hoje = date.today()
            data_hoje_str = '{}/{}/{}'.format(data_hoje.day, data_hoje.month, data_hoje.year)
            ano_atual = data_hoje.strftime("%Y")
            mes_atual = data_hoje.strftime("%m")
            dia_atual = data_hoje.strftime("%d")

            if ano_text != ano_atual:
                self.mensagem_alerta(f'Você está lançando o consumo deste '
                                     f'item no ano de {ano_text}!\n\n'
                                     f'Data Atual: {data_hoje_str}')

            if mes_text != mes_atual:
                self.mensagem_alerta(f'Você está lançando o consumo deste '
                                     f'item no mês {mes_text}!\n\n'
                                     f'Data Atual: {data_hoje_str}')

            if dia_text != dia_atual:
                self.mensagem_alerta(f'Você está lançando o consumo deste '
                                     f'item no dia {dia_text}!\n\n'
                                     f'Data Atual: {data_hoje_str}')

            extrai_total = self.jutando_tabelas_extraidas()

            soma_qtde_item = 0
            for item_para_saldo in extrai_total:
                id_mat1, cod_est1, descr_est1, ref_est1, um_est1, qtde_est1, local1, saldo1, \
                data_os1, cod_os1, descr_os1, ref_os1, um_os1, qtde_os1 = item_para_saldo
                qtde_est_float1 = float(qtde_est1)
                if id_mat_sel_str == id_mat1 and cod_os1 != "":
                    soma_qtde_item = soma_qtde_item + qtde_est_float1

            if soma_qtde_item == qtde_prod_float:
                self.mensagem_alerta('Este material já foi consumido na estrutura!')
                self.limpa_manual()
            else:
                nova_lista_total = []

                for itens_extraidos in extrai_total:
                    id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo, \
                    data_os, cod_os, descr_os, ref_os, um_os, qtde_os = itens_extraidos
                    qtde_est_float = float(qtde_est)

                    if id_mat_sel_str == id_mat and cod_os == "":
                        if qtde_manu_float > qtde_est_float:
                            self.mensagem_alerta('A quantidade é maior do que a necessidade da estrutura!')
                            dados = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo,
                                     data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                            nova_lista_total.append(dados)

                        elif qtde_manu_float == qtde_est_float:
                            dados = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_manu_red, local,
                                     saldo, data_string, cod_est, descr_est, ref_est, um_est, qtde_manu_red)
                            nova_lista_total.append(dados)

                        else:
                            dados = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_manu_red, local,
                                     saldo, data_string, cod_est, descr_est, ref_est, um_est, qtde_manu_red)

                            dif_do_saldo_qtde = "%.3f" % (qtde_est_float - qtde_manu_float)

                            dados1 = (id_mat, cod_est, descr_est, ref_est, um_est, dif_do_saldo_qtde, local,
                                      saldo, "", "", "", "", "", "")
                            nova_lista_total.append(dados)
                            nova_lista_total.append(dados1)

                    else:
                        dados = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo,
                                 data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                        nova_lista_total.append(dados)

                tabela_estrutura = []
                tabela_consumo_os = []
                for itens in nova_lista_total:
                    id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo, \
                    data_os, cod_os, descr_os, ref_os, um_os, qtde_os = itens

                    lista_est = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo)
                    tabela_estrutura.append(lista_est)

                    lista_os = (id_mat, data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                    tabela_consumo_os.append(lista_os)

                lanca_tabela(self.table_ConsumoOS, tabela_consumo_os)
                lanca_tabela(self.table_Estrutura, tabela_estrutura)
                self.pintar_tabelas()
                self.btn_Salvar.setEnabled(True)
                self.limpa_manual()

                self.table_ConsumoOS.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                self.table_Estrutura.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_manual(self):
        try:
            self.line_Cod_Manu.clear()
            self.line_Cod_Manu.setFocus()

            self.line_Ref_Manu.clear()
            self.line_Saldo_Manu.clear()
            self.line_Descricao_Manu.clear()
            self.line_Compl_Manu.clear()
            self.line_Local_Manu.clear()
            self.line_UM_Manu.clear()
            self.line_Qtde_Manu.clear()

            self.line_Cod_Subs.clear()
            self.line_Ref_Subs.clear()
            self.line_Descricao_Subs.clear()
            self.line_Compl_Subs.clear()
            self.line_UM_Subs.clear()
            self.line_Qtde_Subs.clear()

            self.remove_layout_substituto()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tudo(self):
        try:
            self.line_Num_OP.clear()
            self.line_Codigo.clear()
            self.label_Emissao.clear()
            self.line_Referencia.clear()
            self.line_Descricao.clear()
            self.line_UM.clear()
            self.line_Qtde.clear()
            self.line_UM.clear()
            self.table_Estrutura.clearContents()
            self.table_ConsumoOS.clearContents()
            self.line_Cod_Manu.setEnabled(False)
            self.line_Qtde_Manu.setEnabled(False)
            self.line_Cod_Manu.clear()
            self.line_Descricao_Manu.clear()
            self.line_Ref_Manu.clear()
            self.line_UM_Manu.clear()
            self.line_Qtde_Manu.clear()
            self.line_Local_Manu.clear()
            self.line_Saldo_Manu.clear()
            self.line_Num_OP.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_cod_subs(self):
        try:
            codigo_produto = self.line_Cod_Subs.text()
            if len(codigo_produto) == 0:
                self.mensagem_alerta('O campo "Código" não pode estar vazio')
                self.line_Cod_Subs.clear()
            elif int(codigo_produto) == 0:
                self.mensagem_alerta('O campo "Código" não pode ser "0"')
                self.line_Cod_Subs.clear()
            else:
                self.verifica_sql_produto_subs()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_sql_produto_subs(self):
        try:
            codigo_produto = self.line_Cod_Subs.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT descricao, COALESCE(obs, '') as obs, unidade, "
                           f"COALESCE(localizacao, ''), quantidade "
                           f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cursor.fetchall()
            if not detalhes_produto:
                self.mensagem_alerta('Este código de produto não existe!')
                self.line_Cod_Subs.clear()
            else:
                self.lanca_dados_produto_subs()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_produto_subs(self):
        try:
            codigo_manu = self.line_Cod_Manu.text()
            if not codigo_manu:
                self.mensagem_alerta('Primeiro defina o item que deseja consumir,\n '
                                     'para depois indicar o seu substituto!')
            else:
                codigo_produto = self.line_Cod_Subs.text()
                cur = conecta.cursor()
                cur.execute(f"SELECT descricao, COALESCE(obs, '') as obs, unidade, COALESCE(localizacao, ''), "
                            f"quantidade "
                            f"FROM produto where codigo = {codigo_produto};")
                detalhes_produto = cur.fetchall()
                descricao_id, referencia_id, unidade_id, local_id, quantidade_id = detalhes_produto[0]

                id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut = self.dados_os()

                cursor = conecta.cursor()
                cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                               f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                               f"((SELECT quantidade FROM ordemservico where numero = {numero_os}) * "
                               f"(estprod.quantidade)) AS Qtde, "
                               f"COALESCE(prod.localizacao, ''), prod.quantidade "
                               f"FROM estrutura_produto as estprod "
                               f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                               f"where estprod.id_estrutura = {id_estrut} AND prod.codigo = {codigo_produto} "
                               f"ORDER BY prod.descricao;")
                item_estrutura = cursor.fetchall()

                if not item_estrutura:
                    self.mensagem_alerta('Este produto não faz parte da estrutura!')
                    self.line_Cod_Subs.clear()

                else:
                    id_mat_selecao, cod_prod, descr_prod, ref_prod, um_prod, qtde_prod, local_prod, saldo_prod = \
                        item_estrutura[0]
                    id_mat_sel_str = str(id_mat_selecao)
                    qtde_prod_float = float(qtde_prod)

                    extrai_total = self.jutando_tabelas_extraidas()

                    soma_qtde_item = 0
                    for item_para_saldo in extrai_total:
                        id_mat1, cod_est1, descr_est1, ref_est1, um_est1, qtde_est1, local1, saldo1, \
                        data_os1, cod_os1, descr_os1, ref_os1, um_os1, qtde_os1 = item_para_saldo
                        qtde_est_float1 = float(qtde_est1)
                        if id_mat_sel_str == id_mat1 and cod_os1 != "":
                            soma_qtde_item = soma_qtde_item + qtde_est_float1

                    if soma_qtde_item == qtde_prod_float:
                        self.mensagem_alerta('Este material já foi consumido na estrutura!')
                        self.line_Cod_Subs.clear()

                    else:
                        self.line_Descricao_Subs.setText(descricao_id)
                        self.line_Ref_Subs.setText(referencia_id)
                        self.line_UM_Subs.setText(unidade_id)
                        self.line_Qtde_Subs.setEnabled(True)
                        self.line_Qtde_Subs.setFocus()
                        self.btn_Consome_Subs.setEnabled(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_qtde_subs(self):
        try:
            qtdezinha = self.line_Qtde_Manu.text()
            if not qtdezinha:
                self.mensagem_alerta('O campo "Qtde:" não pode estar vazio')
                self.line_Qtde_Manu.clear()
                self.line_Qtde_Manu.setFocus()
            else:
                if "," in qtdezinha:
                    qtdezinha_com_ponto = qtdezinha.replace(',', '.')
                    qtdezinha_float = float(qtdezinha_com_ponto)
                else:
                    qtdezinha_float = float(qtdezinha)

                if qtdezinha_float == 0:
                    self.mensagem_alerta('O campo "Qtde:" não pode ser "0"')
                    self.line_Qtde_Manu.clear()
                    self.line_Qtde_Manu.setFocus()
                else:
                    qtdezinha = self.line_Qtde_Subs.text()
                    if "," in qtdezinha:
                        qtdezinha_com_ponto = qtdezinha.replace(',', '.')
                        qtdezinha_float = float(qtdezinha_com_ponto)
                    else:
                        qtdezinha_float = float(qtdezinha)
                    if len(qtdezinha) == 0:
                        self.mensagem_alerta('O campo "Qtde:" não pode estar vazio')
                        self.line_Qtde_Subs.clear()
                        self.line_Qtde_Subs.setFocus()
                    elif qtdezinha_float == 0:
                        self.mensagem_alerta('O campo "Qtde:" não pode ser "0"')
                        self.line_Qtde_Subs.clear()
                        self.line_Qtde_Subs.setFocus()
                    else:
                        unidadezinha = self.line_UM_Manu.text()
                        qtdezinha = self.line_Qtde_Manu.text()
                        saldozinho = self.line_Saldo_Manu.text()
                        if "," in qtdezinha:
                            qtdezinha_com_ponto = qtdezinha.replace(',', '.')
                            qtdezinha_float = float(qtdezinha_com_ponto)
                        else:
                            qtdezinha_float = float(qtdezinha)
                        if "," in saldozinho:
                            saldozinho_com_ponto = saldozinho.replace(',', '.')
                            saldozinho_float = float(saldozinho_com_ponto)
                        else:
                            saldozinho_float = float(saldozinho)
                        if qtdezinha_float > saldozinho_float:
                            diferenca = round((qtdezinha_float - saldozinho_float), 2)
                            self.mensagem_alerta(f'Saldo deste produto é insuficiente!\n '
                                                 f'Falta {diferenca} {unidadezinha} para consumir este produto')
                            self.line_Qtde_Manu.clear()
                        else:
                            self.manipulando_dados_subs()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipulando_dados_subs(self):
        try:
            codigo_manu = self.line_Cod_Manu.text()

            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut = self.dados_os()

            cur = conecta.cursor()
            cur.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade "
                        f"FROM produto where codigo = {codigo_manu};")
            detalhes_produto = cur.fetchall()
            descr_manu, ref_manu, um_manu = detalhes_produto[0]

            qtde_manu = self.line_Qtde_Manu.text()
            if "," in qtde_manu:
                qtde_manu_com_ponto = qtde_manu.replace(',', '.')
                qtde_manu_float = float(qtde_manu_com_ponto)
            else:
                qtde_manu_float = float(qtde_manu)

            qtde_manu_red = "%.3f" % qtde_manu_float

            codigo_subs = self.line_Cod_Subs.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"((SELECT quantidade FROM ordemservico where numero = {numero_os}) * "
                           f"(estprod.quantidade)) AS Qtde, "
                           f"COALESCE(prod.localizacao, ''), prod.quantidade "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id_estrutura = {id_estrut} AND prod.codigo = {codigo_subs} "
                           f"ORDER BY prod.descricao;")
            item_estrutura = cursor.fetchall()

            id_mat_selecao, cod_prod, descr_prod, ref_prod, um_prod, qtde_prod, local_prod, saldo_prod = \
                item_estrutura[0]
            id_mat_sel_str = str(id_mat_selecao)
            qtde_prod_float = float(qtde_prod)

            qtde_subs = self.line_Qtde_Subs.text()
            if "," in qtde_subs:
                qtde_subs_com_ponto = qtde_subs.replace(',', '.')
                qtde_subs_float = float(qtde_subs_com_ponto)
            else:
                qtde_subs_float = float(qtde_subs)

            qtde_subs_red = "%.3f" % qtde_subs_float

            entrega = self.date_Consumo.text()
            entrega_mov = datetime.strptime(entrega, '%d/%m/%Y').date()
            data_string = '{}/{}/{}'.format(entrega_mov.day, entrega_mov.month, entrega_mov.year)

            ano_text = entrega_mov.strftime("%Y")
            mes_text = entrega_mov.strftime("%m")
            dia_text = entrega_mov.strftime("%d")

            data_hoje = date.today()
            data_hoje_str = '{}/{}/{}'.format(data_hoje.day, data_hoje.month, data_hoje.year)
            ano_atual = data_hoje.strftime("%Y")
            mes_atual = data_hoje.strftime("%m")
            dia_atual = data_hoje.strftime("%d")

            if ano_text != ano_atual:
                self.mensagem_alerta(f'Você está lançando o consumo deste '
                                     f'item no ano de {ano_text}!\n\n'
                                     f'Data Atual: {data_hoje_str}')

            if mes_text != mes_atual:
                self.mensagem_alerta(f'Você está lançando o consumo deste '
                                     f'item no mês {mes_text}!\n\n'
                                     f'Data Atual: {data_hoje_str}')

            if dia_text != dia_atual:
                self.mensagem_alerta(f'Você está lançando o consumo deste '
                                     f'item no dia {dia_text}!\n\n'
                                     f'Data Atual: {data_hoje_str}')

            extrai_total = self.jutando_tabelas_extraidas()

            soma_qtde_item = 0
            for item_para_saldo in extrai_total:
                id_mat1, cod_est1, descr_est1, ref_est1, um_est1, qtde_est1, local1, saldo1, \
                data_os1, cod_os1, descr_os1, ref_os1, um_os1, qtde_os1 = item_para_saldo
                qtde_est_float1 = float(qtde_est1)
                if id_mat_sel_str == id_mat1 and cod_os1 != "":
                    soma_qtde_item = soma_qtde_item + qtde_est_float1

            if soma_qtde_item == qtde_prod_float:
                self.mensagem_alerta('Este material já foi consumido na estrutura!')
                self.limpa_subs()
            else:
                nova_lista_total = []

                for itens_extraidos in extrai_total:
                    id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo, \
                    data_os, cod_os, descr_os, ref_os, um_os, qtde_os = itens_extraidos
                    qtde_est_float = float(qtde_est)
                    saldo_float = float(saldo)

                    if id_mat_sel_str == id_mat and cod_os == "":
                        if qtde_subs_float > qtde_est_float:
                            self.mensagem_alerta('A quantidade é maior do que resta consumir na "OP"!')
                            dados = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo,
                                     data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                            nova_lista_total.append(dados)

                        elif qtde_subs_float == qtde_est_float:
                            dados = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_subs_red, local,
                                     saldo, data_string, codigo_manu, descr_manu, ref_manu, um_manu,
                                     qtde_manu_red)
                            nova_lista_total.append(dados)

                        else:
                            dados = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_subs_red, local, saldo,
                                     data_string, codigo_manu, descr_manu, ref_manu, um_manu, qtde_manu_red)

                            dif_do_saldo_qtde = "%.3f" % (qtde_est_float - qtde_subs_float)
                            dados1 = (id_mat, cod_est, descr_est, ref_est, um_est, dif_do_saldo_qtde, local,
                                      saldo_float, "", "", "", "", "", "")
                            nova_lista_total.append(dados)
                            nova_lista_total.append(dados1)

                    else:
                        dados = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo,
                                 data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                        nova_lista_total.append(dados)

                tabela_estrutura = []
                tabela_consumo_os = []
                for itens in nova_lista_total:
                    id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo, \
                    data_os, cod_os, descr_os, ref_os, um_os, qtde_os = itens

                    lista_est = (id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo)
                    tabela_estrutura.append(lista_est)

                    lista_os = (id_mat, data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                    tabela_consumo_os.append(lista_os)

                lanca_tabela(self.table_ConsumoOS, tabela_consumo_os)
                lanca_tabela(self.table_Estrutura, tabela_estrutura)
                self.pintar_tabelas()
                self.btn_Salvar.setEnabled(True)
                self.limpa_manual()

                self.table_ConsumoOS.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                self.table_Estrutura.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_subs(self):
        try:
            self.line_Cod_Subs.clear()
            self.line_Cod_Subs.setFocus()

            self.line_Ref_Subs.clear()
            self.line_Descricao_Subs.clear()
            self.line_Compl_Subs.clear()
            self.line_UM_Subs.clear()
            self.line_Qtde_Subs.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def exibe_substituto(self):
        try:
            if self.widget_Subs.isHidden() is True:
                self.adiciona_layout_substituto()
                self.line_Cod_Manu.setFocus()
            else:
                self.remove_layout_substituto()
                self.btn_Consome_Manu.setHidden(False)
                self.line_Qtde_Subs.setEnabled(True)
                self.line_Cod_Manu.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_item(self):
        try:
            consumo_os = extrair_tabela(self.table_ConsumoOS)
            extrai_total = self.jutando_tabelas_extraidas()

            if not consumo_os:
                self.mensagem_alerta('A tabela "Consumo OP" não tem itens para excluir')

            else:
                tabela_estrutura = []
                tabela_consumo_os = []
                linha_selecao = self.table_ConsumoOS.currentRow()

                item_para_excluir = extrai_total[linha_selecao]

                id_mat_exc, cod_exc, desc_exc, ref_exc, um_exc, qtde_exc, local_exc, saldo_exc, \
                data_os_exc, cod_os_exc, desc_os_exc, ref_os_exc, um_os_exc, qtde_os_exc = item_para_excluir

                item_encontrado = [s for s in extrai_total if id_mat_exc in s]
                linhas_item_encontrado = len(item_encontrado)

                achado_sem_codigo = 0

                for item in item_encontrado:
                    id_mat_enc, cod_est_enc, desc_est_enc, ref_est_enc, umest__enc, qtde_est_enc, \
                    local_enc, saldo_enc, data_enc, cod_os_enc, desc_os_enc, ref_os_enc, \
                    um_os_enc, qtde_os_enc = item
                    if cod_os_enc == "":
                        achado_sem_codigo = achado_sem_codigo + 1

                if cod_os_exc == "":
                    self.mensagem_alerta('Escolha um item para excluir!')
                else:
                    for dados in enumerate(extrai_total):
                        id_mat_est, cod_est, desc_est, ref_est, um_est, qtde_est, local_est, saldo_est, \
                        data_os, cod_os, desc_os, ref_os, um_os, qtde_os = dados[1]

                        index = dados[0]

                        id_mat_sel, cod_sel, desc_sel, ref_sel, um_sel, qtde_sel, local_sel, saldo_sel, \
                        data_os_sel, cod_os_sel, desc_os_sel, ref_os_sel, um_os_sel, \
                        qtde_os_sel = extrai_total[linha_selecao]

                        if linha_selecao == index:
                            if linhas_item_encontrado == 1:
                                lista_est = (id_mat_est, cod_est, desc_est, ref_est, um_est,
                                             qtde_est, local_est, saldo_est)
                                tabela_estrutura.append(lista_est)

                                lista_os = (id_mat_est, "", "", "", "", "", "")
                                tabela_consumo_os.append(lista_os)

                            elif achado_sem_codigo == 0:
                                lista_est = (id_mat_est, cod_est, desc_est, ref_est, um_est, qtde_est,
                                             local_est, saldo_est)
                                tabela_estrutura.append(lista_est)

                                lista_os = (id_mat_est, "", "", "", "", "", "")
                                tabela_consumo_os.append(lista_os)

                        elif id_mat_est == id_mat_sel and linhas_item_encontrado != 1 and cod_os == "":
                            qtde_atual = float(qtde_est) + float(qtde_sel)
                            qtde_redondo = "%.3f" % qtde_atual
                            lista_est = (id_mat_est, cod_est, desc_est, ref_est, um_est, qtde_redondo,
                                         local_est, saldo_est)
                            tabela_estrutura.append(lista_est)

                            lista_os = (id_mat_est, "", "", "", "", "", "")
                            tabela_consumo_os.append(lista_os)

                        else:
                            lista_est = (id_mat_est, cod_est, desc_est, ref_est, um_est, qtde_est, local_est, saldo_est)
                            tabela_estrutura.append(lista_est)

                            lista_os = (id_mat_est, data_os, cod_os, desc_os, ref_os, um_os, qtde_os)
                            tabela_consumo_os.append(lista_os)

                    lanca_tabela(self.table_ConsumoOS, tabela_consumo_os)
                    lanca_tabela(self.table_Estrutura, tabela_estrutura)
                    self.pintar_tabelas()

                    self.table_ConsumoOS.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                    self.table_Estrutura.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_salvamento1(self):
        try:
            num_op = self.line_Num_OP.text()

            estrutura = extrair_tabela(self.table_Estrutura)
            consumo_os = extrair_tabela(self.table_ConsumoOS)

            linhas_est = len(estrutura)
            diferentes = 0
            sem_saldo = 0
            prod_sem_saldo = []

            for linha_est in range(linhas_est):
                id_mat_os, data_os, cod_os, descr_os, ref_os, um_os, qtde_os = consumo_os[linha_est]
                if not cod_os:
                    diferentes = diferentes + 1
                else:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, codigo, quantidade FROM produto where codigo = {cod_os};")
                    dados_produto = cursor.fetchall()
                    id_produto, codigo, quantidade = dados_produto[0]

                    quantidade_str = str(quantidade)

                    if quantidade < 0:
                        sem_saldo = sem_saldo + 1
                        dados = (cod_os, descr_os, quantidade_str)
                        prod_sem_saldo.append(dados)

            if diferentes > 0:
                self.mensagem_alerta(f'Esta Ordem de Produção tem divergências com a estrutura!')
            elif sem_saldo > 0:
                texto_composto = ""
                if len(prod_sem_saldo) > 1:
                    for titi in prod_sem_saldo:
                        cod_os, descr_os, quantidade = titi
                        texto = "- " + cod_os + " - " + descr_os + " - Saldo: " + quantidade
                        texto_composto = texto_composto + "\n" + texto

                    self.mensagem_alerta(f'Os produtos abaixo estão sem saldo para '
                                         f'encerrar a\n'
                                         f'Ordem de Produção Nº {num_op}\n'
                                         f'{texto_composto}!')
                else:
                    cod_os, descr_os, quantidade = prod_sem_saldo[0]
                    texto = "- " + cod_os + " - " + descr_os + " - Saldo: " + quantidade
                    self.mensagem_alerta(f'O produto abaixo está sem saldo para '
                                         f'encerrar a\n'
                                         f'Ordem de Produção Nº {num_op}\n'
                                         f'{texto}!')
            else:
                self.verifica_salvamento2()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_salvamento2(self):
        try:
            consumo_os = extrair_tabela(self.table_ConsumoOS)

            if not consumo_os:
                self.mensagem_alerta('A Tabela "Consumo OP" não possui produtos lançados!')
            else:
                self.salvar_lista()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_lista(self):
        try:
            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, id_estrut = self.dados_os()
            itens_manipula_total = self.select_mistura()

            select_antigo = []
            for dados in itens_manipula_total:
                id_mat, cod, descr, ref, um, qtde, local, saldo, \
                data_os, cod_os, descr_os, ref_os, um_os, qtde_os = dados

                if cod_os != "":
                    id_mat_str = str(id_mat)

                    cod_str = str(cod)

                    qtde_float = float(qtde)
                    qtde_red = "%.3f" % qtde_float
                    qtde_str = str(qtde_red)

                    saldo_float = float(saldo)
                    saldo_red = "%.3f" % saldo_float
                    saldo_str = str(saldo_red)

                    cod_os_str = str(cod_os)

                    qtde_os_float = float(qtde_os)
                    qtde_os_red = "%.3f" % qtde_os_float
                    qtde_os_str = str(qtde_os_red)

                    dados_os = (id_mat_str, cod_str, descr, ref, um, qtde_str, local, saldo_str,
                                data_os, cod_os_str, descr_os, ref_os, um_os, qtde_os_str)
                    select_antigo.append(dados_os)

            select_final = []
            extrai_total = self.jutando_tabelas_extraidas()

            for dadus in extrai_total:
                id_mat, cod, descr, ref, um, qtde, local, saldo, \
                data_os, cod_os, descr_os, ref_os, um_os, qtde_os = dadus

                if cod_os != "":
                    dados_fi = (id_mat, cod, descr, ref, um, qtde, local, saldo,
                                data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                    select_final.append(dados_fi)

            itens_para_adicionar = []
            itens_para_excluir = []
            for i in range(len(select_final)):
                if not select_final[i] in select_antigo:
                    itens_para_adicionar.append(select_final[i])

            for i in range(len(select_antigo)):
                if not select_antigo[i] in select_final:
                    itens_para_excluir.append(select_antigo[i])

            if not itens_para_excluir and not itens_para_adicionar:
                self.mensagem_alerta('Não foi consumido nenhum item nesta OP!')
            else:
                if not itens_para_excluir:
                    pass
                else:
                    for dado_exc in itens_para_excluir:
                        id_mat, cod, descr, ref, um, qtde, local, saldo, \
                        data_os, cod_os, descr_os, ref_os, um_os, qtde_os = dado_exc

                        cursor = conecta.cursor()
                        cursor.execute(
                            f"SELECT id, movimentacao, codigo FROM produtoos where numero = {numero_os} "
                            f"and codigo = {cod_os} "
                            f"and id_estrut_prod = {id_mat} "
                            f"and quantidade = {qtde_os} "
                            f"and QTDE_ESTRUT_PROD = {qtde};")
                        resultado = cursor.fetchall()

                        if len(resultado) > 1:
                            id_prod_servico, id_movimento, codigo_item = resultado[0]

                            cursor = conecta.cursor()
                            cursor.execute(f"DELETE FROM produtoos WHERE ID = {id_prod_servico};")

                            cursor = conecta.cursor()
                            cursor.execute(f"DELETE FROM movimentacao WHERE ID = {id_movimento};")
                        else:
                            id_prod_servico, id_movimento, codigo_item = resultado[0]

                            cursor = conecta.cursor()
                            cursor.execute(f"DELETE FROM produtoos WHERE ID = {id_prod_servico};")

                            cursor = conecta.cursor()
                            cursor.execute(f"DELETE FROM movimentacao WHERE ID = {id_movimento};")

                if not itens_para_adicionar:
                    pass
                else:
                    for dado_adi in itens_para_adicionar:
                        id_mat, cod, descr, ref, um, qtde, local, saldo, \
                        data_os, cod_os, descr_os, ref_os, um_os, qtde_os = dado_adi
                        cod_os_int = int(cod_os)

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT id, codigo FROM produto where codigo = {cod_os_int};")
                        dados_produto = cursor.fetchall()
                        id_produto, codigo = dados_produto[0]

                        cursor = conecta.cursor()
                        cursor.execute("select GEN_ID(GEN_MOVIMENTACAO_ID,0) from rdb$database;")
                        ultimo_id0 = cursor.fetchall()
                        ultimo_id1 = ultimo_id0[0]
                        ultimo_id = int(ultimo_id1[0]) + 1

                        numero_os_str = f"OP {str(numero_os)}"

                        data_certa = datetime.strptime(data_os, '%d/%m/%Y').date()

                        qtde_os_float = float(qtde_os)
                        qtde_os_red = "%.3f" % qtde_os_float

                        qtde_float = float(qtde)
                        qtde_red = "%.3f" % qtde_float

                        id_mat_int = int(id_mat)

                        cursor = conecta.cursor()
                        cursor.execute(f"Insert into movimentacao "
                                       f"(ID, PRODUTO, OBS, TIPO, QUANTIDADE, DATA, CODIGO, funcionario,  "
                                       f"localestoque) "
                                       f"values (GEN_ID(GEN_MOVIMENTACAO_ID,1), "
                                       f"{id_produto}, '{numero_os_str}', 210, {qtde_os_red}, '{data_certa}', "
                                       f"{codigo}, "
                                       f"6, 3);")

                        cursor = conecta.cursor()
                        cursor.execute(f"Insert into produtoos (ID, DATA, PRODUTO, QUANTIDADE, MESTRE, "
                                       f"funcionarios,  movimentacao,  NUMERO, CODIGO, ID_ESTRUT_PROD, "
                                       f"QTDE_ESTRUT_PROD, DATA_CONSUMO) "
                                       f"values (GEN_ID(GEN_PRODUTOOS_ID,1), '{data_certa}', "
                                       f"{id_produto}, {qtde_os_red}, {id_os}, 6, {ultimo_id}, "
                                       f"{numero_os}, {codigo}, {id_mat_int}, {qtde_red}, '{data_certa}');")

                somando_negativos = 0
                cod_item_negativo = ""
                if not itens_para_adicionar:
                    pass
                else:
                    for dado_adi in itens_para_adicionar:
                        id_mat, cod, descr, ref, um, qtde, local, saldo, \
                        data_os, cod_os, descr_os, ref_os, um_os, qtde_os = dado_adi

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT codigo, quantidade FROM produto where codigo = {cod_os};")
                        estoque_item = cursor.fetchall()
                        codiguzinho, quantidade = estoque_item[0]

                        if quantidade < 0:
                            cod_item_negativo = codiguzinho
                            somando_negativos = somando_negativos + 1

                if somando_negativos > 0:
                    conecta.rollback()
                    self.mensagem_alerta(f"O código {cod_item_negativo} ficou com saldo "
                                         f"negativo e o consumo não foi salvo!")
                else:
                    conecta.commit()
                    self.mensagem_alerta("Material lançado com sucesso!")

            self.reiniciar()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def reiniciar(self):
        try:
            self.line_Num_OP.clear()
            self.line_Codigo.clear()
            self.label_Emissao.clear()
            self.line_Referencia.clear()
            self.line_Descricao.clear()
            self.line_UM.clear()
            self.line_Qtde.clear()
            self.table_Estrutura.setRowCount(0)
            self.table_ConsumoOS.setRowCount(0)
            self.line_Cod_Manu.setEnabled(False)
            self.line_Qtde_Manu.setEnabled(False)
            self.line_Cod_Manu.clear()
            self.line_Descricao_Manu.clear()
            self.line_Ref_Manu.clear()
            self.line_UM_Manu.clear()
            self.line_Qtde_Manu.clear()
            self.line_Local_Manu.clear()
            self.line_Saldo_Manu.clear()

            self.remove_layout_substituto()
            self.btn_Consome_Manu.setHidden(False)
            self.line_Qtde_Subs.setEnabled(True)

            self.line_Num_OP.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    opconsumo = TelaOpConsumirV2()
    opconsumo.show()
    qt.exec_()

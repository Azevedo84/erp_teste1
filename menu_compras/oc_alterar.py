import sys
from banco_dados.conexao import conecta
from forms.tela_oc_alterar import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.lines import definir_data_atual
from comandos.cores import cor_amarelo, cor_branco, cor_cinza_claro
from comandos.telas import tamanho_aplicacao, icone
from comandos.conversores import valores_para_float, valores_para_virgula
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidgetItem, QMessageBox
from PyQt5.QtGui import QColor
from datetime import datetime, date, timedelta
import inspect
import os
import traceback


class TelaOcAlterar(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_compra_sol.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Req_Abertas)
        layout_cabec_tab(self.table_Produtos_OC)

        self.definir_botoes_e_comandos()
        self.definir_bloqueios()
        self.definir_validador()
        self.definir_entrega()
        definir_data_atual(self.date_Emissao)
        self.manipula_dados_req()

        self.escolher_produto = []

        self.table_Produtos_OC.viewport().installEventFilter(self)

        self.lista_requisicoes = []
        self.lista_produtos_oc = []

        self.closeEvent = self.ao_fechar
        self.carregar_dados()

        self.processando = False

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

    def alterar_item(self):
        try:
            nome_tabela = self.table_Produtos_OC

            dados_tab = extrair_tabela(nome_tabela)
            if not dados_tab:
                self.mensagem_alerta(f'A tabela "Tabela Requisições Abertas" está vazia!')
            else:
                linha = nome_tabela.currentRow()
                if linha >= 0:
                    data_str = self.date_Emissao.text()
                    data_emissao = datetime.strptime(data_str, '%d/%m/%Y')

                    num_req, item_req, cod_prod, descr, ref, um, qtde, unit, ipi, tot, entr, qtde_nf = dados_tab[linha]

                    qtde_float = valores_para_float(qtde)
                    q_nf_float = valores_para_float(qtde_nf)

                    if qtde_float == q_nf_float:
                        self.mensagem_alerta(f'O Código {cod_prod} não pode ser alterado pois já foi '
                                             f'encerrado e possui Notas Fiscais de compra vinculadas!')
                    else:
                        dados_tab[linha].append(data_emissao.strftime('%d/%m/%Y'))
                        self.abrir_tela_alteracao_prod(dados_tab[linha])

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def abrir_tela_alteracao_prod(self, lista_dados):
        try:
            from menu_compras.oc_alterar_prod import TelaOcAlterarProduto

            self.escolher_produto = TelaOcAlterarProduto(lista_dados)
            self.escolher_produto.produto_escolhido.connect(self.atualizar_produto_entry)
            self.escolher_produto.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def atualizar_produto_entry(self, produto):
        nome_tabela = self.table_Produtos_OC

        dados_tab = extrair_tabela(nome_tabela)
        if not dados_tab:
            self.mensagem_alerta(f'A tabela "Tabela Requisições Abertas" está vazia!')
        else:
            linha = nome_tabela.currentRow()
            if linha >= 0:
                for coluna, valor in enumerate(produto):
                    item = QTableWidgetItem(valor)
                    nome_tabela.setItem(linha, coluna, item)

                self.atualiza_valor_total()
                self.pinta_tabela()
                self.limpa_dados_produtos()
                self.line_Codigo.setFocus()

    def ao_fechar(self, event):
        try:
            print("encerrar")
            event.accept()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def carregar_dados(self):
        try:
            print("Dados carregados")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def dados_banco_convertidos(self, id_oc):
        try:
            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT ocprod.numero, ocprod.dataentrega, prod.codigo, prod.descricao, prod.embalagem, "
                f"prod.obs, prod.unidade, "
                f"ocprod.quantidade, ocprod.unitario, "
                f"ocprod.ipi, ocprod.id_prod_req, ocprod.produzido "
                f"FROM produtoordemcompra as ocprod "
                f"INNER JOIN produto as prod ON ocprod.produto = prod.ID "
                f"where ocprod.mestre = {id_oc};")
            dados_produtos = cursor.fetchall()

            tabela = []
            erros = 0
            for i in dados_produtos:
                numero, entrega, cod, descr, embal, ref, um, qtde, unit, ipi, id_req, qtde_nf = i

                if not id_req:
                    self.mensagem_alerta(f'O Código {cod} está sem a requisição vinculada com a OC!')
                    erros += 1
                else:
                    data_for = entrega.strftime('%d/%m/%Y')

                    if qtde_nf:
                        qtde_nf_float = float(qtde_nf)
                        qtde_nf_vi = valores_para_virgula(str(qtde_nf_float))
                    else:
                        qtde_nf_vi = "0,00"

                    if qtde:
                        qtde_float = float(qtde)
                        qtde_virg = valores_para_virgula(str(qtde_float))
                    else:
                        qtde_float = 0.00
                        qtde_virg = "0,00"

                    if unit:
                        unit_float = float(unit)
                        unit_2casas = ("%.4f" % unit_float)
                        unit_string = valores_para_virgula(unit_2casas)
                        unit_fim = "R$ " + unit_string
                    else:
                        unit_float = 0.00
                        unit_fim = "R$ 0,0000"

                    if ipi:
                        ipi_float = valores_para_float(ipi)
                        ipi_2casas = ("%.2f" % ipi_float)
                        ipi_string = valores_para_virgula(ipi_2casas)
                        ipi_fim = ipi_string + "%"
                    else:
                        ipi_fim = "0,00%"

                    total_float = qtde_float * unit_float
                    total_2casas = ("%.2f" % total_float)
                    total_string = valores_para_virgula(total_2casas)
                    total_fim = "R$ " + total_string

                    dados = (id_req, cod, descr, ref, um, qtde_virg, unit_fim, ipi_fim, total_fim, data_for, qtde_nf_vi)
                    tabela.append(dados)

            return tabela

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_bloqueios(self):
        try:
            self.line_CodForn.setReadOnly(True)
            self.line_NomeForn.setReadOnly(True)

            self.date_Emissao.setReadOnly(True)

            self.line_Descricao.setReadOnly(True)
            self.line_UM.setReadOnly(True)
            self.line_ValorTotal.setReadOnly(True)

            self.line_Total_Merc.setReadOnly(True)
            self.line_Total_Ipi.setReadOnly(True)
            self.line_Total_Geral.setReadOnly(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_botoes_e_comandos(self):
        try:
            self.line_NumOC.editingFinished.connect(self.verifica_line_oc)

            self.line_Codigo.editingFinished.connect(self.verifica_line_codigo)
            self.line_Qtde.editingFinished.connect(self.verifica_line_qtde)
            self.line_Unit.editingFinished.connect(self.atualiza_mascara_unit)
            self.line_Ipi.editingFinished.connect(self.atualiza_mascara_ipi)

            self.line_ValorTotal.editingFinished.connect(self.verifica_entrega)
            self.btn_AdicionarProd.clicked.connect(self.verifica_entrega)

            self.line_Frete.returnPressed.connect(lambda: self.atualiza_mascara_frete())
            self.line_Desconto.returnPressed.connect(lambda: self.atualiza_mascara_desconto())

            self.btn_ExcluirItem.clicked.connect(self.excluir_item_tab_produtos)
            self.btn_ExcluirTudo.clicked.connect(self.excluir_tudo_tab_produtos)

            self.btn_Excluir_OC.clicked.connect(self.excluir_ordem)

            self.btn_Limpar.clicked.connect(self.limpar_tudo)

            self.btn_Alterar_Item.clicked.connect(self.alterar_item)

            self.btn_Salvar.clicked.connect(self.verifica_salvamento)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_validador(self):
        try:
            validator = QtGui.QRegExpValidator(QtCore.QRegExp(r'\d+'), self.line_NumOC)
            self.line_NumOC.setValidator(validator)

            validator = QtGui.QRegExpValidator(QtCore.QRegExp(r'\d+'), self.line_CodForn)
            self.line_CodForn.setValidator(validator)

            validator = QtGui.QRegExpValidator(QtCore.QRegExp(r'\d+'), self.line_Codigo)
            self.line_Codigo.setValidator(validator)

            validator = QtGui.QDoubleValidator(0, 9999999.000, 3, self.line_Qtde)
            locale = QtCore.QLocale("pt_BR")
            validator.setLocale(locale)
            self.line_Qtde.setValidator(validator)

            validator = QtGui.QDoubleValidator(0, 999.00, 2, self.line_Ipi)
            locale = QtCore.QLocale("pt_BR")
            validator.setLocale(locale)
            self.line_Ipi.setValidator(validator)

            validator = QtGui.QDoubleValidator(0, 9999999.0000, 4, self.line_Unit)
            locale = QtCore.QLocale("pt_BR")
            validator.setLocale(locale)
            self.line_Unit.setValidator(validator)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_entrega(self):
        try:
            data_hoje = date.today()
            data_entrega = data_hoje + timedelta(days=14)
            self.date_Entrega.setDate(data_entrega)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_oc(self):
        if not self.processando:
            try:
                self.processando = True

                num_oc = self.line_NumOC.text()

                self.limpa_dados_produtos()
                self.limpa_dados_fornecedor()
                self.limpa_dados_valores_obs()

                self.table_Produtos_OC.setRowCount(0)
                self.table_Req_Abertas.setRowCount(0)

                definir_data_atual(self.date_Emissao)

                self.manipula_dados_req()
                self.line_NumOC.setFocus()

                if num_oc:
                    if int(num_oc) == 0:
                        self.mensagem_alerta('O campo "Nº OC" não pode ser "0"')
                        self.line_NumOC.clear()
                        self.line_NumOC.setFocus()
                    else:
                        self.verifica_sql_oc(num_oc)

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_sql_oc(self, num_oc):
        try:
            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT oc.id, oc.numero, oc.data, oc.fornecedor "
                f"FROM ordemcompra as oc "
                f"where oc.entradasaida = 'E' "
                f"and oc.numero = {num_oc};")
            dados_oc = cursor.fetchall()

            if not dados_oc:
                self.mensagem_alerta('Este número de OC não existe!')
                self.line_NumOC.clear()
            else:
                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT oc.numero, oc.data, forn.registro, forn.razao, oc.frete, oc.descontos, oc.obs "
                    f"FROM ordemcompra as oc "
                    f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.ID "
                    f"where oc.entradasaida = 'E' "
                    f"and oc.numero = {num_oc} "
                    f"and oc.status = 'A';")
                dados_oc_aberta = cursor.fetchall()

                if not dados_oc_aberta:
                    self.mensagem_alerta('Este número de OC já foi encerrada!')
                    self.line_NumOC.clear()
                else:
                    self.lanca_dados_ordemcompra(dados_oc_aberta)

                    id_oc = dados_oc[0][0]

                    cursor = conecta.cursor()
                    cursor.execute(
                        f"SELECT reqprod.numero, reqprod.item, prod.codigo, prod.descricao, COALESCE(prod.obs, ''), "
                        f"prod.unidade, ocprod.quantidade, ocprod.unitario, ocprod.ipi, ocprod.dataentrega, "
                        f"ocprod.produzido "
                        f"FROM produtoordemcompra as ocprod "
                        f"INNER JOIN produtoordemrequisicao as reqprod ON ocprod.id_prod_req = reqprod.id "
                        f"INNER JOIN produto as prod ON ocprod.produto = prod.ID "
                        f"where ocprod.mestre = {id_oc};")
                    dados_produtos = cursor.fetchall()

                    if dados_produtos:
                        self.lanca_dados_produtoordemcompra(dados_produtos)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_ordemcompra(self, dados_oc):
        try:
            num_oc, emissao, cod_forn, nom_forn, frete, desconto, obs = dados_oc[0]

            self.date_Emissao.setDate(emissao)

            self.line_CodForn.setText(cod_forn.strip())
            self.line_NomeForn.setText(nom_forn)

            self.line_Frete.setText(str(frete))
            self.atualiza_mascara_frete()

            self.line_Desconto.setText(str(desconto))
            self.atualiza_mascara_desconto()

            self.line_Obs.setText(obs)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_produtoordemcompra(self, dados_produtos):
        try:
            tabela = []

            for i in dados_produtos:
                num_req, item_req, cod, descr, ref, um, qtde, unit, ipi, entrega, qtde_nf = i

                data_for = entrega.strftime('%d/%m/%Y')

                if qtde_nf:
                    qtde_nf_float = float(qtde_nf)
                    qtde_nf_vi = valores_para_virgula(str(qtde_nf_float))
                else:
                    qtde_nf_vi = "0,00"

                if qtde:
                    qtde_float = float(qtde)
                    qtde_virg = valores_para_virgula(str(qtde_float))
                else:
                    qtde_float = 0.00
                    qtde_virg = "0,00"

                if unit:
                    unit_float = float(unit)
                    unit_2casas = ("%.4f" % unit_float)
                    unit_string = valores_para_virgula(unit_2casas)
                    unit_fim = "R$ " + unit_string
                else:
                    unit_float = 0.00
                    unit_fim = "R$ 0,0000"

                if ipi:
                    ipi_float = valores_para_float(ipi)
                    ipi_2casas = ("%.2f" % ipi_float)
                    ipi_string = valores_para_virgula(ipi_2casas)
                    ipi_fim = ipi_string + "%"
                else:
                    ipi_fim = "0,00%"

                total_float = qtde_float * unit_float
                total_2casas = ("%.2f" % total_float)
                total_string = valores_para_virgula(total_2casas)
                total_fim = "R$ " + total_string

                dados = (num_req, item_req, cod, descr, ref, um, qtde_virg, unit_fim, ipi_fim, total_fim,
                         data_for, qtde_nf_vi)
                tabela.append(dados)

            if tabela:
                lanca_tabela(self.table_Produtos_OC, tabela)
                self.lista_produtos_oc = tabela

                self.atualiza_valor_total()
                self.pinta_tabela()
                self.limpa_dados_produtos()
                self.line_Codigo.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_codigo(self):
        if not self.processando:
            try:
                self.processando = True

                cod_produto = self.line_Codigo.text()

                self.line_Num_Req.clear()
                self.line_Item_Req.clear()
                self.line_Descricao.clear()
                self.line_Referencia.clear()
                self.line_UM.clear()
                self.line_Qtde.clear()
                self.line_Unit.clear()
                self.line_Ipi.clear()
                self.line_ValorTotal.clear()

                self.line_Referencia.setStyleSheet(f"background-color: {cor_branco};")

                self.definir_entrega()

                if cod_produto:
                    if int(cod_produto) == 0:
                        self.mensagem_alerta('O campo "Código:" não pode ser "0"')
                        self.limpa_dados_produtos()
                        self.line_Codigo.setFocus()
                    else:
                        self.verifica_sql_codigo(cod_produto)

                else:
                    self.limpa_dados_produtos()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_sql_codigo(self, cod_produto):
        try:
            cur = conecta.cursor()
            cur.execute(f"SELECT prod.descricao, COALESCE(prod.obs, '') , prod.unidade, "
                        f"prod.localizacao, prod.quantidade, conj.conjunto, prod.embalagem "
                        f"FROM produto as prod "
                        f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                        f"where codigo = {cod_produto};")
            detalhes_produto = cur.fetchall()

            if not detalhes_produto:
                self.mensagem_alerta('Este Código de Produto não existe!')
                self.limpa_dados_produtos()
                self.line_Codigo.setFocus()
            else:
                self.lanca_dados_codigo(cod_produto, detalhes_produto)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_codigo(self, cod_produto, detalhes_produto):
        try:
            descricao, referencia, un, local, saldo, conj, embalagem = detalhes_produto[0]

            self.line_Descricao.setText(descricao)
            self.line_UM.setText(un)

            itens_encontrados = []

            dados_req_abertas = extrair_tabela(self.table_Req_Abertas)
            for i in dados_req_abertas:
                if cod_produto in i:
                    num_req, item_req, cod, descr, ref, um, qtde = i
                    tt = (num_req, item_req, cod, descr, ref, um, qtde)
                    itens_encontrados.append(tt)

            if not itens_encontrados:
                self.mensagem_alerta('Indique o Número e o Item da sequência da Requisição!')
                self.line_Num_Req.setFocus()

            elif len(itens_encontrados) > 1:
                num_reqs, item_reqs, cods, descrs, refs, ums, qtdes = itens_encontrados[0]

                self.mensagem_alerta('Indique o Item da sequência da Requisição!')
                self.line_Referencia.setText(referencia)
                if embalagem == "SIM":
                    self.line_Referencia.setStyleSheet(f"background-color: {cor_amarelo};")
                elif embalagem == "SER":
                    self.line_Referencia.setStyleSheet(f"background-color: {cor_amarelo};")
                else:
                    self.line_Referencia.setStyleSheet(f"background-color: {cor_branco};")

                self.line_Num_Req.setText(str(num_reqs))
                self.line_Item_Req.setFocus()

            else:
                num_reqs, item_reqs, cods, descrs, refs, ums, qtdes = itens_encontrados[0]

                self.line_Referencia.setText(refs)

                if embalagem == "SIM":
                    self.line_Referencia.setStyleSheet(f"background-color: {cor_amarelo};")
                elif embalagem == "SER":
                    self.line_Referencia.setStyleSheet(f"background-color: {cor_amarelo};")
                else:
                    self.line_Referencia.setStyleSheet(f"background-color: {cor_branco};")

                self.line_Num_Req.setText(str(num_reqs))
                self.line_Item_Req.setText(str(item_reqs))
                self.line_Qtde.setText(str(qtdes))
                self.line_Qtde.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_qtde(self):
        if not self.processando:
            try:
                self.processando = True

                qtde = self.line_Qtde.text()

                if qtde:
                    qtde_com_virgula = valores_para_virgula(qtde)

                    self.line_Qtde.setText(qtde_com_virgula)

                    self.calcular_valor_total_prod()

                    self.line_Unit.setFocus()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def atualiza_mascara_unit(self):
        if not self.processando:
            try:
                self.processando = True

                unit = self.line_Unit.text()

                if unit:
                    unit_float = valores_para_float(unit)
                    unit_2casas = ("%.4f" % unit_float)
                    valor_string = valores_para_virgula(unit_2casas)
                    valor_final = "R$ " + valor_string
                    self.line_Unit.setText(valor_final)

                    self.calcular_valor_total_prod()

                    self.line_Ipi.setFocus()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def atualiza_mascara_ipi(self):
        if not self.processando:
            try:
                self.processando = True

                ipi = self.line_Ipi.text()

                ipi_float = valores_para_float(ipi)
                ipi_2casas = ("%.2f" % ipi_float)
                valor_string = valores_para_virgula(ipi_2casas)
                valor_final = valor_string + "%"
                self.line_Ipi.setText(valor_final)

                self.calcular_valor_total_prod()

                self.line_ValorTotal.setFocus()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def calcular_valor_total_prod(self):
        try:
            qtde = self.line_Qtde.text()
            unit = self.line_Unit.text()

            if qtde and unit:
                qtde_float = valores_para_float(qtde)

                unit_float = valores_para_float(unit)

                valor_total = qtde_float * unit_float

                total_2casas = ("%.2f" % valor_total)
                valor_string = valores_para_virgula(total_2casas)

                valor_final = "R$ " + valor_string
                self.line_ValorTotal.setText(valor_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_entrega(self):
        if not self.processando:
            try:
                self.processando = True

                data_entrega_str = self.date_Entrega.text()
                try:
                    data_entrega = datetime.strptime(data_entrega_str, '%d/%m/%Y')

                    data_atual = datetime.combine(date.today(), datetime.min.time())

                    if data_entrega < data_atual:
                        self.mensagem_alerta(f'A data de entrega não pode ser menor que a atual!')
                    else:
                        self.verifica_dados_completos()

                except ValueError:
                    self.mensagem_alerta(f'A data de entrega não está no formato correto '
                                                                f'(dd/mm/aaaa)!')

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_dados_completos(self):
        try:
            num_oc = self.line_NumOC.text()
            cod_produto = self.line_Codigo.text()
            num_req = self.line_Num_Req.text()
            item_req = self.line_Item_Req.text()
            qtde = self.line_Qtde.text()
            unit = self.line_Unit.text()

            if not num_oc:
                self.verifica_line_oc()
            elif not cod_produto:
                self.verifica_line_codigo()
            elif not num_req:
                self.mensagem_alerta('O campo "Nº Req" não pode estar vazio')
                self.line_Num_Req.setFocus()
            elif not item_req:
                self.mensagem_alerta('O campo "Item Req" não pode estar vazio')
                self.line_Item_Req.setFocus()
            elif not qtde:
                self.mensagem_alerta('O campo "Qtde" não pode estar vazio')
                self.line_Qtde.setFocus()
            elif not unit:
                self.mensagem_alerta('O campo "R$/Unid" não pode estar vazio')
                self.line_Unit.setFocus()
            else:
                self.manipula_dados_tabela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela(self):
        try:
            self.calcular_valor_total_prod()

            cod_produto = self.line_Codigo.text()
            descr = self.line_Descricao.text()
            ref = self.line_Referencia.text()
            um = self.line_UM.text()

            num_req = self.line_Num_Req.text()
            item_req = self.line_Item_Req.text()

            qtde = self.line_Qtde.text()
            unit = self.line_Unit.text()
            ipi = self.line_Ipi.text()
            total = self.line_ValorTotal.text()
            entrega = self.date_Entrega.text()
            qtde_nf = "0,00"

            item_encontrado = self.verifica_item_requisicao(num_req, item_req)

            if item_encontrado:
                dados = [num_req, item_req, cod_produto, descr, ref, um, qtde, unit, ipi, total, entrega, qtde_nf]

                extrai_produtos = extrair_tabela(self.table_Produtos_OC)

                ja_existe = False
                for i in extrai_produtos:
                    num_req_e, item_req_e, cod_produto_e, descr_e, ref_e, um_e, qtde_e, unit_e, ipi_e, total_e, \
                    entrega_e, qtde_nf_e = i

                    if num_req_e == num_req and item_req_e == item_req:
                        ja_existe = True
                        break

                if not ja_existe:
                    extrai_produtos.append(dados)
                    if extrai_produtos:
                        lanca_tabela(self.table_Produtos_OC, extrai_produtos)

                    self.atualiza_valor_total()
                    self.pinta_tabela()
                    self.limpa_dados_produtos()
                    self.line_Codigo.setFocus()

                    self.excluir_item_requisicao(item_encontrado)
                else:
                    self.mensagem_alerta(f'O item selecionado já está presente na tabela'
                                                                f'"Produtos OC".')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_item_requisicao(self, num_req_m, item_req_m):
        try:
            itens_encontrados = []

            dados_req_abertas = extrair_tabela(self.table_Req_Abertas)
            for indice, i in enumerate(dados_req_abertas):
                num_req, item_req, cod, descr, ref, um, qtde = i

                if num_req_m == num_req and item_req_m == item_req:
                    tt = (indice, num_req, item_req, cod, descr, ref, um, qtde)
                    itens_encontrados.append(tt)

            if not itens_encontrados:
                self.mensagem_alerta('Número ou o Item da sequência da Requisição não encontrados!')
                self.line_Num_Req.setFocus()

                itens_encontrados = []

            elif len(itens_encontrados) > 1:
                self.mensagem_alerta('Indique o Item da sequência da Requisição!')

                itens_encontrados = []

            return itens_encontrados

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_item_requisicao(self, item_encontrado):
        try:
            linha_selecao = item_encontrado[0][0]

            nome_tabela = self.table_Req_Abertas

            extrai_recomendados = extrair_tabela(nome_tabela)
            if not extrai_recomendados:
                self.mensagem_alerta(f'A tabela "Tabela Requisições Abertas" está vazia!')
            else:
                if linha_selecao >= 0:
                    nome_tabela.removeRow(linha_selecao)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_item_tab_produtos(self):
        try:
            nome_tabela = self.table_Produtos_OC

            dados_tab = extrair_tabela(nome_tabela)
            if not dados_tab:
                self.mensagem_alerta(f'A tabela "Produtos Ordem de Compra" está vazia!')
            else:
                linha = nome_tabela.currentRow()
                if linha >= 0:
                    num_req, item_req, cod_pr, desc, ref, um, qtde, unit, ipi, total, entr, q_nf = dados_tab[linha]

                    q_nf_float = valores_para_float(q_nf)

                    if q_nf_float:
                        self.mensagem_alerta(f'O Código {cod_pr} não pode ser excluído em razão de '
                                             f'Notas Fiscais de compra vinculadas!')
                    else:
                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT prodreq.numero, prodreq.item,  "
                                       f"prod.codigo, prod.descricao as DESCRICAO, "
                                       f"CASE prod.embalagem when 'SIM' then COALESCE(prodreq.referencia, '') "
                                       f"else COALESCE(prod.obs, '') end as REFERENCIA, "
                                       f"prod.unidade, prodreq.quantidade "
                                       f"FROM produtoordemrequisicao as prodreq "
                                       f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                                       f"WHERE prodreq.numero = {num_req} "
                                       f"and prodreq.item = {item_req} "
                                       f"ORDER BY prodreq.numero;")
                        extrair_req = cursor.fetchall()

                        if extrair_req:
                            num_req_r, item_req_r, cod_r, descr_r, ref_r, um_r, qtde_r = extrair_req[0]

                            dados = [num_req, item_req, cod_r, descr_r, ref_r, um_r, qtde_r]
                            nome_tabela.removeRow(linha)

                            extrai_produtos = extrair_tabela(self.table_Req_Abertas)

                            ja_existe = False
                            for i in extrai_produtos:
                                num_req_e, item_req_e, cod_e, descr_e, ref_e, um_e, qtde_e = i

                                if num_req_e == num_req and item_req_e == item_req:
                                    ja_existe = True
                                    break

                            if not ja_existe:
                                extrai_produtos.append(dados)
                                if extrai_produtos:
                                    extrai_produtos_ord = sorted(extrai_produtos, key=self.chave_de_ordenacao)
                                    lanca_tabela(self.table_Req_Abertas, extrai_produtos_ord)

                            self.atualiza_valor_total()
                            self.pinta_tabela()
                            self.limpa_dados_produtos()
                            self.line_Codigo.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_tudo_tab_produtos(self):
        try:
            nome_tabela = self.table_Produtos_OC

            pode_excluir = True

            dados_tab = extrair_tabela(nome_tabela)
            if not dados_tab:
                self.mensagem_alerta(f'A tabela "Produtos Ordem de Compra" está vazia!')
            else:
                for i in dados_tab:
                    num_req, item_req, cod_pr, desc, ref, um, qtde, unit, ipi, total, entr, q_nf = i

                    q_nf_float = valores_para_float(q_nf)

                    if q_nf_float:
                        pode_excluir = False
                        break

                if pode_excluir:
                    nome_tabela.setRowCount(0)
                else:
                    self.mensagem_alerta(f'Os produtos não podem ser excluídos em razão de '
                                         f'Notas Fiscais de compra vinculadas!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_ordem(self):
        try:
            num_oc = self.line_NumOC.text()

            nome_tabela = self.table_Produtos_OC
            dados_tab = extrair_tabela(nome_tabela)

            pode_excluir = True

            if not num_oc:
                self.mensagem_alerta(f'O campo "Nº OC" não pode estar vazio!')

            else:
                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT oc.id, oc.numero, oc.data, oc.fornecedor "
                    f"FROM ordemcompra as oc "
                    f"where oc.entradasaida = 'E' "
                    f"and oc.numero = {num_oc};")
                dados_oc = cursor.fetchall()

                if not dados_oc:
                    self.mensagem_alerta('Este número de OC não existe!')
                else:
                    if int(num_oc) == 0:
                        self.mensagem_alerta('O campo "Nº OC" não pode ser "0"')

                    elif not dados_tab:
                        self.mensagem_alerta(f'A tabela "Produtos Ordem de Compra" está vazia!')
                    else:
                        for i in dados_tab:
                            num_req, item_req, cod_pr, desc, ref, um, qtde, unit, ipi, total, entr, q_nf = i

                            q_nf_float = valores_para_float(q_nf)

                            if q_nf_float:
                                pode_excluir = False
                                break

                        if pode_excluir:
                            msg = f'Tem certeza que deseja Excluir a Ordem de Compra Nº {num_oc}?'

                            if self.pergunta_confirmacao(msg):
                                cursor = conecta.cursor()
                                cursor.execute(
                                    f"SELECT oc.numero, oc.data, forn.registro, forn.razao, oc.frete, "
                                    f"oc.descontos, oc.obs "
                                    f"FROM ordemcompra as oc "
                                    f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.ID "
                                    f"where oc.entradasaida = 'E' "
                                    f"and oc.numero = {num_oc} "
                                    f"and oc.status = 'A';")
                                dados_oc_aberta = cursor.fetchall()

                                if dados_oc_aberta:
                                    id_oc = dados_oc[0][0]

                                    cursor = conecta.cursor()
                                    cursor.execute(
                                        f"SELECT ocprod.id, reqprod.id, reqprod.item, prod.codigo, prod.descricao, "
                                        f"COALESCE(prod.obs, ''), prod.unidade, ocprod.quantidade, "
                                        f"ocprod.unitario, ocprod.ipi, ocprod.dataentrega, "
                                        f"ocprod.produzido "
                                        f"FROM produtoordemcompra as ocprod "
                                        f"INNER JOIN produtoordemrequisicao as reqprod ON "
                                        f"ocprod.id_prod_req = reqprod.id "
                                        f"INNER JOIN produto as prod ON ocprod.produto = prod.ID "
                                        f"where ocprod.mestre = {id_oc};")
                                    dados_produtos = cursor.fetchall()

                                    if dados_produtos:
                                        for i in dados_produtos:
                                            id_oc_prod = i[0]
                                            id_req = i[1]

                                            cursor = conecta.cursor()
                                            cursor.execute(f"DELETE FROM produtoordemcompra WHERE ID = {id_oc_prod};")

                                            cursor = conecta.cursor()
                                            cursor.execute(f"UPDATE produtoordemrequisicao SET STATUS = 'A', "
                                                           f"PRODUZIDO = 0 WHERE id = {id_req};")

                                            conecta.commit()

                                    cursor = conecta.cursor()
                                    cursor.execute(f"DELETE FROM ordemcompra WHERE ID = {id_oc};")

                                    conecta.commit()

                                    self.limpar_tudo()
                        else:
                            self.mensagem_alerta(f'A ordem de compra não pode ser excluída em razão '
                                                 f'de Notas Fiscais de compra vinculadas!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def chave_de_ordenacao(self, item):
        try:
            chave1 = int(item[0])
            chave2 = int(item[1])

            chave_final = (chave1, chave2)

            return chave_final

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pinta_tabela(self):
        try:
            dados_tabela = extrair_tabela(self.table_Produtos_OC)

            for index, dados in enumerate(dados_tabela):
                num_req, item_req, cod_produto, descr, ref, um, qtde, unit, ipi, total, entrega, qtde_nf = dados

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, descricao, embalagem FROM produto where codigo = {cod_produto};")
                dados_produto = cursor.fetchall()
                ides, descr, embalagem = dados_produto[0]
                if embalagem == "SIM":
                    self.table_Produtos_OC.item(index, 4).setBackground(QColor(cor_amarelo))
                elif embalagem == "SER":
                    self.table_Produtos_OC.item(index, 4).setBackground(QColor(cor_amarelo))

                qtde_nf_float = valores_para_float(qtde_nf)

                if qtde_nf_float > 0.00:
                    num_colunas = len(dados_tabela[0])
                    for i in range(num_colunas):
                        self.table_Produtos_OC.item(index, i).setBackground(QColor(cor_cinza_claro))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_req(self):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT prodreq.numero, prodreq.item,  "
                           f"prod.codigo, prod.descricao as DESCRICAO, "
                           f"CASE prod.embalagem when 'SIM' then COALESCE(prodreq.referencia, '') "
                           f"else COALESCE(prod.obs, '') end as REFERENCIA, "
                           f"prod.unidade, prodreq.quantidade "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"WHERE prodreq.status = 'A' ORDER BY prodreq.numero;")
            extrair_req = cursor.fetchall()

            if extrair_req:
                lanca_tabela(self.table_Req_Abertas, extrair_req)
                self.lista_requisicoes = extrair_req

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def atualiza_mascara_frete(self):
        try:
            frete = self.line_Frete.text()

            frete_float = valores_para_float(frete)
            frete_2casas = ("%.2f" % frete_float)
            valor_string = valores_para_virgula(frete_2casas)
            valor_final = "R$ " + valor_string
            self.line_Frete.setText(valor_final)

            self.atualiza_valor_total()
            self.line_Desconto.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def atualiza_mascara_desconto(self):
        try:
            desconto = self.line_Desconto.text()

            desconto_float = valores_para_float(desconto)
            desconto_2casas = ("%.2f" % desconto_float)
            valor_string = valores_para_virgula(desconto_2casas)
            valor_final = "R$ " + valor_string
            self.line_Desconto.setText(valor_final)

            self.atualiza_valor_total()
            self.line_Obs.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def atualiza_valor_total(self):
        try:
            extrai_produtos = extrair_tabela(self.table_Produtos_OC)

            total_mercadorias = 0.00
            total_ipi = 0.00

            if extrai_produtos:
                for i in extrai_produtos:
                    num_req, item_req, cod_produto, descr, ref, um, qtde, unit, ipi, total, entrega, qtde_nf = i

                    qtde_float = valores_para_float(qtde)

                    unit_float = valores_para_float(unit)

                    ipi_float = valores_para_float(ipi)

                    total_ipi += qtde_float * (unit_float * (ipi_float / 100))
                    total_mercadorias += qtde_float * unit_float

            frete = self.line_Frete.text()
            frete_float = valores_para_float(frete)

            desconto = self.line_Desconto.text()
            desconto_float = valores_para_float(desconto)

            total_geral = total_ipi + total_mercadorias + frete_float - desconto_float

            total_ipi_2casas = ("%.2f" % total_ipi)
            valor_ipi_string = valores_para_virgula(total_ipi_2casas)
            valor_ipi_final = "R$ " + valor_ipi_string
            self.line_Total_Ipi.setText(valor_ipi_final)

            total_merc_2casas = ("%.2f" % total_mercadorias)
            valor_merc_string = valores_para_virgula(total_merc_2casas)
            valor_merc_final = "R$ " + valor_merc_string
            self.line_Total_Merc.setText(valor_merc_final)

            total_geral_2casas = ("%.2f" % total_geral)
            valor_geral_string = valores_para_virgula(total_geral_2casas)
            valor_geral_final = "R$ " + valor_geral_string
            self.line_Total_Geral.setText(valor_geral_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_produtos(self):
        try:
            self.line_Num_Req.clear()
            self.line_Item_Req.clear()

            self.line_Codigo.clear()
            self.line_Descricao.clear()
            self.line_Referencia.clear()
            self.line_UM.clear()

            self.line_Qtde.clear()
            self.line_Unit.clear()
            self.line_Ipi.clear()
            self.line_ValorTotal.clear()

            self.line_Referencia.setStyleSheet(f"background-color: {cor_branco};")

            self.definir_entrega()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_fornecedor(self):
        try:
            self.line_CodForn.clear()
            self.line_NomeForn.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_valores_obs(self):
        try:
            self.line_Frete.clear()
            self.line_Desconto.clear()
            self.line_Total_Ipi.setText("R$ 0,00")
            self.line_Total_Merc.setText("R$ 0,00")
            self.line_Total_Geral.setText("R$ 0,00")
            self.line_Obs.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpar_tudo(self):
        try:
            self.line_NumOC.clear()

            self.limpa_dados_produtos()
            self.limpa_dados_fornecedor()
            self.limpa_dados_valores_obs()

            self.table_Produtos_OC.setRowCount(0)
            self.table_Req_Abertas.setRowCount(0)

            definir_data_atual(self.date_Emissao)

            self.manipula_dados_req()
            self.line_NumOC.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def eventFilter(self, sources, event):
        try:
            nome_tabela = self.table_Produtos_OC

            if (event.type() == QtCore.QEvent.MouseButtonDblClick and event.buttons() == QtCore.Qt.LeftButton
                    and sources is nome_tabela.viewport()):

                item = nome_tabela.currentItem()

                if item:
                    data_str = self.date_Emissao.text()
                    data_emissao = datetime.strptime(data_str, '%d/%m/%Y')

                    extrai_recomendados = extrair_tabela(nome_tabela)
                    item_sel = extrai_recomendados[item.row()]

                    num_req, item_req, cod_prod, descr, ref, um, qtde, unit, ipi, total, entrega, qtde_nf = item_sel

                    qtde_float = valores_para_float(qtde)
                    q_nf_float = valores_para_float(qtde_nf)

                    if qtde_float == q_nf_float:
                        self.mensagem_alerta(f'O Código {cod_prod} não pode ser alterado pois já foi '
                                             f'encerrado e possui Notas Fiscais de compra vinculadas!')
                    else:
                        item_sel.append(data_emissao.strftime('%d/%m/%Y'))
                        self.abrir_tela_alteracao_prod(item_sel)

            return super(QMainWindow, self).eventFilter(sources, event)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_salvamento(self):
        try:
            numero_oc = self.line_NumOC.text()
            numero_oc_int = int(numero_oc)

            dados_tabela = extrair_tabela(self.table_Produtos_OC)

            if not numero_oc:
                self.mensagem_alerta("Não pode ser salvo sem informar o número da Ordem!")
            elif not dados_tabela:
                self.mensagem_alerta("Não pode ser salvo a Ordem sem produtos lançados na tabela!")
            else:
                excluir_prod_oc = self.salvar_excluir_produtos_oc(numero_oc_int)
                altera_prod_oc = self.verifica_alteracao_produtos_oc(numero_oc_int)
                altera_dados_oc = self.verifica_alteracao_oc(numero_oc_int)

                if excluir_prod_oc or altera_prod_oc or altera_dados_oc:
                    self.mensagem_alerta(f"Ordem de Compra Nº {numero_oc} foi alterada com sucesso!")
                    self.limpar_tudo()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_alteracao_oc(self, numero_oc):
        try:
            alteracao = False

            data_str = self.date_Emissao.text()
            emissao = datetime.strptime(data_str, '%d/%m/%Y').date()

            frete = self.line_Frete.text()
            frete_float = valores_para_float(frete)

            desconto = self.line_Desconto.text()
            desconto_float = valores_para_float(desconto)

            obs_n = self.line_Obs.text()
            obs = obs_n.upper()

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT oc.numero, oc.data, oc.fornecedor, oc.frete, oc.descontos, oc.obs "
                f"FROM ordemcompra as oc "
                f"where oc.entradasaida = 'E' "
                f"and oc.numero = {numero_oc} "
                f"and oc.status = 'A';")
            dados_oc_aberta = cursor.fetchall()

            if dados_oc_aberta:
                num_oc_b, emissao_b, id_forn_b, frete_b, descont_b, obs_b = dados_oc_aberta[0]

                campos_atualizados = []
                if emissao != emissao_b:
                    campos_atualizados.append(f"data = '{emissao}'")
                if frete_float != frete_b:
                    campos_atualizados.append(f"frete = {frete_float}")
                if desconto_float != descont_b:
                    campos_atualizados.append(f"descontos = {desconto_float}")
                if obs != obs_b:
                    campos_atualizados.append(f"obs = '{obs}'")

                if campos_atualizados:
                    campos_update = ", ".join(campos_atualizados)
                    cursor.execute(f"UPDATE ordemcompra SET {campos_update} "
                                   f"WHERE numero = {numero_oc} "
                                   f"AND entradasaida = 'E';")

                    conecta.commit()

                    alteracao = True

            return alteracao

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_alteracao_produtos_oc(self, numero_oc):
        try:
            alteracao = False

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT oc.id, oc.numero, oc.data, oc.fornecedor "
                f"FROM ordemcompra as oc "
                f"where oc.entradasaida = 'E' "
                f"and oc.numero = {numero_oc};")
            dados_oc = cursor.fetchall()

            if dados_oc:
                id_oc = dados_oc[0][0]

                produtos_tabela = extrair_tabela(self.table_Produtos_OC)

                for i in produtos_tabela:
                    num_req_t, item_t, cod_t, desc_t, ref_t, um_t, qtde_t, unit_t, ipi_t, tot_t, entr_t, qtde_nf_t = i

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT prodreq.id, prodreq.numero, prodreq.item,  "
                                   f"prod.codigo, prod.descricao as DESCRICAO, "
                                   f"CASE prod.embalagem when 'SIM' then COALESCE(prodreq.referencia, '') "
                                   f"else COALESCE(prod.obs, '') end as REFERENCIA, "
                                   f"prod.unidade, prodreq.quantidade "
                                   f"FROM produtoordemrequisicao as prodreq "
                                   f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                                   f"WHERE prodreq.numero = {num_req_t} "
                                   f"and prodreq.item = {item_t} "
                                   f"ORDER BY prodreq.numero;")
                    extrair_req = cursor.fetchall()

                    id_req = extrair_req[0][0]

                    cur = conecta.cursor()
                    cur.execute(f"SELECT prod.id, prod.descricao, COALESCE(prod.obs, '') , prod.unidade, "
                                f"prod.localizacao, prod.quantidade, conj.conjunto, prod.embalagem "
                                f"FROM produto as prod "
                                f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                                f"where codigo = {cod_t};")
                    detalhes_produto = cur.fetchall()
                    id_prod = detalhes_produto[0][0]

                    qtde_t_float = valores_para_float(qtde_t)
                    unit_t_float = valores_para_float(unit_t)
                    ipi_t_float = valores_para_float(ipi_t)
                    entr_t_dt = datetime.strptime(entr_t, '%d/%m/%Y').date()

                    cursor = conecta.cursor()
                    cursor.execute(
                        f"SELECT ocprod.id, reqprod.numero, reqprod.item, prod.codigo, prod.descricao, "
                        f"COALESCE(prod.obs, ''), prod.unidade, ocprod.quantidade, ocprod.unitario, "
                        f"ocprod.ipi, ocprod.dataentrega, "
                        f"ocprod.produzido "
                        f"FROM produtoordemcompra as ocprod "
                        f"INNER JOIN produtoordemrequisicao as reqprod ON ocprod.id_prod_req = reqprod.id "
                        f"INNER JOIN produto as prod ON ocprod.produto = prod.ID "
                        f"where ocprod.mestre = {id_oc} "
                        f"and reqprod.numero = {num_req_t} "
                        f"and reqprod.item = {item_t} "
                        f"and prod.codigo = {cod_t};")
                    prod_banco = cursor.fetchall()

                    if prod_banco:
                        id_oc_pord, n_req_b, item_b, cod_b, desc_b, ref_b, um_b, qtde_b, unit_b, ipi_b, \
                        entr_b, nf_b = prod_banco[0]

                        qtde_b_float = valores_para_float(qtde_b)
                        unit_b_float = valores_para_float(unit_b)
                        ipi_b_float = valores_para_float(ipi_b)

                        campos_atualizados = []

                        if qtde_t_float != qtde_b_float:
                            campos_atualizados.append(f"QUANTIDADE = {qtde_t_float}")
                        if unit_t_float != unit_b_float:
                            campos_atualizados.append(f"UNITARIO = {unit_t_float}")
                        if ipi_t_float != ipi_b_float:
                            campos_atualizados.append(f"IPI = {ipi_t_float}")
                        if entr_t_dt != entr_b:
                            campos_atualizados.append(f"DATAENTREGA = '{entr_t_dt}'")

                        if campos_atualizados:
                            campos_update = ", ".join(campos_atualizados)

                            cursor = conecta.cursor()
                            cursor.execute(f"UPDATE produtoordemcompra SET {campos_update} "
                                           f"WHERE id = {id_oc_pord};")

                            conecta.commit()

                            alteracao = True
                    else:
                        cursor = conecta.cursor()
                        cursor.execute(f"Insert into produtoordemcompra (ID, MESTRE, PRODUTO, QUANTIDADE, "
                                       f"UNITARIO, IPI, DATAENTREGA, NUMERO, CODIGO, PRODUZIDO, ID_PROD_REQ) "
                                       f"values (GEN_ID(GEN_PRODUTOORDEMCOMPRA_ID,1), {id_oc}, "
                                       f"{id_prod}, {qtde_t_float}, {unit_t_float}, {ipi_t_float}, "
                                       f"'{entr_t_dt}', {numero_oc}, '{int(cod_t)}', 0.0, {id_req});")

                        cursor = conecta.cursor()
                        cursor.execute(f"UPDATE produtoordemrequisicao SET STATUS = 'B', "
                                       f"PRODUZIDO = {qtde_t_float} WHERE id = {id_req};")

                        conecta.commit()

                        alteracao = True

            return alteracao

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_excluir_produtos_oc(self, numero_oc):
        try:
            alteracao = False

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT oc.id, oc.numero, oc.data, oc.fornecedor "
                f"FROM ordemcompra as oc "
                f"where oc.entradasaida = 'E' "
                f"and oc.numero = {numero_oc};")
            dados_oc = cursor.fetchall()

            id_oc = dados_oc[0][0]

            produtos_modificados = extrair_tabela(self.table_Produtos_OC)

            produtos_originais_tuple = []
            produtos_modificados_tuple = []

            for i in produtos_modificados:
                num_req_m, item_m, cod_m, desc_m, ref_m, um_m, qtde_m, unit_m, ipi_m, tot_m, entr_m, qtde_nf_m = i

                num_req_m_int = int(num_req_m)
                item_req_m_int = int(item_m)

                dados = (num_req_m_int, item_req_m_int, cod_m)
                produtos_modificados_tuple.append(dados)

            for ii in self.lista_produtos_oc:
                num_req_t, item_t, cod_t, desc_t, ref_t, um_t, qtde_t, unit_t, ipi_t, tot_t, entr_t, qtde_nf_t = ii

                dados = (num_req_t, item_t, cod_t)
                produtos_originais_tuple.append(dados)

            diferenca = set(produtos_originais_tuple).difference(set(produtos_modificados_tuple))
            if diferenca:
                for iii in diferenca:
                    num_req, item, cod = iii

                    cursor = conecta.cursor()
                    cursor.execute(
                        f"SELECT ocprod.id, reqprod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, ''), "
                        f"prod.unidade, ocprod.quantidade, ocprod.unitario, ocprod.ipi, ocprod.dataentrega, "
                        f"ocprod.produzido "
                        f"FROM produtoordemcompra as ocprod "
                        f"INNER JOIN produtoordemrequisicao as reqprod ON ocprod.id_prod_req = reqprod.id "
                        f"INNER JOIN produto as prod ON ocprod.produto = prod.ID "
                        f"where ocprod.mestre = {id_oc} "
                        f"and reqprod.numero = {num_req} "
                        f"and reqprod.item = {item} "
                        f"and prod.codigo = {cod};")
                    prod_banco = cursor.fetchall()

                    if prod_banco:
                        id_oc_pord = prod_banco[0][0]
                        id_req = prod_banco[0][1]

                        cursor = conecta.cursor()
                        cursor.execute(f"DELETE FROM produtoordemcompra WHERE ID = {id_oc_pord};")

                        cursor = conecta.cursor()
                        cursor.execute(f"UPDATE produtoordemrequisicao SET STATUS = 'A', "
                                       f"PRODUZIDO = 0 WHERE id = {id_req};")

                        conecta.commit()

                        alteracao = True

            return alteracao

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaOcAlterar()
    tela.show()
    qt.exec_()

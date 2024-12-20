import sys
from banco_dados.conexao import conecta
from forms.tela_nf_compra_incluir import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import lanca_tabela, layout_cabec_tab, extrair_tabela
from comandos.lines import definir_data_atual
from comandos.telas import tamanho_aplicacao, icone
from comandos.conversores import valores_para_float, valores_para_virgula
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import inspect
import os
import traceback
from datetime import datetime, date


class TelaNFCompraIncluir(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        self.processando = False

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_compra_sol.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_OC_Abertas)
        layout_cabec_tab(self.table_Produtos_NF)

        self.line_CodForn.editingFinished.connect(self.verifica_line_fornecedor)
        self.line_Codigo.editingFinished.connect(self.verifica_line_codigo)

        self.line_NCM.editingFinished.connect(self.passar_por_ncm)

        self.line_Qtde.editingFinished.connect(self.verifica_line_qtde)
        self.line_Ipi.editingFinished.connect(self.atualiza_mascara_ipi)

        self.line_Unit.editingFinished.connect(self.verifica_line_unit)
        self.btn_AdicionarProd.clicked.connect(self.verifica_line_unit)

        self.line_Frete.editingFinished.connect(self.atualiza_mascara_frete)
        self.line_Desconto.editingFinished.connect(self.atualiza_mascara_desconto)

        self.btn_Limpar.clicked.connect(self.limpar_tudo)

        self.btn_ExcluirItem.clicked.connect(self.excluir_item_tab_produtos)
        self.btn_ExcluirTudo.clicked.connect(self.excluir_tudo_tab_produtos)

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)

        self.table_OC_Abertas.viewport().installEventFilter(self)

        definir_data_atual(self.date_Entrega)
        self.definir_validador()
        self.definir_bloqueios()
        self.definir_combo_localestoque()

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

    def definir_bloqueios(self):
        try:
            self.line_NomeForn.setReadOnly(True)

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

    def definir_validador(self):
        try:
            validator = QtGui.QRegExpValidator(QtCore.QRegExp(r'\d+'), self.line_Num_NF)
            self.line_Num_NF.setValidator(validator)

            validator = QtGui.QRegExpValidator(QtCore.QRegExp(r'\d+'), self.line_Num_OC)
            self.line_Num_OC.setValidator(validator)

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

    def definir_combo_localestoque(self):
        try:
            self.combo_Local_Estoque.clear()

            tabela = []

            branco = ""
            tabela.append(branco)

            cur = conecta.cursor()
            cur.execute(f"SELECT id, nome FROM LOCALESTOQUE order by nome;")
            detalhes = cur.fetchall()

            for dadus in detalhes:
                ides, local = dadus
                msg = f"{ides} - {local}"
                tabela.append(msg)

            self.combo_Local_Estoque.addItems(tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_fornecedor(self):
        if not self.processando:
            try:
                self.processando = True

                cod_fornecedor = self.line_CodForn.text()
                if not cod_fornecedor:
                    self.mensagem_alerta('O campo "Cód. For.:" não pode estar vazio')
                    self.line_CodForn.clear()
                    self.line_CodForn.setFocus()
                elif int(cod_fornecedor) == 0:
                    self.mensagem_alerta('O campo "Cód. For.:" não pode ser "0"')
                    self.line_CodForn.clear()
                    self.line_CodForn.setFocus()
                else:
                    self.verifica_sql_fornecedor()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_sql_fornecedor(self):
        try:
            cod_fornecedor = self.line_CodForn.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {cod_fornecedor};")
            dados_fornecedor = cursor.fetchall()

            if not dados_fornecedor:
                self.mensagem_alerta('Este Código de Fornecedor não existe!')
                self.line_CodForn.clear()
                self.line_CodForn.setFocus()
            else:
                nom_fornecedor = dados_fornecedor[0][1].strip()
                self.line_NomeForn.setText(nom_fornecedor)
                self.line_Codigo.setFocus()

            self.oc_total_aberto_por_fornec(cod_fornecedor)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def oc_total_aberto_por_fornec(self, cod_fornecedor):
        try:
            tabela = []

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT oc.numero, prodoc.item, prodoc.codigo, "
                f"prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, (prodoc.quantidade - prodoc.produzido) "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                f"LEFT JOIN produtoordemSOLICITACAO as prodsol ON prodreq.id_prod_sol = prodsol.id "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"where oc.entradasaida = 'E' "
                f"AND oc.STATUS = 'A' "
                f"AND prodoc.produzido < prodoc.quantidade and forn.registro = {cod_fornecedor} "
                f"ORDER BY oc.numero;")
            dados_oc = cursor.fetchall()

            if dados_oc:
                for i in dados_oc:
                    oc, item_oc, cod, descr, ref, um, qtde = i

                    dados = (oc, item_oc, cod, descr, ref, um, qtde)
                    tabela.append(dados)

            if tabela:
                lanca_tabela(self.table_OC_Abertas, tabela)
            else:
                self.mensagem_alerta("Não existem compras pendentes para este Fornecedor!")
                self.line_CodForn.clear()
                self.line_NomeForn.clear()
                self.line_CodForn.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def eventFilter(self, sources, event):
        try:
            nome_tabela = self.table_OC_Abertas

            if (event.type() == QtCore.QEvent.MouseButtonDblClick and event.buttons() == QtCore.Qt.LeftButton
                    and sources is nome_tabela.viewport()):

                item = nome_tabela.currentItem()

                if item:
                    extrai_recomendados = extrair_tabela(nome_tabela)
                    item_sel = extrai_recomendados[item.row()]

                    num_oc, item_oc, cod_prod, descr, ref, um, qtde = item_sel

                    self.line_Codigo.setText(cod_prod)
                    self.line_Descricao.setText(descr)
                    self.line_Referencia.setText(ref)
                    self.line_UM.setText(um)

                    self.line_Num_OC.setText(num_oc)
                    self.line_Item_OC.setText(item_oc)

                    self.line_NCM.setFocus()

            return super(QMainWindow, self).eventFilter(sources, event)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_codigo(self):
        if not self.processando:
            try:
                self.processando = True

                cod_produto = self.line_Codigo.text()

                self.line_Descricao.clear()
                self.line_Referencia.clear()
                self.line_UM.clear()
                self.line_Qtde.clear()
                self.line_Unit.clear()
                self.line_Ipi.clear()
                self.line_ValorTotal.clear()

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
                self.lanca_dados_codigo(cod_produto)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_codigo(self, cod_produto):
        try:
            itens_encontrados = []

            dados_oc_abertas = extrair_tabela(self.table_OC_Abertas)
            for i in dados_oc_abertas:
                if cod_produto in i:
                    num_oc, item_oc, cod_prod, descr, ref, um, qtde = i
                    tt = (num_oc, item_oc, cod_prod, descr, ref, um, qtde)
                    itens_encontrados.append(tt)

            if not itens_encontrados:
                self.mensagem_alerta('Este produto não foi encontrado nas ordens de compra pendentes!')
                self.line_Codigo.clear()
                self.line_Codigo.setFocus()

            elif len(itens_encontrados) > 1:
                num_ocs, item_ocs, cods, descrs, refs, ums, qtdes = itens_encontrados[0]

                self.mensagem_alerta('Indique o Item da sequência da Ordem de Compra!')

                self.line_Codigo.setText(cods)
                self.line_Descricao.setText(descrs)
                self.line_Referencia.setText(refs)
                self.line_UM.setText(ums)
                self.line_Num_OC.setText(num_ocs)

                self.line_Item_OC.setFocus()

            else:
                num_ocs, item_ocs, cods, descrs, refs, ums, qtdes = itens_encontrados[0]

                self.line_Codigo.setText(cods)
                self.line_Descricao.setText(descrs)
                self.line_Referencia.setText(refs)
                self.line_UM.setText(ums)
                self.line_Num_OC.setText(num_ocs)
                self.line_Item_OC.setText(item_ocs)

                self.line_NCM.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_produtos(self):
        try:
            self.line_Codigo.clear()
            self.line_Descricao.clear()
            self.line_Referencia.clear()
            self.line_UM.clear()
            self.line_NCM.clear()
            self.line_Num_OC.clear()
            self.line_Item_OC.clear()
            self.definir_combo_localestoque()

            self.line_Qtde.clear()
            self.line_Unit.clear()
            self.line_Ipi.clear()
            self.line_ValorTotal.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_qtde(self):
        if not self.processando:
            try:
                self.processando = True

                qtde = self.line_Qtde.text()

                if len(qtde) == 0:
                    self.mensagem_alerta('O campo "Qtde:" não pode estar vazio')
                    self.line_Qtde.clear()
                    self.line_Qtde.setFocus()
                elif qtde == "0":
                    self.mensagem_alerta('O campo "Qtde:" não pode ser "0"')
                    self.line_Qtde.clear()
                    self.line_Qtde.setFocus()
                else:
                    qtde_com_virgula = valores_para_virgula(qtde)

                    self.line_Qtde.setText(qtde_com_virgula)
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

                self.line_Unit.setFocus()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def passar_por_ncm(self):
        if not self.processando:
            try:
                self.processando = True

                ncm = self.line_NCM.text()

                if not ncm:
                    self.mensagem_alerta('O campo "NCM" não pode estar vazio')

                self.combo_Local_Estoque.setFocus()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_line_unit(self):
        if not self.processando:
            try:
                self.processando = True

                unit = self.line_Unit.text()

                if len(unit) == 0:
                    self.mensagem_alerta('O campo "R$/Unid:" não pode estar vazio')
                    self.line_Unit.clear()
                    self.line_Unit.setFocus()
                elif unit == "0":
                    self.mensagem_alerta('O campo "R$/Unid:" não pode ser "0"')
                    self.line_Unit.clear()
                    self.line_Unit.setFocus()
                else:
                    self.atualiza_mascara_unit()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def atualiza_mascara_unit(self):
        try:
            print("passei")
            unit = self.line_Unit.text()

            unit_float = valores_para_float(unit)
            unit_2casas = ("%.4f" % unit_float)
            valor_string = valores_para_virgula(unit_2casas)
            valor_final = "R$ " + valor_string
            self.line_Unit.setText(valor_final)

            self.calcular_valor_total_prod()
            self.verifica_entrega()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

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
        try:
            self.verifica_line_unit()

            data_entrega_str = self.date_Entrega.text()
            try:
                data_entrega = datetime.strptime(data_entrega_str, '%d/%m/%Y')

                data_atual = datetime.combine(date.today(), datetime.min.time())

                if data_entrega < data_atual:
                    msg = f'Tem certeza que deseja lançar a Nota Fiscal com a data {data_entrega}?'

                    if self.pergunta_confirmacao(msg):
                        self.verifica_dados_completos()
                else:
                    self.verifica_dados_completos()
            except ValueError:
                self.mensagem_alerta(f'A data de entrega não está no formato correto '
                                     f'(dd/mm/aaaa)!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_dados_completos(self):
        try:
            num_nf = self.line_Num_NF.text()
            cod_fornecedor = self.line_CodForn.text()
            cod_produto = self.line_Codigo.text()

            ncm = self.line_NCM.text()

            local_est = self.combo_Local_Estoque.currentText()
            num_oc = self.line_Num_OC.text()
            item_oc = self.line_Item_OC.text()
            qtde = self.line_Qtde.text()
            unit = self.line_Unit.text()

            if not cod_fornecedor:
                self.mensagem_alerta('O campo "Cód. For." não pode estar vazio')
            elif not num_nf:
                self.mensagem_alerta('O campo "Nº NF" não pode estar vazio')
            elif not cod_produto:
                self.mensagem_alerta('O campo "Código" não pode estar vazio')
            elif not ncm:
                self.mensagem_alerta('O campo "NCM" não pode estar vazio')
            elif not local_est:
                self.mensagem_alerta('O campo "Local de Estoque" não pode estar vazio')
            elif not num_oc:
                self.mensagem_alerta('O campo "Nº OC" não pode estar vazio')
            elif not item_oc:
                self.mensagem_alerta('O campo "Item OC" não pode estar vazio')
            elif not qtde:
                self.mensagem_alerta('O campo "Qtde" não pode estar vazio')
            elif not unit:
                self.mensagem_alerta('O campo "R$/Unid" não pode estar vazio')
            else:
                ncm_sem_espaco = ncm.strip()

                if len(ncm_sem_espaco) != 8:
                    self.mensagem_alerta('O campo "NCM" não está no formato correto!')
                else:
                    self.verifica_ncm_valores()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_ncm_valores(self):
        try:
            self.calcular_valor_total_prod()

            cod_fornecedor = self.line_CodForn.text()

            ncm = self.line_NCM.text()
            ncm_sem_espaco = ncm.strip()
            ncm_4_digitos = ncm_sem_espaco[:4]

            num_oc = self.line_Num_OC.text()
            item_oc = self.line_Item_OC.text()

            qtde = self.line_Qtde.text()
            qtde_float = valores_para_float(qtde)

            unit = self.line_Unit.text()
            unit_float = valores_para_float(unit)

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT oc.numero, prodoc.item, prodoc.codigo, "
                f"prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, (prodoc.quantidade - prodoc.produzido), prod.ncm, prodoc.unitario, prodoc.ipi "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                f"LEFT JOIN produtoordemSOLICITACAO as prodsol ON prodreq.id_prod_sol = prodsol.id "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"where oc.entradasaida = 'E' "
                f"AND oc.STATUS = 'A' "
                f"AND prodoc.produzido < prodoc.quantidade "
                f"and forn.registro = {cod_fornecedor} "
                f"and oc.numero = {num_oc} "
                f"and prodoc.item = {item_oc} "
                f"ORDER BY oc.numero;")
            dados_oc = cursor.fetchall()

            if dados_oc:
                oc_bc, item_oc_bc, cod_bc, descr_bc, ref_bc, um_bc, qtde_bc, ncm_bc, unit_bc, ipi_bc = dados_oc[0]

                qtde_bc_float = valores_para_float(qtde_bc)
                unit_bc_float = valores_para_float(unit_bc)

                ncm_bc_sem_espaco = ncm_bc.strip()
                ncm_bc_num = ncm_bc_sem_espaco.replace(".", "").replace("-", "")
                ncm_bc_4_digitos = ncm_bc_num[:4]

                if qtde_float > qtde_bc_float:
                    self.mensagem_alerta("A quantidade da NF não está de acordo com o saldo da Ordem de Compra!")
                elif unit_float != unit_bc_float:
                    self.mensagem_alerta("O valor unitário não está de acordo com a Ordem de Compra!")
                elif ncm_4_digitos != ncm_bc_4_digitos:
                    self.mensagem_alerta("A NCM não está de acordo com o cadastro do produto!")
                else:
                    self.verifica_dados_ordem()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_dados_ordem(self):
        try:
            self.calcular_valor_total_prod()

            cod = self.line_Codigo.text()
            descr = self.line_Descricao.text()
            ref = self.line_Referencia.text()
            um = self.line_UM.text()
            ncm = self.line_NCM.text()
            est = self.combo_Local_Estoque.currentText()
            num_oc = self.line_Num_OC.text()
            item_oc = self.line_Item_OC.text()
            qtde = self.line_Qtde.text()
            unit = self.line_Unit.text()
            ipi = self.line_Ipi.text()
            total = self.line_ValorTotal.text()

            item_encontrado = self.verifica_item_oc(num_oc, item_oc)

            if item_encontrado:
                dados = [num_oc, item_oc, est, cod, descr, ref, um, ncm, qtde, unit, ipi, total]

                extrai_produtos = extrair_tabela(self.table_Produtos_NF)

                ja_existe = False
                for i in extrai_produtos:
                    num_oc_e, item_oc_e, est_e, cod_e, descr_e, ref_e, um_e, ncm_e, qtde_e, unit_e, ipi_e, total_e = i

                    if num_oc_e == num_oc and item_oc_e == item_oc:
                        ja_existe = True
                        break

                if not ja_existe:
                    extrai_produtos.append(dados)
                    if extrai_produtos:
                        lanca_tabela(self.table_Produtos_NF, extrai_produtos)

                    self.atualiza_valor_total()
                    self.limpa_dados_produtos()
                    self.line_Codigo.setFocus()

                    self.excluir_item_oc(item_encontrado)
                else:
                    self.mensagem_alerta(f'O item selecionado já está presente na tabela '
                                         f'"Produtos OC".')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_item_oc(self, num_oc_m, item_oc_m):
        try:
            itens_encontrados = []

            dados_oc_abertas = extrair_tabela(self.table_OC_Abertas)
            for indice, i in enumerate(dados_oc_abertas):
                num_oc, item_oc, cod, descr, ref, um, qtde = i

                if num_oc_m == num_oc and item_oc_m == item_oc:
                    tt = (indice, num_oc, item_oc, cod, descr, ref, um, qtde)
                    itens_encontrados.append(tt)

            if not itens_encontrados:
                self.mensagem_alerta('Número ou o Item da sequência da Ordem de Compra não encontrados!')
                self.line_Num_OC.setFocus()

                itens_encontrados = []

            elif len(itens_encontrados) > 1:
                self.mensagem_alerta('Indique o Item da sequência da Ordem de Compra!')

                itens_encontrados = []

            return itens_encontrados

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def atualiza_valor_total(self):
        try:
            extrai_produtos = extrair_tabela(self.table_Produtos_NF)

            total_mercadorias = 0.00
            total_ipi = 0.00

            if extrai_produtos:
                for i in extrai_produtos:
                    num_oc, item_oc, est, cod_produto, descr, ref, um, ncm, qtde, unit, ipi, total = i

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

    def excluir_item_oc(self, item_encontrado):
        try:
            linha_selecao = item_encontrado[0][0]

            nome_tabela = self.table_OC_Abertas

            extrai_recomendados = extrair_tabela(nome_tabela)
            if not extrai_recomendados:
                self.mensagem_alerta(f'A tabela "Tabela Ordens de Compras Abertas" está vazia!')
            else:
                if linha_selecao >= 0:
                    nome_tabela.removeRow(linha_selecao)

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
            self.line_Num_NF.clear()

            self.limpa_dados_produtos()
            self.limpa_dados_fornecedor()
            self.limpa_dados_valores_obs()

            self.table_Produtos_NF.setRowCount(0)
            self.table_OC_Abertas.setRowCount(0)

            definir_data_atual(self.date_Entrega)

            self.line_Num_NF.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def atualiza_mascara_frete(self):
        if not self.processando:
            try:
                self.processando = True

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

            finally:
                self.processando = False

    def atualiza_mascara_desconto(self):
        if not self.processando:
            try:
                self.processando = True

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

            finally:
                self.processando = False

    def excluir_item_tab_produtos(self):
        try:
            nome_tabela = self.table_Produtos_NF

            dados_tab = extrair_tabela(nome_tabela)
            if dados_tab:
                linha = nome_tabela.currentRow()
                if linha >= 0:
                    num_oc, item_oc, est, cod_pr, desc, ref, um, ncm, qtde, unit, ipi, total = dados_tab[linha]

                    cursor = conecta.cursor()
                    cursor.execute(
                        f"SELECT oc.numero, prodoc.item, prodoc.codigo, "
                        f"prod.descricao, COALESCE(prod.obs, ''), "
                        f"prod.unidade, (prodoc.quantidade - prodoc.produzido) "
                        f"FROM ordemcompra as oc "
                        f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                        f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                        f"LEFT JOIN produtoordemSOLICITACAO as prodsol ON prodreq.id_prod_sol = prodsol.id "
                        f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                        f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                        f"where oc.entradasaida = 'E' "
                        f"AND oc.STATUS = 'A' "
                        f"AND prodoc.produzido < prodoc.quantidade "
                        f"and oc.numero = {num_oc} "
                        f"and prodoc.item = {item_oc} "
                        f"and prodoc.codigo = {cod_pr} "
                        f"ORDER BY oc.numero;")
                    dados_oc = cursor.fetchall()

                    if dados_oc:
                        oc, item_oc, cod, descr, ref, um, qtde = dados_oc[0]

                        dados = [oc, item_oc, cod, descr, ref, um, qtde]
                        nome_tabela.removeRow(linha)

                        extrai_produtos = extrair_tabela(self.table_OC_Abertas)

                        ja_existe = False
                        for i in extrai_produtos:
                            num_oc_e, item_oc_e, cod_e, descr_e, ref_e, um_e, qtde_e = i

                            if num_oc_e == num_oc and item_oc_e == item_oc:
                                ja_existe = True
                                break

                        if not ja_existe:
                            extrai_produtos.append(dados)
                            if extrai_produtos:
                                extrai_produtos_ord = sorted(extrai_produtos, key=self.chave_de_ordenacao)
                                lanca_tabela(self.table_OC_Abertas, extrai_produtos_ord)

                        self.atualiza_valor_total()
                        self.limpa_dados_produtos()
                        self.line_Codigo.setFocus()

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

    def excluir_tudo_tab_produtos(self):
        try:
            nome_tabela = self.table_Produtos_NF

            dados_tab = extrair_tabela(nome_tabela)
            if dados_tab:
                nome_tabela.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_salvamento(self):
        try:
            cod_fornecedor = self.line_CodForn.text()
            nome_fornecedor = self.line_NomeForn.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {cod_fornecedor};")
            dados_fornecedor = cursor.fetchall()

            dados_tabela = extrair_tabela(self.table_Produtos_NF)

            if not dados_fornecedor:
                self.mensagem_alerta(f'O Fornecedor {nome_fornecedor} não está cadastrado!')
            elif not dados_tabela:
                self.mensagem_alerta(f'A tabela "Produtos Nota Fiscal" não pode estar vazia!')

            else:
                for i in dados_tabela:
                    num_oc, item_oc, est, cod_produto, descr, ref, um, ncm, qtde, unit, ipi, total = i

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaNFCompraIncluir()
    tela.show()
    qt.exec_()

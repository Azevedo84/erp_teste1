import sys
from banco_dados.conexao import conecta
from forms.tela_oc_incluir import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.lines import definir_data_atual
from comandos.cores import cor_amarelo, cor_branco
from comandos.telas import tamanho_aplicacao, icone
from comandos.conversores import valores_para_float, valores_para_virgula
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtGui import QColor
from datetime import datetime, date, timedelta
import inspect
import os
import traceback


class TelaOcIncluir(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_compra_sol.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Req_Abertas)
        layout_cabec_tab(self.table_Produtos_OC)

        self.lista_requisicoes = []

        self.line_NumOC.editingFinished.connect(self.verifica_line_oc)

        self.date_Emissao.editingFinished.connect(self.verifica_emissao)
        self.line_CodForn.editingFinished.connect(self.verifica_line_fornecedor)

        self.line_Codigo.editingFinished.connect(self.verifica_line_codigo)
        self.line_Qtde.editingFinished.connect(self.verifica_line_qtde)
        self.line_Ipi.editingFinished.connect(self.atualiza_mascara_ipi)
        self.line_Unit.editingFinished.connect(self.verifica_line_unit)

        self.line_ValorTotal.editingFinished.connect(self.verifica_entrega)
        self.btn_AdicionarProd.clicked.connect(self.verifica_entrega)

        self.line_Frete.editingFinished.connect(self.atualiza_mascara_frete)
        self.line_Desconto.editingFinished.connect(self.atualiza_mascara_desconto)
        self.date_Entrega.editingFinished.connect(self.verifica_entrega)

        self.processando = False

        self.btn_ExcluirItem.clicked.connect(self.excluir_item_tab_produtos)
        self.btn_ExcluirTudo.clicked.connect(self.excluir_tudo_tab_produtos)

        self.btn_Limpar.clicked.connect(self.limpar_tudo)

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)

        self.definir_validador()
        self.definir_bloqueios()
        self.definir_entrega()
        definir_data_atual(self.date_Emissao)
        self.manipula_dados_req()

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

    def definir_entrega(self):
        try:
            data_hoje = date.today()
            data_entrega = data_hoje + timedelta(days=14)
            self.date_Entrega.setDate(data_entrega)

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

    def verifica_line_oc(self):
        if not self.processando:
            try:
                self.processando = True

                num_oc = self.line_NumOC.text()
                if len(num_oc) == 0:
                    self.mensagem_alerta('O campo "Nº OC:" não pode estar vazio')
                    self.line_NumOC.clear()
                    self.line_NumOC.setFocus()
                elif int(num_oc) == 0:
                    self.mensagem_alerta('O campo "Nº OC:" não pode ser "0"')
                    self.line_NumOC.clear()
                    self.line_NumOC.setFocus()
                else:
                    self.verifica_sql_oc()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_sql_oc(self):
        try:
            num_oc = self.line_NumOC.text()

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT oc.numero, oc.data, oc.status FROM ordemcompra as oc "
                f"where oc.entradasaida = 'E' and oc.numero = {num_oc};")
            dados_oc = cursor.fetchall()

            if dados_oc:
                self.mensagem_alerta('Este número de OC já existe!')
                self.line_NumOC.clear()
                self.line_NumOC.setFocus()
            else:
                self.date_Emissao.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_emissao(self):
        if not self.processando:
            try:
                self.processando = True

                data_emissao_str = self.date_Emissao.text()

                try:
                    data_emissao = datetime.strptime(data_emissao_str, '%d/%m/%Y')

                    data_atual = date.today()

                    if data_emissao.year < data_atual.year:
                        self.mensagem_alerta(f'O ano da emissão é inferior '
                                             f'a {data_atual.year}!')
                        self.line_CodForn.setFocus()

                    elif data_emissao.year == data_atual.year:
                        if data_emissao.date() <= data_atual:
                            self.line_CodForn.setFocus()

                        else:
                            self.mensagem_alerta(f'A data de emissão é maior que a atual!')
                            self.line_CodForn.setFocus()

                    else:
                        self.mensagem_alerta(f'A data de emissão é maior que a atual!')

                except ValueError:
                    msg = f'A data de emissão não está no formato correto (dd/mm/aaaa)!'
                    print(msg)
                    self.mensagem_alerta(msg)

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_line_fornecedor(self):
        if not self.processando:
            try:
                self.processando = True

                cod_fornecedor = self.line_CodForn.text()
                if len(cod_fornecedor) == 0:
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
                    self.line_Unit.setFocus()

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
            unit = self.line_Unit.text()

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
                        self.mensagem_alerta(f'A data de entrega não pode ser menor que a '
                                             f'atual!')
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

            item_encontrado = self.verifica_item_requisicao(num_req, item_req)

            if item_encontrado:
                dados = [num_req, item_req, cod_produto, descr, ref, um, qtde, unit, ipi, total, entrega]

                extrai_produtos = extrair_tabela(self.table_Produtos_OC)

                ja_existe = False
                for i in extrai_produtos:
                    num_req_e, item_req_e, cod_produto_e, descr_e, ref_e, um_e, qtde_e, unit_e, ipi_e, total_e, \
                    entrega_e = i

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
                    self.mensagem_alerta(f'O item selecionado já está presente na tabela '
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
                    num_req, item_req, cod_pr, desc, ref, um, qtde, unit, ipi, total, entr = dados_tab[linha]

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
            nome_tabela = self.table_Produtos_OC

            dados_tab = extrair_tabela(nome_tabela)
            if not dados_tab:
                self.mensagem_alerta(f'A tabela "Produtos Ordem de Compra" está vazia!')
            else:
                nome_tabela.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pinta_tabela(self):
        try:
            dados_tabela = extrair_tabela(self.table_Produtos_OC)

            for index, dados in enumerate(dados_tabela):
                num_req, item_req, cod_produto, descr, ref, um, qtde, unit, ipi, total, entrega = dados

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, descricao, embalagem FROM produto where codigo = {cod_produto};")
                dados_produto = cursor.fetchall()
                ides, descr, embalagem = dados_produto[0]
                if embalagem == "SIM":
                    self.table_Produtos_OC.item(index, 4).setBackground(QColor(cor_amarelo))
                elif embalagem == "SER":
                    self.table_Produtos_OC.item(index, 4).setBackground(QColor(cor_amarelo))

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

    def atualiza_valor_total(self):
        try:
            extrai_produtos = extrair_tabela(self.table_Produtos_OC)

            total_mercadorias = 0.00
            total_ipi = 0.00

            if extrai_produtos:
                for i in extrai_produtos:
                    num_req, item_req, cod_produto, descr, ref, um, qtde, unit, ipi, total, entrega = i

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

    def verifica_salvamento(self):
        try:
            cod_fornecedor = self.line_CodForn.text()
            nome_fornecedor = self.line_NomeForn.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {cod_fornecedor};")
            dados_fornecedor = cursor.fetchall()

            if not dados_fornecedor:
                self.mensagem_alerta(f'O Fornecedor {nome_fornecedor} não está cadastrado!')
            else:
                testar_erros = 0
                dados_alterados = extrair_tabela(self.table_Produtos_OC)
                for itens in dados_alterados:
                    num_req, item_req, cod_produto, descr, ref, um, qtde, unit, ipi, total, entrega = itens

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, descricao FROM produto where codigo = {cod_produto};")
                    dados_produto = cursor.fetchall()
                    if not dados_produto:
                        self.mensagem_alerta(f'O produto {descr} não está cadastrado')
                        testar_erros = testar_erros + 1
                        break

                if testar_erros == 0:
                    self.salvar_ordem()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_ordem(self):
        try:
            cod_fornecedor = self.line_CodForn.text()

            numero_oc = self.line_NumOC.text()
            numero_oc_int = int(numero_oc)

            emissao_oc = self.date_Emissao.date()
            data_emi = emissao_oc.toString("yyyy-MM-dd")

            frete = self.line_Frete.text()
            frete_oc_float = valores_para_float(frete)

            desconto = self.line_Desconto.text()
            desconto_oc_float = valores_para_float(desconto)

            obs = self.line_Obs.text()
            obs_m = obs.upper()
            obs_m = obs_m.encode("utf-8").decode("utf-8")
            print(len(obs_m))

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {cod_fornecedor};")
            dados_fornecedor = cursor.fetchall()
            id_fornecedor, razao = dados_fornecedor[0]

            cursor = conecta.cursor()
            cursor.execute("select GEN_ID(GEN_ORDEMCOMPRA_ID,0) from rdb$database;")
            ultimo_oc0 = cursor.fetchall()
            ultimo_oc1 = ultimo_oc0[0]
            ultimo_oc = int(ultimo_oc1[0]) + 1

            cursor = conecta.cursor()
            cursor.execute(f"Insert into ordemcompra "
                           f"(ID, ENTRADASAIDA, NUMERO, DATA, STATUS, FORNECEDOR, LOCALESTOQUE, FRETE, DESCONTOS, OBS) "
                           f"values (GEN_ID(GEN_ORDEMCOMPRA_ID,1), "
                           f"'E', {numero_oc_int}, '{data_emi}', 'A', {id_fornecedor}, '1', {frete_oc_float}, "
                           f"{desconto_oc_float}, '{obs_m}');")

            dados_alterados = extrair_tabela(self.table_Produtos_OC)

            for indice, itens in enumerate(dados_alterados, start=1):
                num_req, item_req, cod_produto, descr, ref, um, qtde, unit, ipi, total, entrega = itens

                codigo_int = int(cod_produto)

                entrega_prod = datetime.strptime(entrega, '%d/%m/%Y').date()

                qtde_item_float = valores_para_float(qtde)

                valor_unit_float = valores_para_float(unit)

                ipi_item_float = valores_para_float(ipi)

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, descricao FROM produto where codigo = {codigo_int};")
                dados_produto = cursor.fetchall()

                id_produto, descricao = dados_produto[0]

                cursor = conecta.cursor()
                cursor.execute(f"SELECT prodreq.id, prodreq.numero, prodreq.item,  "
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

                id_req = extrair_req[0][0]

                cursor = conecta.cursor()
                cursor.execute(f"Insert into produtoordemcompra (ID, MESTRE, ITEM, PRODUTO, QUANTIDADE, UNITARIO, "
                               f"IPI, DATAENTREGA, NUMERO, CODIGO, PRODUZIDO, ID_PROD_REQ) "
                               f"values (GEN_ID(GEN_PRODUTOORDEMCOMPRA_ID,1), {ultimo_oc}, "
                               f"{indice}, {id_produto}, {qtde_item_float}, {valor_unit_float}, {ipi_item_float}, "
                               f"'{entrega_prod}', {numero_oc_int}, '{codigo_int}', 0.0, {id_req});")

                cursor = conecta.cursor()
                cursor.execute(f"UPDATE produtoordemrequisicao SET STATUS = 'B', "
                               f"PRODUZIDO = {qtde_item_float} WHERE id = {id_req};")

            conecta.commit()

            self.mensagem_alerta(f'Ordem de Compra foi lançada com sucesso!')

            self.limpar_tudo()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaOcIncluir()
    tela.show()
    qt.exec_()

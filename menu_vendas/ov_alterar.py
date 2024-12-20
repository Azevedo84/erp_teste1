import sys
from banco_dados.conexao import conecta
from forms.tela_ov_alterar import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.telas import tamanho_aplicacao, icone
from comandos.conversores import valores_para_float
from comandos.cores import cor_cinza_claro
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtGui import QColor
from datetime import date, timedelta, datetime
import inspect
import os
import traceback


class TelaOvAlterar(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_vendas.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_OV)
        layout_cabec_tab(self.table_PI_Aberto)

        self.campos_bloqueados()

        self.line_Num_OV.editingFinished.connect(self.verifica_num_ov_e_id_cliente)

        self.combo_Cliente.activated.connect(self.verifica_num_ov_e_id_cliente)

        self.line_Codigo_Manu.editingFinished.connect(self.verifica_line_codigo_manual)
        self.line_Qtde_Manu.editingFinished.connect(self.mascara_qtde_manual)
        self.line_Unit_Manu.editingFinished.connect(self.mascara_valor_unitario_manual)
        self.line_Ipi_Manu.editingFinished.connect(self.mascara_ipi_manual)
        self.line_Num_PI_Manu.editingFinished.connect(self.verifica_dados_completos_manual)

        self.line_Frete.editingFinished.connect(self.mascara_frete)
        self.line_Desconto.editingFinished.connect(self.mascara_desconto)

        self.btn_Adicionar.clicked.connect(self.verifica_dados_completos_manual)
        self.btn_ExcluirTudo.clicked.connect(self.excluir_tudo_ov)
        self.btn_ExcluirItem.clicked.connect(self.excluir_item_ov)

        self.btn_Limpar.clicked.connect(self.limpa_tudo)

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)

        self.definir_emissao()
        self.definir_entrega()
        self.manipula_dados_pi()
        self.line_Num_OV.setFocus()

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

    def campos_bloqueados(self):
        try:
            self.line_Descricao_Manu.setReadOnly(True)
            self.line_Referencia_Manu.setReadOnly(True)
            self.line_UM_Manu.setReadOnly(True)
            self.line_Valor_Manu.setReadOnly(True)

            self.line_Total_Ipi.setReadOnly(True)
            self.line_Total_Merc.setReadOnly(True)
            self.line_Total_Geral.setReadOnly(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_pi(self):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.id, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prodint.qtde, prodint.data_previsao "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A';")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    num_ped, cliente, cod, descr, ref, um, qtde, entrega = i

                    data = f'{entrega.day}/{entrega.month}/{entrega.year}'

                    dados = (num_ped, cliente, cod, descr, um, qtde, data)

                    tabela_nova.append(dados)
            if tabela_nova:
                lanca_tabela(self.table_PI_Aberto, tabela_nova)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_emissao(self):
        try:
            data_hoje = date.today()
            self.date_Emissao.setDate(data_hoje)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_entrega(self):
        try:
            data_entrega = date.today() + timedelta(7)
            self.date_Entrega_Manu.setDate(data_entrega)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_num_ov_e_id_cliente(self):
        if not self.processando:
            try:
                self.processando = True

                num_ov = self.line_Num_OV.text()
                cliente = self.combo_Cliente.currentText()

                if num_ov and cliente:
                    self.verifica_num_ov()
                    self.verifica_id_cliente()
                else:
                    if num_ov:
                        self.combo_Cliente.setFocus()
                    if cliente:
                        self.line_Num_OV.setFocus()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_num_ov(self):
        try:
            self.processando = True

            num_ov = self.line_Num_OV.text()

            if num_ov:
                if int(num_ov) == 0:
                    self.mensagem_alerta('O campo "Nº OV" não pode ser "0"!')
                    self.limpa_tudo()
                    self.line_Num_OV.clear()
                else:
                    self.combo_Cliente.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_id_cliente(self):
        try:
            cliente = self.combo_Cliente.currentText()

            num_ov = self.line_Num_OV.text()

            clientetete = cliente.find(" - ")
            id_cliente = cliente[:clientetete]

            self.limpa_dados_manual()
            self.limpa_dados_pedido()
            self.limpa_tabelas()
            self.manipula_dados_pi()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.numero, oc.data, oc.status FROM ordemcompra as oc "
                           f"where oc.entradasaida = 'S' "
                           f"and oc.numero = {num_ov} "
                           f"and cliente = {id_cliente};")
            dados_ov = cursor.fetchall()
            if dados_ov:
                self.lanca_dados_ov()
                self.lanca_produtos_ov()
            else:
                self.mensagem_alerta("Este número de OV não existe!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_ov(self):
        try:
            num_ov = self.line_Num_OV.text()

            cliente = self.combo_Cliente.currentText()
            clientetete = cliente.find(" - ")
            id_cliente = cliente[:clientetete]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.data, oc.cliente, oc.status, COALESCE(oc.descontos, '') as descon, "
                           f"COALESCE(oc.frete, '') as frete, "
                           f"COALESCE(oc.obs, '') as obs "
                           f"FROM ordemcompra as oc "
                           f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                           f"where oc.entradasaida = 'S' "
                           f"and oc.numero = {num_ov} "
                           f"and oc.cliente = {id_cliente};")
            dados_interno = cursor.fetchall()

            emissao, cliente, status, desconto, frete, obs = dados_interno[0]

            self.date_Emissao.setDate(emissao)

            if status == "A":
                self.label_Status.setText("ABERTO")
                self.liberar_campos_pi()
            elif status == "B":
                self.label_Status.setText("BAIXADO")
                self.bloquear_campos_pi()
            else:
                self.label_Status.setText(status)

            self.line_Desconto.setText(str(desconto))
            self.line_Frete.setText(str(frete))
            self.line_Obs.setText(obs)

            item_count = self.combo_Cliente.count()
            for i in range(item_count):
                item_text = self.combo_Cliente.itemText(i)

                if item_text:
                    clientetete = item_text.find(" - ")
                    id_cliente = int(item_text[:clientetete])

                    if cliente == int(id_cliente):
                        self.combo_Cliente.setCurrentText(item_text)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_produtos_ov(self):
        try:
            num_ov = self.line_Num_OV.text()

            cliente = self.combo_Cliente.currentText()
            clientetete = cliente.find(" - ")
            id_cliente = cliente[:clientetete]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE(prodoc.id_pedido, '') as pi, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prodoc.quantidade, prodoc.unitario, COALESCE(prodoc.ipi, '') as ipi, "
                           f"prodoc.dataentrega, prodoc.produzido "
                           f"FROM PRODUTOORDEMCOMPRA as prodoc "
                           f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                           f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                           f"where oc.entradasaida = 'S' "
                           f"and oc.numero = {num_ov} "
                           f"and oc.cliente = {id_cliente}")
            dados_produtos = cursor.fetchall()

            extrai_produtos = extrair_tabela(self.table_OV)

            if dados_produtos:
                for i in dados_produtos:
                    num_pi, cod, descr, ref, um, qtde, unit, ipi, entrega, qtde_ent = i

                    data = f'{entrega.day}/{entrega.month}/{entrega.year}'

                    if qtde and unit:
                        qtde_float = valores_para_float(qtde)
                        unit_float = valores_para_float(unit)

                        total = str("%.2f" % (qtde_float * unit_float))
                        if "." in total:
                            string_com_virgula = total.replace('.', ',')
                            valor_final = string_com_virgula
                        else:
                            valor_final = total

                    else:
                        valor_final = "0,00"

                    dados = [num_pi, cod, descr, ref, um, qtde, unit, ipi, valor_final, data, qtde_ent]

                    ja_existe = False
                    for iii in extrai_produtos:
                        num_pi_e, cod_e, des_e, ref_e, um_e, qtde_e, unit_e, ipi_e, tot_e, ent_e, qtde_ent_e = iii

                        if cod == cod_e:
                            ja_existe = True
                            break

                    if not ja_existe:
                        extrai_produtos.append(dados)

            if extrai_produtos:
                lanca_tabela(self.table_OV, extrai_produtos, zebra=False)
                self.soma_totais()
                self.remover_itens_pi()
                self.pintar_tabela_ov()

                self.line_Codigo_Manu.setFocus()
                self.limpa_dados_manual()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def bloquear_campos_pi(self):
        try:
            self.combo_Cliente.setEnabled(False)
            self.line_Obs.setReadOnly(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def liberar_campos_pi(self):
        try:
            self.combo_Cliente.setEnabled(True)
            self.line_Obs.setReadOnly(False)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tudo(self):
        self.limpa_num_pi_e_cliente()
        self.limpa_dados_manual()
        self.limpa_dados_pedido()
        self.limpa_tabelas()
        self.label_Status.setText("Status")

    def limpa_num_pi_e_cliente(self):
        try:
            self.line_Num_OV.clear()
            self.combo_Cliente.setCurrentText("")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_pedido(self):
        try:
            self.line_Obs.clear()
            self.definir_emissao()
            self.line_Frete.clear()
            self.line_Total_Ipi.clear()
            self.line_Total_Merc.clear()
            self.line_Desconto.clear()
            self.line_Total_Geral.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_manual(self):
        try:
            self.line_Codigo_Manu.clear()
            self.line_Referencia_Manu.clear()
            self.line_Descricao_Manu.clear()
            self.line_UM_Manu.clear()
            self.line_Qtde_Manu.clear()
            self.line_Ipi_Manu.clear()
            self.line_Unit_Manu.clear()
            self.line_Valor_Manu.clear()
            self.line_Num_PI_Manu.clear()
            self.definir_entrega()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tabelas(self):
        try:
            self.table_OV.setRowCount(0)
            self.table_PI_Aberto.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pintar_tabela_ov(self):
        try:
            qtable_widget = self.table_OV

            dados_tabela = extrair_tabela(self.table_OV)

            for index, dados in enumerate(dados_tabela):
                num_pi, cod, desc, ref, um, qtde, unit, ipi, tot, ent, qtde_ent = dados

                qtde_float = valores_para_float(qtde)
                qtde_ent_float = valores_para_float(qtde_ent)

                if qtde_float == qtde_ent_float:
                    qtable_widget.item(index, 0).setBackground(QColor(cor_cinza_claro))
                    qtable_widget.item(index, 1).setBackground(QColor(cor_cinza_claro))
                    qtable_widget.item(index, 2).setBackground(QColor(cor_cinza_claro))
                    qtable_widget.item(index, 3).setBackground(QColor(cor_cinza_claro))
                    qtable_widget.item(index, 4).setBackground(QColor(cor_cinza_claro))
                    qtable_widget.item(index, 5).setBackground(QColor(cor_cinza_claro))
                    qtable_widget.item(index, 6).setBackground(QColor(cor_cinza_claro))
                    qtable_widget.item(index, 7).setBackground(QColor(cor_cinza_claro))
                    qtable_widget.item(index, 8).setBackground(QColor(cor_cinza_claro))
                    qtable_widget.item(index, 9).setBackground(QColor(cor_cinza_claro))
                    qtable_widget.item(index, 10).setBackground(QColor(cor_cinza_claro))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_codigo_manual(self):
        if not self.processando:
            try:
                self.processando = True

                codigo_produto = self.line_Codigo_Manu.text()

                num_oc = self.line_Num_OV.text()

                if not num_oc:
                    self.mensagem_alerta('O campo "Nº OC" não pode estar vazio!')
                    self.line_Num_OV.setFocus()
                    self.line_Codigo_Manu.clear()
                else:
                    cliente = self.combo_Cliente.currentText()
                    if not cliente:
                        self.mensagem_alerta('O campo "Cliente" não pode estar vazio!')
                        self.combo_Cliente.setFocus()
                        self.line_Codigo_Manu.clear()
                    else:
                        if not codigo_produto:
                            self.line_Codigo_Manu.clear()
                        elif int(codigo_produto) == 0:
                            self.mensagem_alerta('O campo "Código" não pode ser "0"!')
                            self.line_Codigo_Manu.clear()
                        else:
                            status = self.label_Status.text()
                            if status == "BAIXADO":
                                self.mensagem_alerta("Este pedido já está encerrado e não pode ser "
                                                     "adicionado produtos!")
                                self.line_Codigo_Manu.clear()
                            else:
                                self.verifica_sql_produto_manual()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_sql_produto_manual(self):
        try:
            codigo_produto = self.line_Codigo_Manu.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade, localizacao, quantidade "
                           f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cursor.fetchall()
            if not detalhes_produto:
                self.mensagem_alerta('Este código de produto não existe!')
                self.line_Codigo_Manu.clear()
            else:
                self.verifica_pi_manual()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_pi_manual(self):
        try:
            codigo_produto = self.line_Codigo_Manu.text()
            cliente = self.combo_Cliente.currentText()
            clientetete = cliente.find(" - ") + 3
            nome_cliente = cliente[clientetete:]

            tem_produto_pi = 0

            extrai_produtos_pi = extrair_tabela(self.table_PI_Aberto)
            if extrai_produtos_pi:
                for i in extrai_produtos_pi:
                    num_pi, cliente, cod, descr, um, qtde, entrega = i

                    if nome_cliente == cliente and cod == codigo_produto:
                        print(nome_cliente, cliente, cod, codigo_produto)
                        self.line_Num_PI_Manu.setText(str(num_pi))
                        self.line_Qtde_Manu.setText(str(qtde))
                        self.lanca_dados_produto_manual()

                        tem_produto_pi += 1

                        break

            if not tem_produto_pi:
                self.mensagem_alerta("Não existe Pedido Interno deste produto para este cliente!")
                self.line_Codigo_Manu.clear()
                self.line_Codigo_Manu.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_produto_manual(self):
        try:
            codigo_produto = self.line_Codigo_Manu.text()

            cur = conecta.cursor()
            cur.execute(f"SELECT descricao, COALESCE(obs, '') as ref, unidade "
                        f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            descr, ref, um = detalhes_produto[0]

            self.line_Descricao_Manu.setText(descr)
            self.line_Referencia_Manu.setText(ref)
            self.line_UM_Manu.setText(um)
            self.line_Qtde_Manu.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_qtde_manual(self):
        if not self.processando:
            try:
                self.processando = True

                qtde = self.line_Qtde_Manu.text()
                if qtde:
                    qtde_float = valores_para_float(qtde)
                    qtde_str = str("%.3f" % qtde_float)

                    if "." in qtde_str:
                        string_com_virgula = qtde_str.replace('.', ',')
                        valor_final = string_com_virgula
                    else:
                        valor_final = qtde_str

                    self.line_Qtde_Manu.setText(valor_final)

                    self.line_Unit_Manu.setFocus()
                    self.calculo_valor_total_manual()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def mascara_valor_unitario_manual(self):
        try:
            unit = self.line_Unit_Manu.text()
            if unit:
                unit_float = valores_para_float(unit)
                unit_str = str("%.4f" % unit_float)

                if "." in unit_str:
                    string_com_virgula = unit_str.replace('.', ',')
                    valor_final = string_com_virgula
                else:
                    valor_final = unit_str

                self.line_Unit_Manu.setText(valor_final)

                self.line_Ipi_Manu.setFocus()
                self.calculo_valor_total_manual()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_ipi_manual(self):
        try:
            ipi = self.line_Ipi_Manu.text()
            if ipi:
                ipi_float = valores_para_float(ipi)
                ipi_str = str("%.2f" % ipi_float)

                if "." in ipi_str:
                    string_com_virgula = ipi_str.replace('.', ',')
                    valor_final = string_com_virgula
                else:
                    valor_final = ipi_str

                self.line_Ipi_Manu.setText(valor_final)

                self.date_Entrega_Manu.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_frete(self):
        try:
            self.soma_totais()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_desconto(self):
        try:
            self.soma_totais()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_valor_total_manual(self):
        try:
            qtde = self.line_Qtde_Manu.text()
            unit = self.line_Unit_Manu.text()

            if qtde and unit:
                qtde_float = valores_para_float(qtde)
                unit_float = valores_para_float(unit)

                total = str("%.2f" % (qtde_float * unit_float))
                if "." in total:
                    string_com_virgula = total.replace('.', ',')
                    valor_final = string_com_virgula
                else:
                    valor_final = total
                self.line_Valor_Manu.setText(valor_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_dados_completos_manual(self):
        if not self.processando:
            try:
                self.processando = True

                num_ov = self.line_Num_OV.text()
                cliente = self.combo_Cliente.currentText()
                cod_produto = self.line_Codigo_Manu.text()
                qtde = self.line_Qtde_Manu.text()
                unit = self.line_Unit_Manu.text()
                num_pi = self.line_Num_PI_Manu.text()

                if not num_ov:
                    self.mensagem_alerta('O campo "Nº OV" não pode estar vazio')
                    self.line_Num_OV.setFocus()
                elif not cliente:
                    self.mensagem_alerta('O campo "Cliente" não pode estar vazio')
                    self.combo_Cliente.setFocus()
                elif not cod_produto:
                    self.mensagem_alerta('O campo "Código:" não pode estar vazio')
                    self.line_Codigo_Manu.setFocus()
                elif not num_pi:
                    self.mensagem_alerta('O campo "Nº PI" não pode estar vazio')
                    self.line_Num_PI_Manu.setFocus()
                elif not qtde:
                    self.mensagem_alerta('O campo "Qtde:" não pode estar vazio')
                    self.line_Qtde_Manu.setFocus()
                elif not unit:
                    self.mensagem_alerta('O campo "R$/Unid:" não pode estar vazio')
                    self.line_Unit_Manu.setFocus()
                else:
                    self.manipula_dados_completos_manual()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def manipula_dados_completos_manual(self):
        try:
            cod = self.line_Codigo_Manu.text()
            descr = self.line_Descricao_Manu.text()
            ref = self.line_Referencia_Manu.text()
            um = self.line_UM_Manu.text()

            qtde = self.line_Qtde_Manu.text()
            unit = self.line_Unit_Manu.text()
            ipi_ini = self.line_Ipi_Manu.text()
            if ipi_ini:
                ipi = ipi_ini
            else:
                ipi = ""

            entrega = self.date_Entrega_Manu.text()
            num_pi = self.line_Num_PI_Manu.text()
            total = self.line_Valor_Manu.text()

            dados = [num_pi, cod, descr, ref, um, qtde, unit, ipi, total, entrega, "0"]

            extrai_produtos = extrair_tabela(self.table_OV)

            ja_existe = False
            for i in extrai_produtos:
                id_req_e, cod_e, descr_e, ref_e, um_e, qtde_e, unit_e, ipi_e, total_e, ent_e, qtde_ent_e = i

                if cod == cod_e:
                    ja_existe = True
                    break

            if not ja_existe:
                extrai_produtos.append(dados)
                lanca_tabela(self.table_OV, extrai_produtos, zebra=False)
                self.soma_totais()
                self.remover_itens_pi()
                self.pintar_tabela_ov()

                self.line_Codigo_Manu.setFocus()
                self.limpa_dados_manual()
            else:
                self.mensagem_alerta(f'O item selecionado já está presente na tabela "Produtos Ordem de Venda".')
                self.limpa_dados_manual()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def soma_totais(self):
        try:
            soma_total_geral = 0.00
            soma_valores = 0.00
            soma_com_ipi = 0.00

            extrai_produtos = extrair_tabela(self.table_OV)
            if extrai_produtos:
                for i in extrai_produtos:
                    qtde = i[5]
                    unit = i[6]
                    ipi = i[7]
                    valor = i[8]

                    qtde_float = valores_para_float(qtde)
                    unit_float = valores_para_float(unit)

                    if ipi:
                        ipi_float = valores_para_float(ipi)
                    else:
                        ipi_float = 0.00

                    valor_com_ipi = qtde_float * ((unit_float * (ipi_float / 100)) + unit_float)

                    total_float = valores_para_float(valor)

                    soma_valores += total_float
                    soma_com_ipi += valor_com_ipi

            frete = self.line_Frete.text()
            if frete:
                valor_frete = valores_para_float(frete)
            else:
                valor_frete = 0

            desconto = self.line_Desconto.text()
            if desconto:
                valor_desconto = valores_para_float(desconto)
            else:
                valor_desconto = 0

            soma_total_geral += soma_com_ipi + valor_frete
            soma_total_geral -= valor_desconto

            total_mercadorias = str("%.2f" % soma_valores)
            self.line_Total_Merc.setText(total_mercadorias)

            total_geral = str("%.2f" % soma_total_geral)
            self.line_Total_Geral.setText(total_geral)

            total_ipi = soma_com_ipi - soma_valores
            total_ipi_final = str("%.2f" % total_ipi)
            self.line_Total_Ipi.setText(total_ipi_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def remover_itens_pi(self):
        try:
            extrai_produtos = extrair_tabela(self.table_PI_Aberto)
            if extrai_produtos:
                pedido_excluir_texto = self.line_Num_PI_Manu.text()
                produto_excluir_texto = self.line_Codigo_Manu.text()

                dados_removidos = [linha for linha in extrai_produtos
                                   if not (linha[0] == pedido_excluir_texto and linha[2] == produto_excluir_texto)]

                if dados_removidos:
                    self.table_PI_Aberto.setRowCount(0)
                    lanca_tabela(self.table_PI_Aberto, dados_removidos)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def adicionar_itens_pi(self, id_cliente, codigo_produto):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.id, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prodint.qtde, prodint.data_previsao "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A' "
                           f"and prod.codigo = {codigo_produto} "
                           f"and ped.id_cliente = {id_cliente};")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    num_ped, cliente, cod, descr, ref, um, qtde, entrega = i

                    data = f'{entrega.day}/{entrega.month}/{entrega.year}'

                    dados = (num_ped, cliente, cod, descr, ref, um, qtde, data)

                    tabela_nova.append(dados)
            if tabela_nova:
                lanca_tabela(self.table_PI_Aberto, tabela_nova)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_tudo_ov(self):
        try:
            itens_encerrados = 0

            tabela_nova = extrair_tabela(self.table_PI_Aberto)

            extrai_tabela_ov = extrair_tabela(self.table_OV)
            if not extrai_tabela_ov:
                self.mensagem_alerta(f'A tabela "Produtos Ordem de Venda" está vazia!')
            else:
                for i in extrai_tabela_ov:
                    qtde_ent = float(i[10])

                    if qtde_ent:
                        itens_encerrados += 1

            if itens_encerrados:
                self.mensagem_alerta('Produtos com notas fiscais vinculadas, não podem ser excluídos!')
            else:
                for dados in extrai_tabela_ov:
                    num_pi, cod, desc, ref, um, qtde, unit, ipi, tot, ent, qtde_ent = dados

                    if num_pi:
                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT ped.id, cli.razao, prod.codigo, prod.descricao, "
                                       f"COALESCE(prod.obs, '') as obs, "
                                       f"prod.unidade, prodint.qtde, prodint.data_previsao "
                                       f"FROM PRODUTOPEDIDOINTERNO as prodint "
                                       f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                                       f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                                       f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                                       f"where prod.codigo = {cod} "
                                       f"and ped.id = {num_pi};")
                        dados_interno = cursor.fetchall()

                        if dados_interno:
                            for i in dados_interno:
                                num_ped, cliente, cod, descr, ref, um, qtde, entrega = i

                                data = f'{entrega.day}/{entrega.month}/{entrega.year}'

                                dados = [num_ped, cliente, cod, descr, ref, um, qtde, data]

                                tabela_nova.append(dados)

                if tabela_nova:
                    tab_ordenada = sorted(tabela_nova, key=lambda x: str(x[0]))
                    lanca_tabela(self.table_PI_Aberto, tab_ordenada)

                self.table_OV.setRowCount(0)
                self.soma_totais()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_item_ov(self):
        try:
            tabela_nova = extrair_tabela(self.table_PI_Aberto)

            extrai_recomendados = extrair_tabela(self.table_OV)
            if not extrai_recomendados:
                self.mensagem_alerta(f'A tabela "Produtos Ordem de Venda" está vazia!')
            else:
                linha_selecao = self.table_OV.currentRow()
                if linha_selecao >= 0:
                    dados = []
                    num_colunas = self.table_OV.columnCount()
                    for coluna in range(num_colunas):
                        item = self.table_OV.item(linha_selecao, coluna)
                        dados.append(item.text() if item else '')
                    num_pi = dados[0]
                    cod_prod = dados[1]
                    qtde_entr = float(dados[10])

                    if qtde_entr > 0:
                        self.mensagem_alerta("Este produto tem notas fiscais vinculadas e não pode ser excluído!")
                    else:
                        if num_pi:
                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT ped.id, cli.razao, prod.codigo, prod.descricao, "
                                           f"COALESCE(prod.obs, '') as obs, "
                                           f"prod.unidade, prodint.qtde, prodint.data_previsao "
                                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                                           f"where prod.codigo = {cod_prod} "
                                           f"and ped.id = {num_pi};")
                            dados_interno = cursor.fetchall()

                            if dados_interno:
                                for i in dados_interno:
                                    num_ped, cliente, cod, descr, ref, um, qtde, entrega = i

                                    data = f'{entrega.day}/{entrega.month}/{entrega.year}'

                                    dados = [num_ped, cliente, cod, descr, ref, um, qtde, data]

                                    tabela_nova.append(dados)
                        if tabela_nova:
                            tab_ordenada = sorted(tabela_nova, key=lambda x: str(x[0]))
                            lanca_tabela(self.table_PI_Aberto, tab_ordenada)

                        self.table_OV.removeRow(linha_selecao)

                        self.soma_totais()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def gerar_dados_ov_banco(self):
        try:
            num_ov = self.line_Num_OV.text()

            cliente = self.combo_Cliente.currentText()
            clientetete = cliente.find(" - ")
            id_cliente = cliente[:clientetete]

            tabela = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE(prodoc.id_pedido, '') as pi, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, prodoc.quantidade, prodoc.unitario, COALESCE(prodoc.ipi, '') as ipi, "
                           f"prodoc.dataentrega, prodoc.produzido "
                           f"FROM PRODUTOORDEMCOMPRA as prodoc "
                           f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                           f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                           f"where oc.entradasaida = 'S' "
                           f"and oc.numero = {num_ov} "
                           f"and oc.cliente = {id_cliente}")
            dados_produtos = cursor.fetchall()
            if dados_produtos:
                for i in dados_produtos:
                    num_pi, cod, descr, ref, um, qtde, unit, ipi, entrega, qtde_ent = i

                    data = f'{entrega.day}/{entrega.month}/{entrega.year}'

                    dados = (num_pi, cod, descr, ref, um, qtde, unit, ipi, data, qtde_ent)
                    tabela.append(dados)

            return tabela

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_salvamento(self):
        try:
            num_ov = self.line_Num_OV.text()
            cliente = self.combo_Cliente.currentText()
            tabela_produtos = extrair_tabela(self.table_OV)

            if not num_ov:
                self.mensagem_alerta('O campo "Nº OV" não pode estar vazio!')
            elif not cliente:
                self.mensagem_alerta('O campo "Cliente" não pode estar vazio!')
            elif not tabela_produtos:
                self.mensagem_alerta('A tabela "Produtos Ordem de Venda" não pode estar vazia"')
            else:
                num_ov = self.line_Num_OV.text()

                cliente = self.combo_Cliente.currentText()
                clientetete = cliente.find(" - ")
                id_cliente = cliente[:clientetete]

                cursor = conecta.cursor()
                cursor.execute(f"SELECT oc.cliente, oc.status, COALESCE(oc.descontos, '') as descon, "
                               f"COALESCE(oc.frete, '') as frete, "
                               f"COALESCE(oc.obs, '') as obs "
                               f"FROM ordemcompra as oc "
                               f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                               f"where oc.entradasaida = 'S' "
                               f"and oc.numero = {num_ov} "
                               f"and oc.cliente = {id_cliente};")
                dados_interno = cursor.fetchall()

                if dados_interno:
                    cliente_b, status_b, desconto_b, frete_b, obs_b = dados_interno[0]
                    if status_b == "B":
                        self.mensagem_alerta("Esta Ordem de Venda (OV) já está encerrada e não pode ser alterada!")
                    else:
                        self.salvar_ov_existente()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_ov_existente(self):
        try:
            qtde_salvamentos = 0

            num_ov = self.line_Num_OV.text()
            num_ov_int = int(num_ov)

            cliente = self.combo_Cliente.currentText()
            clientetete = cliente.find(" - ")
            id_cliente = cliente[:clientetete]

            frete = self.line_Frete.text()
            frete_oc_float = valores_para_float(frete)

            desconto = self.line_Desconto.text()
            desconto_oc_float = valores_para_float(desconto)

            obs = self.line_Obs.text()
            obs_m = obs.upper()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.id, oc.cliente, COALESCE(oc.descontos, '') as descon, "
                           f"COALESCE(oc.frete, '') as frete, "
                           f"COALESCE(oc.obs, '') as obs "
                           f"FROM ordemcompra as oc "
                           f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                           f"where oc.entradasaida = 'S' "
                           f"and oc.numero = {num_ov} "
                           f"and oc.cliente = {id_cliente};")
            dados_interno = cursor.fetchall()

            id_mestre, id_cliente_b, desconto_b, frete_b, obs_b = dados_interno[0]

            desconto_b_float_e = valores_para_float(desconto_b)
            frete_b_float_e = valores_para_float(frete_b)

            if desconto_oc_float != desconto_b_float_e \
                    or frete_oc_float != frete_b_float_e \
                    or obs_m != obs_b:
                print("ALTERAÇÃO NOS DADOS")

                cursor = conecta.cursor()
                cursor.execute(f"UPDATE ordemcompra SET descontos = '{desconto_oc_float}', "
                               f"frete = {frete_oc_float}, "
                               f"obs = '{obs_m}' "
                               f"where id = {id_mestre};")

                qtde_salvamentos += 1

            dados_tabela = extrair_tabela(self.table_OV)
            dados_banco = self.gerar_dados_ov_banco()

            lista_pi_b = []
            lista_pi_a = []

            for item_tab in dados_tabela:
                n_pi_tab, cod_tab, des_tab, ref_tab, um_tab, qtde_tab, unit_tab, ipi_tab, tot_tab, ent_tab, \
                qtde_ent_tab = item_tab

                qtde_tab_float = valores_para_float(qtde_tab)
                qtde_ent_tab_float = valores_para_float(qtde_ent_tab)
                unit_tab_float = valores_para_float(unit_tab)
                ipi_tab_float = valores_para_float(ipi_tab)

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, codigo, embalagem FROM produto where codigo = '{cod_tab}';")
                dados_produto = cursor.fetchall()
                id_produto, codigo, embalagem = dados_produto[0]

                date_entr_tab = datetime.strptime(ent_tab, '%d/%m/%Y').date()
                data_entr_tab_certa = date_entr_tab.strftime('%Y-%m-%d')

                for item_banc in dados_banco:
                    n_pi_bc, cod_bc, des_bc, ref_bc, um_bc, qtde_bc, unit_bc, ipi_bc, ent_bc, qtde_ent_bc = item_banc

                    if cod_tab == cod_bc:
                        if qtde_ent_tab_float > 0:
                            print("ALTERAÇÃO PRODUTO NÃO ENCERRADO", cod_tab, des_tab)

                            cursor = conecta.cursor()
                            cursor.execute(f"UPDATE produtoordemcompra "
                                           f"SET dataentrega = '{data_entr_tab_certa}', "
                                           f"quantidade = {qtde_tab_float}, "
                                           f"unitario = {unit_tab_float}, "
                                           f"ipi = {ipi_tab_float}, "
                                           f"id_pedido = {n_pi_tab} "
                                           f"WHERE mestre = {id_mestre} "
                                           f"and produto = {id_produto};")

                            if n_pi_tab:
                                print("ATUALIZA STATUS PRODUTO PI PARA B", cod_tab, des_tab)

                                cursor = conecta.cursor()
                                cursor.execute(f"UPDATE produtopedidointerno SET STATUS = 'B' "
                                               f"WHERE id_produto = {id_produto} "
                                               f"and id_pedidointerno = {n_pi_tab};")

                                lista_pi_b.append(n_pi_tab)

                            qtde_salvamentos += 1
                        break
                else:
                    print("INSERT PRODUTO NOVO", cod_tab, des_tab)

                    cursor = conecta.cursor()
                    cursor.execute(f"Insert into produtoordemcompra (ID, MESTRE, PRODUTO, QUANTIDADE, UNITARIO, "
                                   f"IPI, DATAENTREGA, NUMERO, CODIGO, PRODUZIDO, ID_PEDIDO) "
                                   f"values (GEN_ID(GEN_PRODUTOORDEMCOMPRA_ID,1), {id_mestre}, "
                                   f"{id_produto}, {qtde_tab_float}, {unit_tab_float}, {ipi_tab_float}, "
                                   f"'{data_entr_tab_certa}', {num_ov_int}, '{cod_tab}', 0.0, {n_pi_tab});")

                    if n_pi_tab:
                        print("ATUALIZA STATUS PRODUTO PI PARA B", cod_tab, des_tab)

                        cursor = conecta.cursor()
                        cursor.execute(f"UPDATE produtopedidointerno SET STATUS = 'B' "
                                       f"WHERE id_produto = {id_produto} "
                                       f"and id_pedidointerno = {n_pi_tab};")

                        lista_pi_b.append(n_pi_tab)

                    qtde_salvamentos += 1

            conjunto_tabela = set([item[1] for item in dados_tabela])
            conjunto_banco = set([item[1] for item in dados_banco])

            codigos_faltantes = conjunto_banco - conjunto_tabela

            if codigos_faltantes:
                for i in codigos_faltantes:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT prodoc.id, COALESCE(prodoc.id_pedido, '') as pi, prod.id, "
                                   f"prod.descricao, "
                                   f"COALESCE(prod.obs, '') as obs, "
                                   f"prod.unidade, prodoc.quantidade, prodoc.unitario, "
                                   f"COALESCE(prodoc.ipi, '') as ipi, "
                                   f"prodoc.dataentrega, prodoc.produzido "
                                   f"FROM PRODUTOORDEMCOMPRA as prodoc "
                                   f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                                   f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                                   f"where oc.entradasaida = 'S' "
                                   f"and oc.numero = {num_ov} "
                                   f"and oc.cliente = {id_cliente} "
                                   f"and prod.codigo = {i}")
                    dados_produtos = cursor.fetchall()
                    id_ocprod, num_pi, id_prods, descr, ref, um, qtde, unit, ipi, entrega, qtde_ent = dados_produtos[0]

                    qtde_ent_float = valores_para_float(qtde_ent)

                    if qtde_ent_float == 0:
                        print("DELETE PRODUTO", i)

                        cursor = conecta.cursor()
                        cursor.execute(f"DELETE FROM PRODUTOORDEMCOMPRA WHERE id = {id_ocprod};")

                        if num_pi:
                            print("ATUALIZA STATUS PRODUTO PI PARA A", i)

                            cursor = conecta.cursor()
                            cursor.execute(f"UPDATE produtopedidointerno SET STATUS = 'A' "
                                           f"WHERE id_produto = {id_prods} "
                                           f"and id_pedidointerno = {num_pi};")

                            lista_pi_a.append(num_pi)
                        qtde_salvamentos += 1

            if qtde_salvamentos:
                print("entrei no commit")
                conecta.commit()
                print("salvado")
                self.mensagem_alerta(f'Ordem de Compra foi alterada com sucesso!')

                if lista_pi_b:
                    lista_sem_duplicatas = list(set(lista_pi_b))
                    for iii in lista_sem_duplicatas:
                        numero_pi = iii[0]

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT ped.id, cli.razao, prod.codigo, prod.descricao, "
                                       f"COALESCE(prod.obs, '') as obs, "
                                       f"prod.unidade, prodint.qtde, prodint.data_previsao "
                                       f"FROM PRODUTOPEDIDOINTERNO as prodint "
                                       f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                                       f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                                       f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                                       f"where prodint.status = 'A' and ped.id = {numero_pi};")
                        dados_interno = cursor.fetchall()

                        if not dados_interno:
                            print("ATUALIZAR STATUS DO PI INTEIRO para B")
                            """
                            cursor = conecta.cursor()
                            cursor.execute(f"UPDATE pedidointerno SET STATUS = 'B' "
                                           f"WHERE id = {numero_pi};")

                            conecta.commit()
                            """

                if lista_pi_a:
                    lista_sem_duplicatasa = list(set(lista_pi_a))
                    for iiia in lista_sem_duplicatasa:
                        numero_pi_a = iiia[0]

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT ped.id, cli.razao, prod.codigo, prod.descricao, "
                                       f"COALESCE(prod.obs, '') as obs, "
                                       f"prod.unidade, prodint.qtde, prodint.data_previsao "
                                       f"FROM PRODUTOPEDIDOINTERNO as prodint "
                                       f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                                       f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                                       f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                                       f"where prodint.status = 'A' and ped.id = {numero_pi_a};")
                        dados_interno_a = cursor.fetchall()

                        if dados_interno_a:
                            print("ATUALIZAR STATUS DO PI INTEIRO para A")
                            """
                            cursor = conecta.cursor()
                            cursor.execute(f"UPDATE pedidointerno SET STATUS = 'B' "
                                           f"WHERE id = {numero_pi};")

                            conecta.commit()
                            """

                self.limpa_tabelas()
                self.limpa_tudo()
                self.manipula_dados_pi()
                self.line_Num_OV.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaOvAlterar()
    tela.show()
    qt.exec_()

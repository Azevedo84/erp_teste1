import sys
from banco_dados.conexao import conecta
from forms.tela_vendas_status import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import lanca_tabela
from comandos.telas import tamanho_aplicacao, icone, cor_widget_cab
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import inspect
import os
import traceback


class TelaVendasStatus(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_vendas.png")
        tamanho_aplicacao(self)
        cor_widget_cab(self.widget_cabecalho)

        self.btn_Consultar_PI.clicked.connect(self.verifica_filtro_pi)
        self.dados_pi_aberto()

        self.btn_Consultar_OV.clicked.connect(self.verifica_filtro_ov)
        self.dados_ov_aberto()

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

    def verifica_filtro_pi(self):
        try:
            cliente = self.combo_Cliente_PI.currentText()
            if cliente == "TODOS":
                id_cliente = ""
            else:
                clientetete = cliente.find(" - ")
                id_cliente = cliente[:clientetete]

            if self.check_Aberto_PI.isChecked():
                aberto = True
            else:
                aberto = False

            if self.check_Baixado_PI.isChecked():
                baixado = True
            else:
                baixado = False

            if id_cliente:
                if aberto and baixado:
                    self.table_PI.setRowCount(0)
                    self.dados_pi_todos_com_cliente(id_cliente)
                elif aberto and not baixado:
                    self.table_PI.setRowCount(0)
                    self.dados_pi_aberto_com_cliente(id_cliente)
                elif not aberto and baixado:
                    self.table_PI.setRowCount(0)
                    self.dados_pi_baixado_com_cliente(id_cliente)
            else:
                if aberto and baixado:
                    self.table_PI.setRowCount(0)
                    self.dados_pi_aberto()
                    self.mensagem_alerta("Defina um cliente para filtrar os pedidos!")
                    self.combo_Cliente_PI.setFocus()
                    self.check_Aberto_PI.setChecked(True)
                    self.check_Baixado_PI.setChecked(False)
                elif aberto and not baixado:
                    self.table_PI.setRowCount(0)
                    self.dados_pi_aberto()
                elif not aberto and baixado:
                    self.table_PI.setRowCount(0)
                    self.dados_pi_aberto()
                    self.mensagem_alerta("Defina um cliente para filtrar os pedidos!")
                    self.combo_Cliente_PI.setFocus()
                    self.check_Aberto_PI.setChecked(True)
                    self.check_Baixado_PI.setChecked(False)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def dados_pi_aberto(self):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.emissao, ped.id, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as ref, prod.unidade, prodint.qtde, "
                           f"prodint.data_previsao, ped.solicitante, COALESCE(ped.num_req_cliente, '') as req, "
                           f"ped.nome_pc, prodint.status, COALESCE(ped.obs, '') as obs "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A';")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    emissao, num_pi, clie, cod, descr, ref, um, qtde, entrega, solic, num_req, pc, status, obs = i

                    emi = f'{emissao.day}/{emissao.month}/{emissao.year}'
                    entreg = f'{entrega.day}/{entrega.month}/{entrega.year}'

                    dados = (emi, num_pi, clie, cod, descr, ref, um, qtde, entreg, solic, num_req, pc, status, obs)
                    tabela_nova.append(dados)
            if tabela_nova:
                lanca_tabela(self.table_PI, tabela_nova)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def dados_pi_baixado_com_cliente(self, id_cliente):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.emissao, ped.id, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as ref, prod.unidade, prodint.qtde, "
                           f"prodint.data_previsao, ped.solicitante, COALESCE(ped.num_req_cliente, '') as req, "
                           f"ped.nome_pc, prodint.status, COALESCE(ped.obs, '') as obs "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'B' and ped.id_cliente = {id_cliente};")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    emissao, num_pi, clie, cod, descr, ref, um, qtde, entrega, solic, num_req, pc, status, obs = i

                    emi = f'{emissao.day}/{emissao.month}/{emissao.year}'
                    entreg = f'{entrega.day}/{entrega.month}/{entrega.year}'

                    dados = (emi, num_pi, clie, cod, descr, ref, um, qtde, entreg, solic, num_req, pc, status, obs)
                    tabela_nova.append(dados)
            if tabela_nova:
                lista_de_listas_ordenada = sorted(tabela_nova, key=lambda x: x[1])
                lanca_tabela(self.table_PI, lista_de_listas_ordenada)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def dados_pi_aberto_com_cliente(self, id_cliente):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.emissao, ped.id, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as ref, prod.unidade, prodint.qtde, "
                           f"prodint.data_previsao, ped.solicitante, COALESCE(ped.num_req_cliente, '') as req, "
                           f"ped.nome_pc, prodint.status, COALESCE(ped.obs, '') as obs "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A' and ped.id_cliente = {id_cliente};")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    emissao, num_pi, clie, cod, descr, ref, um, qtde, entrega, solic, num_req, pc, status, obs = i

                    emi = f'{emissao.day}/{emissao.month}/{emissao.year}'
                    entreg = f'{entrega.day}/{entrega.month}/{entrega.year}'

                    dados = (emi, num_pi, clie, cod, descr, ref, um, qtde, entreg, solic, num_req, pc, status, obs)
                    tabela_nova.append(dados)
            if tabela_nova:
                lista_de_listas_ordenada = sorted(tabela_nova, key=lambda x: x[1])
                lanca_tabela(self.table_PI, lista_de_listas_ordenada)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def dados_pi_todos_com_cliente(self, id_cliente):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.emissao, ped.id, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as ref, prod.unidade, prodint.qtde, "
                           f"prodint.data_previsao, ped.solicitante, COALESCE(ped.num_req_cliente, '') as req, "
                           f"ped.nome_pc, prodint.status, COALESCE(ped.obs, '') as obs "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where ped.id_cliente = {id_cliente};")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    emissao, num_pi, clie, cod, descr, ref, um, qtde, entrega, solic, num_req, pc, status, obs = i

                    emi = f'{emissao.day}/{emissao.month}/{emissao.year}'
                    entreg = f'{entrega.day}/{entrega.month}/{entrega.year}'

                    dados = (emi, num_pi, clie, cod, descr, ref, um, qtde, entreg, solic, num_req, pc, status, obs)
                    tabela_nova.append(dados)
            if tabela_nova:
                lista_de_listas_ordenada = sorted(tabela_nova, key=lambda x: x[1])
                lanca_tabela(self.table_PI, lista_de_listas_ordenada)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_filtro_ov(self):
        try:
            cliente = self.combo_Cliente_OV.currentText()
            if cliente == "TODOS":
                id_cliente = ""
            else:
                clientetete = cliente.find(" - ")
                id_cliente = cliente[:clientetete]

            if self.check_Aberto_OV.isChecked():
                aberto = True
            else:
                aberto = False

            if self.check_Baixado_OV.isChecked():
                baixado = True
            else:
                baixado = False

            if id_cliente:
                if aberto and baixado:
                    self.table_OV.setRowCount(0)
                    self.dados_ov_todos_com_cliente(id_cliente)
                elif aberto and not baixado:
                    self.table_OV.setRowCount(0)
                    self.dados_ov_aberto_com_cliente(id_cliente)
                elif not aberto and baixado:
                    self.table_OV.setRowCount(0)
                    self.dados_ov_baixado_com_cliente(id_cliente)
            else:
                if aberto and baixado:
                    self.table_OV.setRowCount(0)
                    self.dados_ov_aberto()
                    self.mensagem_alerta("Defina um cliente para filtrar os pedidos!")
                    self.combo_Cliente_OV.setFocus()
                    self.check_Aberto_OV.setChecked(True)
                    self.check_Baixado_OV.setChecked(False)
                elif aberto and not baixado:
                    self.table_OV.setRowCount(0)
                    self.dados_ov_aberto()
                elif not aberto and baixado:
                    self.table_OV.setRowCount(0)
                    self.dados_ov_aberto()
                    self.mensagem_alerta("Defina um cliente para filtrar os pedidos!")
                    self.combo_Cliente_OV.setFocus()
                    self.check_Aberto_OV.setChecked(True)
                    self.check_Baixado_OV.setChecked(False)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def dados_ov_aberto(self):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.data, oc.numero, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as ref, prod.unidade, prodoc.quantidade, "
                           f"prodoc.produzido, COALESCE(prodoc.id_pedido, '') as pedi, "
                           f"COALESCE(ped.num_req_cliente, '') as req, "
                           f"COALESCE(oc.obs, '') as obs "
                           f"FROM PRODUTOORDEMCOMPRA as prodoc "
                           f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                           f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                           f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                           f"LEFT JOIN pedidointerno as ped ON prodoc.id_pedido = ped.id "
                           f"where prodoc.quantidade > prodoc.produzido "
                           f"and oc.status = 'A' "
                           f"and oc.entradasaida = 'S';")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    emissao, num_ov, clie, cod, descr, ref, um, qtde, qtde_entr, nume_pi, num_req, obs = i

                    emi = f'{emissao.day}/{emissao.month}/{emissao.year}'

                    dados = (emi, num_ov, clie, cod, descr, ref, um, qtde, qtde_entr, nume_pi, num_req, obs)
                    tabela_nova.append(dados)
            if tabela_nova:
                lanca_tabela(self.table_OV, tabela_nova)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def dados_ov_baixado_com_cliente(self, id_cliente):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.data, oc.numero, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as ref, prod.unidade, prodoc.quantidade, "
                           f"prodoc.produzido, COALESCE(prodoc.id_pedido, '') as pedi, "
                           f"COALESCE(ped.num_req_cliente, '') as req, "
                           f"COALESCE(oc.obs, '') as obs "
                           f"FROM PRODUTOORDEMCOMPRA as prodoc "
                           f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                           f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                           f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                           f"LEFT JOIN pedidointerno as ped ON prodoc.id_pedido = ped.id "
                           f"where oc.status = 'B' "
                           f"and oc.entradasaida = 'S'"
                           f"and oc.cliente = {id_cliente};")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    emissao, num_ov, clie, cod, descr, ref, um, qtde, qtde_entr, nume_pi, num_req, obs = i

                    emi = f'{emissao.day}/{emissao.month}/{emissao.year}'

                    dados = (emi, num_ov, clie, cod, descr, ref, um, qtde, qtde_entr, nume_pi, num_req, obs)
                    tabela_nova.append(dados)
            if tabela_nova:
                lista_de_listas_ordenada = sorted(tabela_nova, key=lambda x: x[1])
                lanca_tabela(self.table_OV, lista_de_listas_ordenada)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def dados_ov_aberto_com_cliente(self, id_cliente):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.data, oc.numero, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as ref, prod.unidade, prodoc.quantidade, "
                           f"prodoc.produzido, COALESCE(prodoc.id_pedido, '') as pedi, "
                           f"COALESCE(ped.num_req_cliente, '') as req, "
                           f"COALESCE(oc.obs, '') as obs "
                           f"FROM PRODUTOORDEMCOMPRA as prodoc "
                           f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                           f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                           f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                           f"LEFT JOIN pedidointerno as ped ON prodoc.id_pedido = ped.id "
                           f"where prodoc.quantidade > prodoc.produzido "
                           f"and oc.status = 'A' "
                           f"and oc.entradasaida = 'S'"
                           f"and oc.cliente = {id_cliente};")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    emissao, num_ov, clie, cod, descr, ref, um, qtde, qtde_entr, nume_pi, num_req, obs = i

                    emi = f'{emissao.day}/{emissao.month}/{emissao.year}'

                    dados = (emi, num_ov, clie, cod, descr, ref, um, qtde, qtde_entr, nume_pi, num_req, obs)
                    tabela_nova.append(dados)
            if tabela_nova:
                lista_de_listas_ordenada = sorted(tabela_nova, key=lambda x: x[1])
                lanca_tabela(self.table_OV, lista_de_listas_ordenada)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def dados_ov_todos_com_cliente(self, id_cliente):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.data, oc.numero, cli.razao, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as ref, prod.unidade, prodoc.quantidade, "
                           f"prodoc.produzido, COALESCE(prodoc.id_pedido, '') as pedi, "
                           f"COALESCE(ped.num_req_cliente, '') as req, "
                           f"COALESCE(oc.obs, '') as obs "
                           f"FROM PRODUTOORDEMCOMPRA as prodoc "
                           f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                           f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                           f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                           f"LEFT JOIN pedidointerno as ped ON prodoc.id_pedido = ped.id "
                           f"where oc.entradasaida = 'S'"
                           f"and oc.cliente = {id_cliente};")
            dados_interno = cursor.fetchall()
            if dados_interno:
                for i in dados_interno:
                    emissao, num_ov, clie, cod, descr, ref, um, qtde, qtde_entr, nume_pi, num_req, obs = i

                    emi = f'{emissao.day}/{emissao.month}/{emissao.year}'

                    dados = (emi, num_ov, clie, cod, descr, ref, um, qtde, qtde_entr, nume_pi, num_req, obs)
                    tabela_nova.append(dados)
            if tabela_nova:
                lista_de_listas_ordenada = sorted(tabela_nova, key=lambda x: x[1])
                lanca_tabela(self.table_OV, lista_de_listas_ordenada)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaVendasStatus()
    tela.show()
    qt.exec_()

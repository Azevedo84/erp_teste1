import sys
from banco_dados.conexao import conecta
from forms.tela_compras_status import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import lanca_tabela, layout_cabec_tab
from comandos.telas import tamanho_aplicacao, icone
from comandos.cores import cor_cinza_claro
from comandos.conversores import valores_para_float
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtGui import QColor
from PyQt5.QtCore import QThread, pyqtSignal
import inspect
import os
import traceback


class SolicitacaoTotalBaixada(QThread):
    new_value = pyqtSignal(list)

    def __init__(self):
        super(SolicitacaoTotalBaixada, self).__init__()

    def run(self):
        cursor = conecta.cursor()
        cursor.execute(f"SELECT (extract(day from req.dataemissao)||'-'||extract(month from req.dataemissao)||'-'||"
                       f"extract(year from req.dataemissao)) AS DATA, "
                       f"COALESCE(prodreq.mestre, ''), prod.codigo, "
                       f"CASE prod.id when 28761 then prodreq.descricao else prod.descricao end as DESCRICAO, "
                       f"CASE prod.embalagem when 'SIM' then COALESCE(prodreq.referencia, '') "
                       f"else COALESCE(prod.obs, '') end as REFERENCIA, "
                       f"prod.unidade, prodreq.quantidade, COALESCE(prodreq.destino, ''), prodreq.status, "
                       f"COALESCE(req.nome_pc, '') "
                       f"FROM produtoordemsolicitacao as prodreq "
                       f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                       f"INNER JOIN ordemsolicitacao as req ON prodreq.mestre = req.idsolicitacao "
                       f"LEFT JOIN produtoordemrequisicao as preq ON prodreq.id = preq.id_prod_sol "
                       f"WHERE prodreq.status = 'B' ORDER BY prodreq.mestre;")
        extrair_req = cursor.fetchall()

        self.new_value.emit(extrair_req)


class SolicitacaoTotal(QThread):
    new_value = pyqtSignal(list)

    def __init__(self):
        super(SolicitacaoTotal, self).__init__()

    def run(self):
        cursor = conecta.cursor()
        cursor.execute(f"SELECT (extract(day from req.dataemissao)||'-'||extract(month from req.dataemissao)||'-'||"
                       f"extract(year from req.dataemissao)) AS DATA, "
                       f"COALESCE(prodreq.mestre, ''), prod.codigo, "
                       f"CASE prod.id when 28761 then prodreq.descricao else prod.descricao end as DESCRICAO, "
                       f"CASE prod.embalagem when 'SIM' then COALESCE(prodreq.referencia, '') "
                       f"else COALESCE(prod.obs, '') end as REFERENCIA, "
                       f"prod.unidade, prodreq.quantidade, COALESCE(prodreq.destino, ''), prodreq.status, "
                       f"COALESCE(req.nome_pc, '') "
                       f"FROM produtoordemsolicitacao as prodreq "
                       f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                       f"INNER JOIN ordemsolicitacao as req ON prodreq.mestre = req.idsolicitacao "
                       f"LEFT JOIN produtoordemrequisicao as preq ON prodreq.id = preq.id_prod_sol "
                       f"ORDER BY prodreq.mestre;")
        result = cursor.fetchall()

        self.new_value.emit(result)


class RequisicaoTotalBaixada(QThread):
    new_value = pyqtSignal(list)

    def __init__(self):
        super(RequisicaoTotalBaixada, self).__init__()

    def run(self):
        dados_man_req = []

        cursor = conecta.cursor()
        cursor.execute(f"SELECT (extract(day from req.data)||'-'||"
                       f"extract(month from req.data)||'-'||"
                       f"extract(year from req.data)) AS DATA, "
                       f"prodreq.numero, prodreq.produto, prodreq.quantidade, "
                       f"prodreq.destino, prodreq.id_prod_sol, prodreq.status "
                       f"FROM produtoordemrequisicao as prodreq "
                       f"INNER JOIN ordemrequisicao as req ON prodreq.mestre = req.id "
                       f"where prodreq.status = 'B' "
                       f"ORDER BY prodreq.numero;")
        select_req = cursor.fetchall()

        for dados_req in select_req:
            data, numero, produto, qtde, destino, id_prod_sol, status = dados_req

            cur = conecta.cursor()
            cur.execute(f"SELECT codigo, descricao, COALESCE(obs, ' ') as obs, unidade "
                        f"FROM produto where id = {produto};")
            detalhes_produtos = cur.fetchall()
            cod, descr, ref, um = detalhes_produtos[0]

            if id_prod_sol is None:
                num_sol = "X"
            else:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, mestre "
                               f"FROM produtoordemsolicitacao "
                               f"WHERE id = {id_prod_sol};")
                extrair_sol = cursor.fetchall()
                id_sol, num_so = extrair_sol[0]

                if not extrair_sol:
                    num_sol = "X"
                else:
                    num_sol = num_so

            dados = (data, numero, cod, descr, ref, um, qtde, destino, num_sol, status)
            dados_man_req.append(dados)

        self.new_value.emit(dados_man_req)


class RequisicaoTotal(QThread):
    new_value = pyqtSignal(list)

    def __init__(self):
        super(RequisicaoTotal, self).__init__()

    def run(self):
        dados_man_req = []

        cursor = conecta.cursor()
        cursor.execute(f"SELECT (extract(day from req.data)||'-'||"
                       f"extract(month from req.data)||'-'||"
                       f"extract(year from req.data)) AS DATA, "
                       f"prodreq.numero, prodreq.produto, prodreq.quantidade, "
                       f"prodreq.destino, prodreq.id_prod_sol, prodreq.status "
                       f"FROM produtoordemrequisicao as prodreq "
                       f"INNER JOIN ordemrequisicao as req ON prodreq.mestre = req.id "
                       f"ORDER BY prodreq.numero;")
        select_req = cursor.fetchall()

        for dados_req in select_req:
            data, numero, produto, qtde, destino, id_prod_sol, status = dados_req

            cur = conecta.cursor()
            cur.execute(f"SELECT codigo, descricao, COALESCE(obs, ' ') as obs, unidade "
                        f"FROM produto where id = {produto};")
            detalhes_produtos = cur.fetchall()
            cod, descr, ref, um = detalhes_produtos[0]

            if id_prod_sol is None:
                num_sol = "X"
            else:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, mestre "
                               f"FROM produtoordemsolicitacao "
                               f"WHERE id = {id_prod_sol};")
                extrair_sol = cursor.fetchall()
                id_sol, num_so = extrair_sol[0]

                if not extrair_sol:
                    num_sol = "X"
                else:
                    num_sol = num_so

            dados = (data, numero, cod, descr, ref, um, qtde, destino, num_sol, status)
            dados_man_req.append(dados)

        self.new_value.emit(dados_man_req)


class OrdemCompraTotalBaixada(QThread):
    new_value = pyqtSignal(list)

    def __init__(self):
        super(OrdemCompraTotalBaixada, self).__init__()

    def run(self):
        tabela = []

        cursor = conecta.cursor()
        cursor.execute(
            f"SELECT COALESCE(prodreq.numero, ''), oc.data, oc.numero, forn.razao, prodoc.codigo, "
            f"prod.descricao, COALESCE(prod.obs, ''), "
            f"prod.unidade, prodoc.quantidade, prodoc.produzido, prodoc.dataentrega, "
            f"COALESCE(prodsol.mestre, ''), oc.STATUS "
            f"FROM ordemcompra as oc "
            f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
            f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
            f"LEFT JOIN produtoordemSOLICITACAO as prodsol ON prodreq.id_prod_sol = prodsol.id "
            f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
            f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
            f"where oc.entradasaida = 'E' AND oc.STATUS = 'B' "
            f"ORDER BY oc.data;")
        dados_oc = cursor.fetchall()

        if dados_oc:
            for i in dados_oc:
                id_req, data, oc, forncec, cod, descr, ref, um, qtde, prod, entr_dt, id_sol, status = i

                emissao = data.strftime("%d/%m/%Y")
                if entr_dt:
                    entrega = entr_dt.strftime("%d/%m/%Y")
                else:
                    entrega = ""

                dados = (emissao, oc, cod, descr, ref, um, qtde, entrega, forncec, id_req, id_sol, status, prod)
                tabela.append(dados)

        self.new_value.emit(tabela)


class OrdemCompraTotal(QThread):
    new_value = pyqtSignal(list)

    def __init__(self):
        super(OrdemCompraTotal, self).__init__()

    def run(self):
        tabela = []

        cursor = conecta.cursor()
        cursor.execute(
            f"SELECT COALESCE(prodreq.numero, ''), oc.data, oc.numero, forn.razao, prodoc.codigo, "
            f"prod.descricao, COALESCE(prod.obs, ''), "
            f"prod.unidade, prodoc.quantidade, prodoc.produzido, prodoc.dataentrega, "
            f"COALESCE(prodsol.mestre, ''), oc.STATUS "
            f"FROM ordemcompra as oc "
            f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
            f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
            f"LEFT JOIN produtoordemSOLICITACAO as prodsol ON prodreq.id_prod_sol = prodsol.id "
            f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
            f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
            f"where oc.entradasaida = 'E' "
            f"ORDER BY oc.data;")
        dados_oc = cursor.fetchall()

        if dados_oc:
            for i in dados_oc:
                id_req, data, oc, forncec, cod, descr, ref, um, qtde, produ, entr_dt, id_sol, status = i

                emissao = data.strftime("%d/%m/%Y")

                if entr_dt:
                    entrega = entr_dt.strftime("%d/%m/%Y")
                else:
                    entrega = ""

                if status == "A":
                    qtde_oc_float = valores_para_float(qtde)
                    qtde_produ_float = valores_para_float(produ)

                    falta_ent = qtde_oc_float - qtde_produ_float

                    if falta_ent > 0:
                        status_f = "A"

                    else:
                        status_f = "B"

                else:
                    status_f = "B"

                dados = (emissao, oc, cod, descr, ref, um, qtde, entrega, forncec, id_req, id_sol, status_f, produ)
                tabela.append(dados)

        self.new_value.emit(tabela)


class TelaComprasStatus(QMainWindow, Ui_Consulta_Sol):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        self.solicita_total_b = SolicitacaoTotalBaixada()
        self.solicita_total = SolicitacaoTotal()

        self.requisita_total_b = RequisicaoTotalBaixada()
        self.requisita_total = RequisicaoTotal()

        self.ordem_total_b = OrdemCompraTotalBaixada()
        self.ordem_total = OrdemCompraTotal()

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_compra_sol.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Solicitacao)
        layout_cabec_tab(self.table_Requisicao)
        layout_cabec_tab(self.table_OC)

        self.btn_Status_Sol.clicked.connect(self.manipula_sol_por_status)
        self.btn_Num_Sol.clicked.connect(self.manipula_sol_por_numero)
        self.line_Num_Sol.editingFinished.connect(self.manipula_sol_por_numero)
        self.line_Codigo_Sol.editingFinished.connect(self.manipula_sol_por_produto)
        self.btn_Prod_Sol.clicked.connect(self.manipula_sol_por_produto)

        self.btn_Status_Req.clicked.connect(self.manipula_req_por_status)
        self.btn_Num_Req.clicked.connect(self.manipula_req_por_numero)
        self.line_Num_Req.editingFinished.connect(self.manipula_req_por_numero)
        self.line_Codigo_Req.editingFinished.connect(self.manipula_req_por_produto)
        self.btn_Prod_Req.clicked.connect(self.manipula_req_por_produto)

        self.btn_Status_OC.clicked.connect(self.manipula_oc_por_status)
        self.line_Num_Fornec.editingFinished.connect(self.manipula_oc_por_status)
        self.btn_Num_OC.clicked.connect(self.manipula_oc_por_numero)
        self.line_Num_OC.editingFinished.connect(self.manipula_oc_por_numero)
        self.line_Codigo_OC.editingFinished.connect(self.manipula_oc_por_produto)
        self.btn_Prod_OC.clicked.connect(self.manipula_oc_por_produto)

        self.processando = False

        self.widget_Progress.setHidden(True)

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

    def manipula_sol_por_status(self):
        try:
            self.widget_Progress.setHidden(False)

            self.table_Solicitacao.setRowCount(0)

            abertas = self.check_Aberto_Sol.isChecked()
            fechadas = self.check_Baixado_Sol.isChecked()

            if abertas and fechadas:
                self.run_sol_total()
            elif abertas:
                self.sol_total_aberto()
            elif fechadas:
                self.run_sol_total_baixada()
            else:
                self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def sol_total_aberto(self):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT (extract(day from req.dataemissao)||'-'||extract(month from req.dataemissao)||'-'||"
                           f"extract(year from req.dataemissao)) AS DATA, "
                           f"COALESCE(prodreq.mestre, ''), prod.codigo, "
                           f"CASE prod.id when 28761 then prodreq.descricao else prod.descricao end as DESCRICAO, "
                           f"CASE prod.embalagem when 'SIM' then COALESCE(prodreq.referencia, '') "
                           f"else COALESCE(prod.obs, '') end as REFERENCIA, "
                           f"prod.unidade, prodreq.quantidade, COALESCE(prodreq.destino, ''), prodreq.status, "
                           f"COALESCE(req.nome_pc, '') "
                           f"FROM produtoordemsolicitacao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"INNER JOIN ordemsolicitacao as req ON prodreq.mestre = req.idsolicitacao "
                           f"LEFT JOIN produtoordemrequisicao as preq ON prodreq.id = preq.id_prod_sol "
                           f"WHERE prodreq.status = 'A' AND preq.id_prod_sol IS NULL ORDER BY prodreq.mestre;")
            extrair_req = cursor.fetchall()

            if extrair_req:
                lanca_tabela(self.table_Solicitacao, extrair_req)

                self.pintar_tabela_solicitacao(self.table_Solicitacao, extrair_req)

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def run_sol_total_baixada(self):
        try:
            self.solicita_total_b.new_value.connect(self.sol_total_baixada)
            self.solicita_total_b.start()
        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def sol_total_baixada(self, extrair_req):
        try:
            if extrair_req:
                lanca_tabela(self.table_Solicitacao, extrair_req)

                self.pintar_tabela_solicitacao(self.table_Solicitacao, extrair_req)

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def run_sol_total(self):
        try:
            self.solicita_total.new_value.connect(self.sol_total)
            self.solicita_total.start()
        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def sol_total(self, extrair_req):
        try:
            if extrair_req:
                lanca_tabela(self.table_Solicitacao, extrair_req)
                self.pintar_tabela_solicitacao(self.table_Solicitacao, extrair_req)

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_sol_por_numero(self):
        try:
            self.widget_Progress.setHidden(False)

            self.table_Solicitacao.setRowCount(0)

            num_sol = self.line_Num_Sol.text()

            if num_sol:
                self.sol_total_por_numero()
            else:
                self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def sol_total_por_numero(self):
        try:
            num_sol = self.line_Num_Sol.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT (extract(day from req.dataemissao)||'-'||extract(month from req.dataemissao)||'-'||"
                           f"extract(year from req.dataemissao)) AS DATA, "
                           f"COALESCE(prodreq.mestre, ''), prod.codigo, "
                           f"CASE prod.id when 28761 then prodreq.descricao else prod.descricao end as DESCRICAO, "
                           f"CASE prod.embalagem when 'SIM' then COALESCE(prodreq.referencia, '') "
                           f"else COALESCE(prod.obs, '') end as REFERENCIA, "
                           f"prod.unidade, prodreq.quantidade, COALESCE(prodreq.destino, ''), prodreq.status, "
                           f"COALESCE(req.nome_pc, '') "
                           f"FROM produtoordemsolicitacao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"INNER JOIN ordemsolicitacao as req ON prodreq.mestre = req.idsolicitacao "
                           f"LEFT JOIN produtoordemrequisicao as preq ON prodreq.id = preq.id_prod_sol "
                           f"WHERE prodreq.mestre = {num_sol}"
                           f"ORDER BY prodreq.mestre;")
            extrair_req = cursor.fetchall()

            if extrair_req:
                lanca_tabela(self.table_Solicitacao, extrair_req)

                self.pintar_tabela_solicitacao(self.table_Solicitacao, extrair_req)

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_sol_por_produto(self):
        if not self.processando:
            try:
                self.processando = True

                self.widget_Progress.setHidden(False)

                self.table_Solicitacao.setRowCount(0)

                codigo_produto = self.line_Codigo_Sol.text()

                if codigo_produto:
                    codigo_produto = self.line_Codigo_Sol.text()
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade, localizacao, quantidade "
                                   f"FROM produto where codigo = {codigo_produto};")
                    detalhes_produto = cursor.fetchall()
                    if not detalhes_produto:
                        self.mensagem_alerta('Este código de produto não existe!')
                        self.line_Codigo_Sol.clear()
                    else:
                        codigo_produto = self.line_Codigo_Sol.text()
                        cur = conecta.cursor()
                        cur.execute(f"SELECT descricao, COALESCE(descricaocomplementar, '') as compl, "
                                    f"COALESCE(obs, '') as obs, unidade, COALESCE(ncm, '') as local, "
                                    f"quantidade, embalagem FROM produto where codigo = {codigo_produto};")
                        detalhes_produto = cur.fetchall()
                        descr, compl, ref, um, ncm, saldo, embalagem = detalhes_produto[0]

                        self.line_Descricao_Sol.setText(descr)
                        self.line_Referencia_Sol.setText(ref)
                        self.line_UM_Sol.setText(um)

                        self.sol_total_por_produto()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo)
                grava_erro_banco(nome_funcao, e, self.nome_arquivo)

            finally:
                self.processando = False

    def sol_total_por_produto(self):
        try:
            cod_prod = self.line_Codigo_Sol.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT (extract(day from req.dataemissao)||'-'||extract(month from req.dataemissao)||'-'||"
                           f"extract(year from req.dataemissao)) AS DATA, "
                           f"COALESCE(prodreq.mestre, ''), prod.codigo, "
                           f"CASE prod.id when 28761 then prodreq.descricao else prod.descricao end as DESCRICAO, "
                           f"CASE prod.embalagem when 'SIM' then COALESCE(prodreq.referencia, '') "
                           f"else COALESCE(prod.obs, '') end as REFERENCIA, "
                           f"prod.unidade, prodreq.quantidade, COALESCE(prodreq.destino, ''), prodreq.status, "
                           f"COALESCE(req.nome_pc, '') "
                           f"FROM produtoordemsolicitacao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"INNER JOIN ordemsolicitacao as req ON prodreq.mestre = req.idsolicitacao "
                           f"LEFT JOIN produtoordemrequisicao as preq ON prodreq.id = preq.id_prod_sol "
                           f"WHERE prod.codigo = {cod_prod}"
                           f"ORDER BY prodreq.mestre;")
            extrair_req = cursor.fetchall()

            if extrair_req:
                lanca_tabela(self.table_Solicitacao, extrair_req)

                self.pintar_tabela_solicitacao(self.table_Solicitacao, extrair_req)

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_req_por_status(self):
        try:
            self.widget_Progress.setHidden(False)

            self.table_Requisicao.setRowCount(0)

            abertas = self.check_Aberto_Req.isChecked()
            fechadas = self.check_Baixado_Req.isChecked()

            if abertas and fechadas:
                self.run_req_total()
            elif abertas:
                self.req_total_aberto()
            elif fechadas:
                self.run_req_total_baixada()
            else:
                self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def req_total_aberto(self):
        try:
            dados_man_req = []
            cursor = conecta.cursor()
            cursor.execute(f"SELECT (extract(day from req.data)||'-'||"
                           f"extract(month from req.data)||'-'||"
                           f"extract(year from req.data)) AS DATA, "
                           f"prodreq.numero, prodreq.produto, prodreq.quantidade, "
                           f"prodreq.destino, prodreq.id_prod_sol, prodreq.status "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN ordemrequisicao as req ON prodreq.mestre = req.id "
                           f"where prodreq.status = 'A';")
            select_req = cursor.fetchall()

            for dados_req in select_req:
                data, numero, produto, qtde, destino, id_prod_sol, status = dados_req

                cur = conecta.cursor()
                cur.execute(f"SELECT codigo, descricao, COALESCE(obs, ' ') as obs, unidade "
                            f"FROM produto where id = {produto};")
                detalhes_produtos = cur.fetchall()
                cod, descr, ref, um = detalhes_produtos[0]

                if id_prod_sol is None:
                    num_sol = "X"
                else:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, mestre "
                                   f"FROM produtoordemsolicitacao "
                                   f"WHERE id = {id_prod_sol};")
                    extrair_sol = cursor.fetchall()
                    id_sol, num_so = extrair_sol[0]

                    if not extrair_sol:
                        num_sol = "X"
                    else:
                        num_sol = num_so

                dados = (data, numero, cod, descr, ref, um, qtde, destino, num_sol, status)
                dados_man_req.append(dados)

            if dados_man_req:
                lanca_tabela(self.table_Requisicao, dados_man_req)

                self.pintar_tabela_requisicao(self.table_Requisicao, dados_man_req)

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def run_req_total_baixada(self):
        try:
            self.requisita_total_b.new_value.connect(self.req_total_baixada)
            self.requisita_total_b.start()
        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def req_total_baixada(self, dados_man_req):
        try:
            if dados_man_req:
                lanca_tabela(self.table_Requisicao, dados_man_req)

                self.pintar_tabela_requisicao(self.table_Requisicao, dados_man_req)

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def run_req_total(self):
        try:
            self.requisita_total.new_value.connect(self.req_total)
            self.requisita_total.start()
        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def req_total(self, dados_man_req):
        try:
            if dados_man_req:
                lanca_tabela(self.table_Requisicao, dados_man_req)

                self.pintar_tabela_requisicao(self.table_Requisicao, dados_man_req)

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_req_por_numero(self):
        try:
            self.widget_Progress.setHidden(False)

            self.table_Requisicao.setRowCount(0)

            num_req = self.line_Num_Req.text()

            if num_req:
                self.req_total_por_numero()
            else:
                self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_req_por_produto(self):
        if not self.processando:
            try:
                self.processando = True

                self.widget_Progress.setHidden(False)

                self.table_Requisicao.setRowCount(0)

                codigo_produto = self.line_Codigo_Req.text()

                if codigo_produto:
                    codigo_produto = self.line_Codigo_Req.text()
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade, localizacao, quantidade "
                                   f"FROM produto where codigo = {codigo_produto};")
                    detalhes_produto = cursor.fetchall()
                    if not detalhes_produto:
                        self.mensagem_alerta('Este código de produto não existe!')
                        self.line_Codigo_Req.clear()
                        self.widget_Progress.setHidden(True)
                    else:
                        codigo_produto = self.line_Codigo_Req.text()
                        cur = conecta.cursor()
                        cur.execute(f"SELECT descricao, COALESCE(descricaocomplementar, '') as compl, "
                                    f"COALESCE(obs, '') as obs, unidade, COALESCE(ncm, '') as local, "
                                    f"quantidade, embalagem FROM produto where codigo = {codigo_produto};")
                        detalhes_produto = cur.fetchall()
                        descr, compl, ref, um, ncm, saldo, embalagem = detalhes_produto[0]

                        self.line_Descricao_Req.setText(descr)
                        self.line_Referencia_Req.setText(ref)
                        self.line_UM_Req.setText(um)

                        self.req_total_por_produto()
                else:
                    self.widget_Progress.setHidden(True)

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo)
                grava_erro_banco(nome_funcao, e, self.nome_arquivo)

            finally:
                self.processando = False

    def req_total_por_numero(self):
        try:
            num_req = self.line_Num_Req.text()

            dados_man_req = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT (extract(day from req.data)||'-'||"
                           f"extract(month from req.data)||'-'||"
                           f"extract(year from req.data)) AS DATA, "
                           f"prodreq.numero, prodreq.produto, prodreq.quantidade, "
                           f"prodreq.destino, prodreq.id_prod_sol, prodreq.status "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN ordemrequisicao as req ON prodreq.mestre = req.id "
                           f"WHERE prodreq.numero = {num_req}"
                           f"ORDER BY prodreq.numero;")
            select_req = cursor.fetchall()

            for dados_req in select_req:
                data, numero, produto, qtde, destino, id_prod_sol, status = dados_req

                cur = conecta.cursor()
                cur.execute(f"SELECT codigo, descricao, COALESCE(obs, ' ') as obs, unidade "
                            f"FROM produto where id = {produto};")
                detalhes_produtos = cur.fetchall()
                cod, descr, ref, um = detalhes_produtos[0]

                if id_prod_sol is None:
                    num_sol = "X"
                else:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, mestre "
                                   f"FROM produtoordemsolicitacao "
                                   f"WHERE id = {id_prod_sol};")
                    extrair_sol = cursor.fetchall()
                    id_sol, num_so = extrair_sol[0]

                    if not extrair_sol:
                        num_sol = "X"
                    else:
                        num_sol = num_so

                dados = (data, numero, cod, descr, ref, um, qtde, destino, num_sol, status)
                dados_man_req.append(dados)

            if dados_man_req:
                lanca_tabela(self.table_Requisicao, dados_man_req)

                self.pintar_tabela_requisicao(self.table_Requisicao, dados_man_req)

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def req_total_por_produto(self):
        try:
            cod_prod = self.line_Codigo_Req.text()

            dados_man_req = []

            cur = conecta.cursor()
            cur.execute(f"SELECT id, codigo FROM produto where codigo = {cod_prod};")
            detalhes_produtos0 = cur.fetchall()
            id_prod0, cod0 = detalhes_produtos0[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT (extract(day from req.data)||'-'||"
                           f"extract(month from req.data)||'-'||"
                           f"extract(year from req.data)) AS DATA, "
                           f"prodreq.numero, prodreq.produto, prodreq.quantidade, "
                           f"prodreq.destino, prodreq.id_prod_sol, prodreq.status "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN ordemrequisicao as req ON prodreq.mestre = req.id "
                           f"WHERE prodreq.produto = {id_prod0}"
                           f"ORDER BY prodreq.numero;")
            select_req = cursor.fetchall()

            for dados_req in select_req:
                data, numero, produto, qtde, destino, id_prod_sol, status = dados_req

                cur = conecta.cursor()
                cur.execute(f"SELECT codigo, descricao, COALESCE(obs, ' ') as obs, unidade "
                            f"FROM produto where id = {produto};")
                detalhes_produtos = cur.fetchall()
                cod, descr, ref, um = detalhes_produtos[0]

                if id_prod_sol is None:
                    num_sol = "X"
                else:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, mestre "
                                   f"FROM produtoordemsolicitacao "
                                   f"WHERE id = {id_prod_sol};")
                    extrair_sol = cursor.fetchall()
                    id_sol, num_so = extrair_sol[0]

                    if not extrair_sol:
                        num_sol = "X"
                    else:
                        num_sol = num_so

                dados = (data, numero, cod, descr, ref, um, qtde, destino, num_sol, status)
                dados_man_req.append(dados)

            if dados_man_req:
                lanca_tabela(self.table_Requisicao, dados_man_req)

                self.pintar_tabela_requisicao(self.table_Requisicao, dados_man_req)

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_oc_por_status(self):
        try:
            self.widget_Progress.setHidden(False)

            self.table_OC.setRowCount(0)

            abertas = self.check_Aberto_OC.isChecked()
            fechadas = self.check_Baixado_OC.isChecked()
            num_fonec = self.line_Num_Fornec.text()

            if num_fonec:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, razao FROM fornecedores where registro = {num_fonec};")
                dados_fornecedor = cursor.fetchall()

                if not dados_fornecedor:
                    self.mensagem_alerta('Este código de Fornecedor não existe!')
                    self.line_Num_Fornec.clear()
                    self.widget_Progress.setHidden(True)
                else:
                    if abertas and fechadas:
                        self.oc_total_forn()
                    elif abertas:
                        self.oc_total_aberto_fornc()
                    elif fechadas:
                        self.oc_total_baixada_fornc()
                    else:
                        self.widget_Progress.setHidden(True)
            else:
                if abertas and fechadas:
                    self.run_oc_total()
                elif abertas:
                    self.oc_total_aberto()
                elif fechadas:
                    self.run_oc_total_baixada()
                else:
                    self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def oc_total_aberto(self):
        try:
            tabela = []

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT COALESCE(prodreq.numero, ''), oc.data, oc.numero, forn.razao, prodoc.codigo, "
                f"prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, prodoc.quantidade, prodoc.produzido, prodoc.dataentrega, "
                f"COALESCE(prodsol.mestre, ''), oc.STATUS "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                f"LEFT JOIN produtoordemSOLICITACAO as prodsol ON prodreq.id_prod_sol = prodsol.id "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"where oc.entradasaida = 'E' AND oc.STATUS = 'A' AND prodoc.produzido < prodoc.quantidade "
                f"ORDER BY oc.numero;")
            dados_oc = cursor.fetchall()

            if dados_oc:
                for i in dados_oc:
                    id_req, data, oc, forncec, cod, descr, ref, um, qtde, prod, entr_dt, id_sol, status = i

                    emissao = data.strftime("%d/%m/%Y")
                    entrega = entr_dt.strftime("%d/%m/%Y")

                    dados = (emissao, oc, cod, descr, ref, um, qtde, entrega, forncec, id_req, id_sol, status, prod)
                    tabela.append(dados)

            if tabela:
                lanca_tabela(self.table_OC, tabela)

                self.pintar_tabela_oc(self.table_OC, tabela)

                self.table_OC.setFocus()
                self.limpa_dados_oc()

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def run_oc_total_baixada(self):
        try:
            self.ordem_total_b.new_value.connect(self.oc_total_baixada)
            self.ordem_total_b.start()
        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def oc_total_baixada(self, tabela):
        try:
            if tabela:
                lanca_tabela(self.table_OC, tabela)

                self.pintar_tabela_oc(self.table_OC, tabela)

                self.table_OC.setFocus()
                self.limpa_dados_oc()

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def run_oc_total(self):
        try:
            self.ordem_total.new_value.connect(self.oc_total)
            self.ordem_total.start()
        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def oc_total(self, tabela):
        try:
            if tabela:
                lanca_tabela(self.table_OC, tabela)
                self.pintar_tabela_oc(self.table_OC, tabela)
                self.table_OC.setFocus()
                self.limpa_dados_oc()

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def oc_total_aberto_fornc(self):
        try:
            num_fonec = self.line_Num_Fornec.text()

            tabela = []

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT COALESCE(prodreq.numero, ''), oc.data, oc.numero, forn.razao, prodoc.codigo, "
                f"prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, prodoc.quantidade, prodoc.produzido, prodoc.dataentrega, "
                f"COALESCE(prodsol.mestre, ''), oc.STATUS "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                f"LEFT JOIN produtoordemSOLICITACAO as prodsol ON prodreq.id_prod_sol = prodsol.id "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"where forn.registro = {num_fonec} "
                f"and oc.entradasaida = 'E' "
                f"AND oc.STATUS = 'A' "
                f"AND prodoc.produzido < prodoc.quantidade "
                f"ORDER BY oc.numero;")
            dados_oc = cursor.fetchall()

            if dados_oc:
                for i in dados_oc:
                    id_req, data, oc, forncec, cod, descr, ref, um, qtde, prod, entr_dt, id_sol, status = i

                    emissao = data.strftime("%d/%m/%Y")
                    entrega = entr_dt.strftime("%d/%m/%Y")

                    dados = (emissao, oc, cod, descr, ref, um, qtde, entrega, forncec, id_req, id_sol, status, prod)
                    tabela.append(dados)

            if tabela:
                lanca_tabela(self.table_OC, tabela)

                self.pintar_tabela_oc(self.table_OC, tabela)
                self.table_OC.setFocus()
                self.limpa_dados_oc()

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def oc_total_baixada_fornc(self):
        try:
            num_fonec = self.line_Num_Fornec.text()

            tabela = []

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT COALESCE(prodreq.numero, ''), oc.data, oc.numero, forn.razao, prodoc.codigo, "
                f"prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, prodoc.quantidade, prodoc.produzido, prodoc.dataentrega, "
                f"COALESCE(prodsol.mestre, ''), oc.STATUS "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                f"LEFT JOIN produtoordemSOLICITACAO as prodsol ON prodreq.id_prod_sol = prodsol.id "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"where forn.registro = {num_fonec} "
                f"and oc.entradasaida = 'E' "
                f"AND oc.STATUS = 'B' "
                f"ORDER BY oc.data;")
            dados_oc = cursor.fetchall()

            if dados_oc:
                for i in dados_oc:
                    id_req, data, oc, forncec, cod, descr, ref, um, qtde, prod, entr_dt, id_sol, status = i

                    emissao = data.strftime("%d/%m/%Y")
                    if entr_dt:
                        entrega = entr_dt.strftime("%d/%m/%Y")
                    else:
                        entrega = ""

                    dados = (emissao, oc, cod, descr, ref, um, qtde, entrega, forncec, id_req, id_sol, status, prod)
                    tabela.append(dados)

            if tabela:
                lanca_tabela(self.table_OC, tabela)

                self.pintar_tabela_oc(self.table_OC, tabela)
                self.table_OC.setFocus()
                self.limpa_dados_oc()

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def oc_total_forn(self):
        try:
            num_fonec = self.line_Num_Fornec.text()

            tabela = []

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT COALESCE(prodreq.numero, ''), oc.data, oc.numero, forn.razao, prodoc.codigo, "
                f"prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, prodoc.quantidade, prodoc.produzido, prodoc.dataentrega, "
                f"COALESCE(prodsol.mestre, ''), oc.STATUS "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                f"LEFT JOIN produtoordemSOLICITACAO as prodsol ON prodreq.id_prod_sol = prodsol.id "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"where forn.registro = {num_fonec} "
                f"and oc.entradasaida = 'E' "
                f"ORDER BY oc.data;")
            dados_oc = cursor.fetchall()

            if dados_oc:
                for i in dados_oc:
                    id_req, data, oc, forncec, cod, descr, ref, um, qtde, produ, entr_dt, id_sol, status = i

                    emissao = data.strftime("%d/%m/%Y")

                    if entr_dt:
                        entrega = entr_dt.strftime("%d/%m/%Y")
                    else:
                        entrega = ""

                    if status == "A":
                        qtde_oc_float = valores_para_float(qtde)
                        qtde_produ_float = valores_para_float(produ)

                        falta_ent = qtde_oc_float - qtde_produ_float

                        if falta_ent > 0:
                            status_f = "A"

                        else:
                            status_f = "B"

                    else:
                        status_f = "B"

                    dados = (emissao, oc, cod, descr, ref, um, qtde, entrega, forncec, id_req, id_sol, status_f, produ)
                    tabela.append(dados)

            if tabela:
                lanca_tabela(self.table_OC, tabela)

                self.pintar_tabela_oc(self.table_OC, tabela)
                self.table_OC.setFocus()
                self.limpa_dados_oc()

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_oc_por_numero(self):
        if not self.processando:
            try:
                self.processando = True

                self.widget_Progress.setHidden(False)

                self.table_OC.setRowCount(0)

                num_oc = self.line_Num_OC.text()

                if num_oc:
                    self.oc_total_por_numero()
                else:
                    self.widget_Progress.setHidden(True)

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def oc_total_por_numero(self):
        try:
            num_oc = self.line_Num_OC.text()

            tabela = []

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT COALESCE(prodreq.numero, ''), oc.data, oc.numero, forn.razao, prodoc.codigo, "
                f"prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, prodoc.quantidade, prodoc.produzido, prodoc.dataentrega, "
                f"COALESCE(prodsol.mestre, ''), oc.STATUS "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                f"LEFT JOIN produtoordemSOLICITACAO as prodsol ON prodreq.id_prod_sol = prodsol.id "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"where oc.numero = {num_oc} "
                f"and oc.entradasaida = 'E' "
                f"ORDER BY oc.data;")
            dados_oc = cursor.fetchall()

            if dados_oc:
                for i in dados_oc:
                    id_req, data, oc, forncec, cod, descr, ref, um, qtde, produ, entr_dt, id_sol, status = i

                    emissao = data.strftime("%d/%m/%Y")

                    if entr_dt:
                        entrega = entr_dt.strftime("%d/%m/%Y")
                    else:
                        entrega = ""

                    if status == "A":
                        qtde_oc_float = valores_para_float(qtde)
                        qtde_produ_float = valores_para_float(produ)

                        falta_ent = qtde_oc_float - qtde_produ_float

                        if falta_ent > 0:
                            status_f = "A"

                        else:
                            status_f = "B"

                    else:
                        status_f = "B"

                    dados = (emissao, oc, cod, descr, ref, um, qtde, entrega, forncec, id_req, id_sol, status_f, produ)
                    tabela.append(dados)

            if tabela:
                lanca_tabela(self.table_OC, tabela)

                self.pintar_tabela_oc(self.table_OC, tabela)

                self.table_OC.setFocus()
                self.limpa_dados_oc()

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_oc_por_produto(self):
        if not self.processando:
            try:
                self.processando = True

                self.widget_Progress.setHidden(False)

                self.table_OC.setRowCount(0)

                codigo_produto = self.line_Codigo_OC.text()

                if codigo_produto:
                    codigo_produto = self.line_Codigo_OC.text()
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade, localizacao, quantidade "
                                   f"FROM produto where codigo = {codigo_produto};")
                    detalhes_produto = cursor.fetchall()
                    if not detalhes_produto:
                        self.mensagem_alerta('Este código de produto não existe!')
                        self.line_Codigo_OC.clear()
                        self.widget_Progress.setHidden(True)
                    else:
                        codigo_produto = self.line_Codigo_OC.text()
                        cur = conecta.cursor()
                        cur.execute(f"SELECT descricao, COALESCE(descricaocomplementar, '') as compl, "
                                    f"COALESCE(obs, '') as obs, unidade, COALESCE(ncm, '') as local, "
                                    f"quantidade, embalagem FROM produto where codigo = {codigo_produto};")
                        detalhes_produto = cur.fetchall()
                        descr, compl, ref, um, ncm, saldo, embalagem = detalhes_produto[0]

                        self.line_Descricao_OC.setText(descr)
                        self.line_Referencia_OC.setText(ref)
                        self.line_UM_OC.setText(um)

                        self.oc_total_por_produto()
                else:
                    self.widget_Progress.setHidden(True)

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo)
                grava_erro_banco(nome_funcao, e, self.nome_arquivo)

            finally:
                self.processando = False

    def oc_total_por_produto(self):
        try:
            cod_prod = self.line_Codigo_OC.text()

            tabela = []

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT COALESCE(prodreq.numero, ''), oc.data, oc.numero, forn.razao, prodoc.codigo, "
                f"prod.descricao, COALESCE(prod.obs, ''), "
                f"prod.unidade, prodoc.quantidade, prodoc.produzido, prodoc.dataentrega, "
                f"COALESCE(prodsol.mestre, ''), oc.STATUS "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"LEFT JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                f"LEFT JOIN produtoordemSOLICITACAO as prodsol ON prodreq.id_prod_sol = prodsol.id "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"where prodoc.codigo = '{cod_prod}' "
                f"and oc.entradasaida = 'E' "
                f"ORDER BY oc.data;")
            dados_oc = cursor.fetchall()

            if dados_oc:
                for i in dados_oc:
                    id_req, data, oc, forncec, cod, descr, ref, um, qtde, produ, entr_dt, id_sol, status = i

                    emissao = data.strftime("%d/%m/%Y")

                    if entr_dt:
                        entrega = entr_dt.strftime("%d/%m/%Y")
                    else:
                        entrega = ""

                    if status == "A":
                        qtde_oc_float = valores_para_float(qtde)
                        qtde_produ_float = valores_para_float(produ)

                        falta_ent = qtde_oc_float - qtde_produ_float

                        if falta_ent > 0:
                            status_f = "A"

                        else:
                            status_f = "B"

                    else:
                        status_f = "B"

                    dados = (emissao, oc, cod, descr, ref, um, qtde, entrega, forncec, id_req, id_sol, status_f, produ)
                    tabela.append(dados)

            if tabela:
                lanca_tabela(self.table_OC, tabela)

                self.pintar_tabela_oc(self.table_OC, tabela)
                self.table_OC.setFocus()

                self.limpa_dados_oc()

            self.widget_Progress.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pintar_tabela_solicitacao(self, nome_tabela, extrai_tabela):
        try:
            for index, itens in enumerate(extrai_tabela):
                status = itens[8]

                if status == "B":
                    nome_tabela.item(index, 0).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 1).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 2).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 3).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 4).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 5).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 6).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 7).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 8).setBackground(QColor(cor_cinza_claro))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pintar_tabela_requisicao(self, nome_tabela, extrai_tabela):
        try:
            for index, itens in enumerate(extrai_tabela):
                status = itens[9]

                if status == "B":
                    nome_tabela.item(index, 0).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 1).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 2).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 3).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 4).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 5).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 6).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 7).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 8).setBackground(QColor(cor_cinza_claro))
                    nome_tabela.item(index, 9).setBackground(QColor(cor_cinza_claro))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pintar_tabela_oc(self, nome_tabela, extrai_tabela):
        try:
            for linha, itens_l in enumerate(extrai_tabela):
                status = itens_l[11]

                if status == "B":
                    for coluna, itens_c in enumerate(itens_l):
                        nome_tabela.item(linha, coluna).setBackground(QColor(cor_cinza_claro))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_oc(self):
        try:
            self.line_Num_OC.clear()
            self.line_Num_Fornec.clear()

            self.line_Codigo_OC.clear()
            self.line_UM_OC.clear()
            self.line_Descricao_OC.clear()
            self.line_Referencia_OC.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    telasolicitaconsulta = TelaComprasStatus()
    telasolicitaconsulta.show()
    qt.exec_()

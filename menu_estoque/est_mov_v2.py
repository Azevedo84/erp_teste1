import sys
from banco_dados.conexao import conecta
from forms.tela_est_mov import *
from banco_dados.controle_erros import grava_erro_banco
from arquivos.chamar_arquivos import definir_caminho_arquivo
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.telas import tamanho_aplicacao, icone
from comandos.excel import lanca_dados_mesclado, lanca_dados_coluna, edita_alinhamento, edita_bordas
from comandos.excel import adiciona_imagem, dataframe_pandas, escritor_dataframe, escritor_direto_dataframe
from comandos.excel import carregar_workbook
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from datetime import date, datetime, timedelta
from pathlib import Path
import inspect
import os
from threading import Thread
import traceback


class TelaEstMovimentacaoV2(QMainWindow, Ui_ConsultaOP):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_estoque.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_OP)

        data_hoje = date.today()

        self.date_Inicio.setDate(data_hoje)
        self.date_Final.setDate(data_hoje)

        self.btn_Conultar.clicked.connect(self.consulta)

        self.btn_Salvar.clicked.connect(self.final)

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

    def consulta(self):
        try:
            self.label_Excel.setText("")
            self.label_Msg.setText("")

            self.widget_Progress.setHidden(False)

            Thread(target=self.consulta_tred).start()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_tred(self):
        try:
            if self.radio_Almox.isChecked():
                dados_tabela = self.define_tabela_almox()
            else:
                dados_tabela = self.define_tabela_total()

            if dados_tabela:
                lanca_tabela(self.table_OP, dados_tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def select_mov_almox(self, data_inicio, data_fim, num_tipo, nome_tipo):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE((extract(day from m.data)||'/'||"
                           f"extract(month from m.data)||'/'||extract(year from m.data)), '') AS DATA, "
                           f"produto.codigo, produto.descricao, "
                           f"COALESCE(produto.obs, ''), produto.unidade, "
                           f"COALESCE(CASE WHEN m.tipo < 200 THEN m.quantidade END, 0) AS Qtde_Entrada, "
                           f"COALESCE(CASE WHEN m.tipo > 200 THEN m.quantidade END, 0) AS Qtde_Saida, "
                           f"(select case when sum(quantidade) is null then 0 else sum(quantidade) end "
                           f"from movimentacao where produto=m.produto "
                           f"and tipo<200 and localestoque=m.localestoque)-"
                           f"(select case when sum(quantidade) is null then 0 else sum(quantidade) end "
                           f"from movimentacao where produto=m.produto "
                           f"and tipo>200 and localestoque=m.localestoque)+"
                           f"(case when ((select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"(((m.tipo<200) and ((m2.data>m.data) or ((m2.data=m.data) "
                           f"and (m2.id>m.id)))) or(m.tipo>200 and m2.data>m.data)) "
                           f"and m2.tipo<200)*-1) is null then 0 else "
                           f"((select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"(((m.tipo<200) and ((m2.data>m.data) or((m2.data=m.data) "
                           f"and (m2.id>m.id)))) or(m.tipo>200 and m2.data>m.data)) "
                           f"and m2.tipo<200)*-1) end) + "
                           f"(case when (select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"((m2.data=m.data and (m2.id>m.id  or (m.tipo<200)) )or(m2.data>m.data)) "
                           f"and m2.tipo>200) is null then 0 else (select sum(m2.quantidade) "
                           f"from movimentacao m2 where m2.localestoque=m.localestoque "
                           f"and m2.produto=m.produto and ((m2.data=m.data "
                           f"and (m2.id>m.id or (m.tipo<200)) )or(m2.data>m.data)) and m2.tipo>200) end) "
                           f"as saldo, "
                           f"CASE WHEN m.tipo = 210 THEN ('OP '|| produtoos.numero) "
                           f"WHEN m.tipo = 110 THEN ('OP '|| ordemservico.numero) "
                           f"WHEN m.tipo = 130 THEN ('NF '|| entradaprod.nota) "
                           f"WHEN m.tipo = 140 THEN ('INVENTÁRIO') "
                           f"WHEN m.tipo = 240 THEN ('INVENTÁRIO') "
                           f"WHEN m.tipo = 230 THEN ('NF '|| saidaprod.numero) "
                           f"WHEN m.tipo = 250 THEN ('Devol. OS '|| produtoservico.numero) "
                           f"WHEN m.tipo = 112 THEN ('OS '|| produtoservico.numero) "
                           f"WHEN m.tipo = 220 THEN 'CI' "
                           f"END AS OS_NF_CI, "
                           f"COALESCE(natop.descricao, ''), localestoque.nome, "
                           f"CASE WHEN m.tipo = 210 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 110 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 130 THEN (fornecedores.razao) "
                           f"WHEN m.tipo = 140 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 230 THEN (clientes.razao) "
                           f"WHEN m.tipo = 250 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 112 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 220 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 240 THEN (funcionarios.funcionario) "
                           f"END AS teste, "
                           f"COALESCE(m.obs, '') "
                           f"FROM movimentacao m "
                           f"INNER JOIN produto ON (m.codigo = produto.codigo) "
                           f"INNER JOIN localestoque ON (m.localestoque = localestoque.id) "
                           f"LEFT JOIN funcionarios ON (m.funcionario = funcionarios.id) "
                           f"LEFT JOIN saidaprod ON (m.id = saidaprod.movimentacao) "
                           f"LEFT JOIN entradaprod ON (m.id = entradaprod.movimentacao) "
                           f"LEFT JOIN produtoservico ON (m.id = produtoservico.movimentacao) "
                           f"LEFT JOIN ordemservico ON (m.id = ordemservico.movimentacao) "
                           f"LEFT JOIN produtoos ON (m.id = produtoos.movimentacao) "
                           f"LEFT JOIN fornecedores ON (entradaprod.fornecedor = fornecedores.id) "
                           f"LEFT JOIN clientes ON (saidaprod.cliente = clientes.id) "
                           f"LEFT JOIN natop ON (( COALESCE( saidaprod.natureza, 0 ) + "
                           f"COALESCE( entradaprod.natureza, 0 ) ) = natop.ID) "
                           f"WHERE m.data >= '{data_inicio}' and m.data <= '{data_fim}' "
                           f"and localestoque.id IN (1, 2) and m.tipo = {num_tipo} "
                           f"order by m.data, {nome_tipo}, m.id;")
            results = cursor.fetchall()

            return results

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def select_mov_total(self, data_inicio, data_fim, num_tipo, nome_tipo):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE((extract(day from m.data)||'/'||"
                           f"extract(month from m.data)||'/'||extract(year from m.data)), '') AS DATA, "
                           f"produto.codigo, produto.descricao, "
                           f"COALESCE(produto.obs, ''), produto.unidade, "
                           f"COALESCE(CASE WHEN m.tipo < 200 THEN m.quantidade END, 0) AS Qtde_Entrada, "
                           f"COALESCE(CASE WHEN m.tipo > 200 THEN m.quantidade END, 0) AS Qtde_Saida, "
                           f"(select case when sum(quantidade) is null then 0 else sum(quantidade) end "
                           f"from movimentacao where produto=m.produto "
                           f"and tipo<200 and localestoque=m.localestoque)-"
                           f"(select case when sum(quantidade) is null then 0 else sum(quantidade) end "
                           f"from movimentacao where produto=m.produto "
                           f"and tipo>200 and localestoque=m.localestoque)+"
                           f"(case when ((select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"(((m.tipo<200) and ((m2.data>m.data) or ((m2.data=m.data) "
                           f"and (m2.id>m.id)))) or(m.tipo>200 and m2.data>m.data)) "
                           f"and m2.tipo<200)*-1) is null then 0 else "
                           f"((select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"(((m.tipo<200) and ((m2.data>m.data) or((m2.data=m.data) "
                           f"and (m2.id>m.id)))) or(m.tipo>200 and m2.data>m.data)) "
                           f"and m2.tipo<200)*-1) end) + "
                           f"(case when (select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"((m2.data=m.data and (m2.id>m.id  or (m.tipo<200)) )or(m2.data>m.data)) "
                           f"and m2.tipo>200) is null then 0 else (select sum(m2.quantidade) "
                           f"from movimentacao m2 where m2.localestoque=m.localestoque "
                           f"and m2.produto=m.produto and ((m2.data=m.data "
                           f"and (m2.id>m.id or (m.tipo<200)) )or(m2.data>m.data)) and m2.tipo>200) end) "
                           f"as saldo, "
                           f"CASE WHEN m.tipo = 210 THEN ('OP '|| produtoos.numero) "
                           f"WHEN m.tipo = 110 THEN ('OP '|| ordemservico.numero) "
                           f"WHEN m.tipo = 130 THEN ('NF '|| entradaprod.nota) "
                           f"WHEN m.tipo = 140 THEN ('INVENTÁRIO') "
                           f"WHEN m.tipo = 240 THEN ('INVENTÁRIO') "
                           f"WHEN m.tipo = 230 THEN ('NF '|| saidaprod.numero) "
                           f"WHEN m.tipo = 250 THEN ('Devol. OS '|| produtoservico.numero) "
                           f"WHEN m.tipo = 112 THEN ('OS '|| produtoservico.numero) "
                           f"WHEN m.tipo = 220 THEN 'CI' "
                           f"END AS OS_NF_CI, "
                           f"COALESCE(natop.descricao, ''), localestoque.nome, "
                           f"CASE WHEN m.tipo = 210 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 110 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 130 THEN (fornecedores.razao) "
                           f"WHEN m.tipo = 140 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 230 THEN (clientes.razao) "
                           f"WHEN m.tipo = 250 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 112 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 220 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 240 THEN (funcionarios.funcionario) "
                           f"END AS teste, "
                           f"COALESCE(m.obs, '') "
                           f"FROM movimentacao m "
                           f"INNER JOIN produto ON (m.codigo = produto.codigo) "
                           f"INNER JOIN localestoque ON (m.localestoque = localestoque.id) "
                           f"LEFT JOIN funcionarios ON (m.funcionario = funcionarios.id) "
                           f"LEFT JOIN saidaprod ON (m.id = saidaprod.movimentacao) "
                           f"LEFT JOIN entradaprod ON (m.id = entradaprod.movimentacao) "
                           f"LEFT JOIN produtoservico ON (m.id = produtoservico.movimentacao) "
                           f"LEFT JOIN ordemservico ON (m.id = ordemservico.movimentacao) "
                           f"LEFT JOIN produtoos ON (m.id = produtoos.movimentacao) "
                           f"LEFT JOIN fornecedores ON (entradaprod.fornecedor = fornecedores.id) "
                           f"LEFT JOIN clientes ON (saidaprod.cliente = clientes.id) "
                           f"LEFT JOIN natop ON (( COALESCE( saidaprod.natureza, 0 ) + "
                           f"COALESCE( entradaprod.natureza, 0 ) ) = natop.ID) "
                           f"WHERE m.data >= '{data_inicio}' and m.data <= '{data_fim}' "
                           f"and m.tipo = {num_tipo} "
                           f"order by m.data, {nome_tipo}, m.id;")
            results = cursor.fetchall()

            return results

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def define_tabela_almox(self):
        try:
            results = []

            data_inicio = self.date_Inicio.text()
            data_inicio_certa = datetime.strptime(data_inicio, '%d/%m/%Y').date()

            data_fim = self.date_Final.text()
            data_fim_certa = datetime.strptime(data_fim, '%d/%m/%Y').date()

            dif = data_fim_certa - data_inicio_certa
            dif1 = dif.days

            if dif1 != 0:
                for data in range(0, dif1 + 1):
                    data_muda = data_inicio_certa + timedelta(days=data)

                    results0 = self.select_movimentis_almox(data_inicio_certa, data_fim_certa)

                    if not results0:
                        results = []
                    else:
                        movimento, ops_entradas = self.tititutu_almox(data_muda, data_muda)
                        for dados in movimento:
                            results.append(dados)
            else:
                data_muda = data_inicio_certa

                results0 = self.select_movimentis_almox(data_inicio_certa, data_fim_certa)

                if not results0:
                    msg = f'Neste período não houve movimentações!'
                    self.label_Msg.setText(msg)
                    self.widget_Progress.setHidden(True)
                    results = []
                else:
                    movimento, ops_entradas = self.tititutu_almox(data_muda, data_muda)
                    for dados in movimento:
                        results.append(dados)

            self.widget_Progress.setHidden(True)

            return results

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def define_tabela_total(self):
        try:
            results = []

            data_inicio = self.date_Inicio.text()
            data_inicio_certa = datetime.strptime(data_inicio, '%d/%m/%Y').date()

            data_fim = self.date_Final.text()
            data_fim_certa = datetime.strptime(data_fim, '%d/%m/%Y').date()

            dif = data_fim_certa - data_inicio_certa
            dif1 = dif.days

            if dif1 != 0:
                for data in range(0, dif1 + 1):
                    data_muda = data_inicio_certa + timedelta(days=data)

                    results0 = self.select_movimentis_total(data_inicio_certa, data_fim_certa)

                    if not results0:
                        results = []
                    else:
                        movimento, ops_entradas = self.tititutu_total(data_muda, data_muda)
                        for dados in movimento:
                            results.append(dados)
            else:
                data_muda = data_inicio_certa

                results0 = self.select_movimentis_total(data_inicio_certa, data_fim_certa)

                if not results0:
                    msg = f'Neste período não houve movimentações!'
                    self.label_Msg.setText(msg)
                    self.widget_Progress.setHidden(True)
                    results = []
                else:
                    movimento, ops_entradas = self.tititutu_total(data_muda, data_muda)
                    for dados in movimento:
                        results.append(dados)

            self.widget_Progress.setHidden(True)

            return results

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def select_movimentis_almox(self, data_inicio, data_fim):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE((extract(day from m.data)||'/'||"
                           f"extract(month from m.data)||'/'||extract(year from m.data)), '') AS DATA, "
                           f"produto.codigo, produto.descricao, "
                           f"COALESCE(produto.obs, ''), produto.unidade, "
                           f"COALESCE(CASE WHEN m.tipo < 200 THEN m.quantidade END, 0) AS Qtde_Entrada, "
                           f"COALESCE(CASE WHEN m.tipo > 200 THEN m.quantidade END, 0) AS Qtde_Saida, "
                           f"(select case when sum(quantidade) is null then 0 else sum(quantidade) end "
                           f"from movimentacao where produto=m.produto "
                           f"and tipo<200 and localestoque=m.localestoque)-"
                           f"(select case when sum(quantidade) is null then 0 else sum(quantidade) end "
                           f"from movimentacao where produto=m.produto "
                           f"and tipo>200 and localestoque=m.localestoque)+"
                           f"(case when ((select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"(((m.tipo<200) and ((m2.data>m.data) or ((m2.data=m.data) "
                           f"and (m2.id>m.id)))) or(m.tipo>200 and m2.data>m.data)) "
                           f"and m2.tipo<200)*-1) is null then 0 else "
                           f"((select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"(((m.tipo<200) and ((m2.data>m.data) or((m2.data=m.data) "
                           f"and (m2.id>m.id)))) or(m.tipo>200 and m2.data>m.data)) "
                           f"and m2.tipo<200)*-1) end) + "
                           f"(case when (select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"((m2.data=m.data and (m2.id>m.id  or (m.tipo<200)) )or(m2.data>m.data)) "
                           f"and m2.tipo>200) is null then 0 else (select sum(m2.quantidade) "
                           f"from movimentacao m2 where m2.localestoque=m.localestoque "
                           f"and m2.produto=m.produto and ((m2.data=m.data "
                           f"and (m2.id>m.id or (m.tipo<200)) )or(m2.data>m.data)) and m2.tipo>200) end) "
                           f"as saldo, "
                           f"CASE WHEN m.tipo = 210 THEN ('OP '|| produtoos.numero) "
                           f"WHEN m.tipo = 110 THEN ('OP '|| ordemservico.numero) "
                           f"WHEN m.tipo = 130 THEN ('NF '|| entradaprod.nota) "
                           f"WHEN m.tipo = 230 THEN ('NF '|| saidaprod.numero) "
                           f"WHEN m.tipo = 250 THEN ('Devol. OS '|| produtoservico.numero) "
                           f"WHEN m.tipo = 112 THEN ('OS '|| produtoservico.numero) "
                           f"WHEN m.tipo = 220 THEN 'CI' "
                           f"END AS OS_NF_CI, "
                           f"COALESCE(natop.descricao, ''), localestoque.nome, "
                           f"COALESCE(funcionarios.funcionario, ''), COALESCE(m.obs, '') "
                           f"FROM movimentacao m "
                           f"INNER JOIN produto ON (m.codigo = produto.codigo) "
                           f"INNER JOIN localestoque ON (m.localestoque = localestoque.id) "
                           f"LEFT JOIN funcionarios ON (m.funcionario = funcionarios.id) "
                           f"LEFT JOIN saidaprod ON (m.id = saidaprod.movimentacao) "
                           f"LEFT JOIN entradaprod ON (m.id = entradaprod.movimentacao) "
                           f"LEFT JOIN produtoservico ON (m.id = produtoservico.movimentacao) "
                           f"LEFT JOIN ordemservico ON (m.id = ordemservico.movimentacao) "
                           f"LEFT JOIN produtoos ON (m.id = produtoos.movimentacao) "
                           f"LEFT JOIN natop ON (( COALESCE( saidaprod.natureza, 0 ) + "
                           f"COALESCE( entradaprod.natureza, 0 ) ) = natop.ID) "
                           f"WHERE m.data >= '{data_inicio}' and m.data <= '{data_fim}' "
                           f"and localestoque.id IN (1, 2) "
                           f"order by m.data, Qtde_Entrada, m.id;")
            select_movis = cursor.fetchall()

            return select_movis

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def select_movimentis_total(self, data_inicio, data_fim):
        try:
            print("total")
            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE((extract(day from m.data)||'/'||"
                           f"extract(month from m.data)||'/'||extract(year from m.data)), '') AS DATA, "
                           f"produto.codigo, produto.descricao, "
                           f"COALESCE(produto.obs, ''), produto.unidade, "
                           f"COALESCE(CASE WHEN m.tipo < 200 THEN m.quantidade END, 0) AS Qtde_Entrada, "
                           f"COALESCE(CASE WHEN m.tipo > 200 THEN m.quantidade END, 0) AS Qtde_Saida, "
                           f"(select case when sum(quantidade) is null then 0 else sum(quantidade) end "
                           f"from movimentacao where produto=m.produto "
                           f"and tipo<200 and localestoque=m.localestoque)-"
                           f"(select case when sum(quantidade) is null then 0 else sum(quantidade) end "
                           f"from movimentacao where produto=m.produto "
                           f"and tipo>200 and localestoque=m.localestoque)+"
                           f"(case when ((select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"(((m.tipo<200) and ((m2.data>m.data) or ((m2.data=m.data) "
                           f"and (m2.id>m.id)))) or(m.tipo>200 and m2.data>m.data)) "
                           f"and m2.tipo<200)*-1) is null then 0 else "
                           f"((select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"(((m.tipo<200) and ((m2.data>m.data) or((m2.data=m.data) "
                           f"and (m2.id>m.id)))) or(m.tipo>200 and m2.data>m.data)) "
                           f"and m2.tipo<200)*-1) end) + "
                           f"(case when (select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"((m2.data=m.data and (m2.id>m.id  or (m.tipo<200)) )or(m2.data>m.data)) "
                           f"and m2.tipo>200) is null then 0 else (select sum(m2.quantidade) "
                           f"from movimentacao m2 where m2.localestoque=m.localestoque "
                           f"and m2.produto=m.produto and ((m2.data=m.data "
                           f"and (m2.id>m.id or (m.tipo<200)) )or(m2.data>m.data)) and m2.tipo>200) end) "
                           f"as saldo, "
                           f"CASE WHEN m.tipo = 210 THEN ('OP '|| produtoos.numero) "
                           f"WHEN m.tipo = 110 THEN ('OP '|| ordemservico.numero) "
                           f"WHEN m.tipo = 130 THEN ('NF '|| entradaprod.nota) "
                           f"WHEN m.tipo = 230 THEN ('NF '|| saidaprod.numero) "
                           f"WHEN m.tipo = 250 THEN ('Devol. OS '|| produtoservico.numero) "
                           f"WHEN m.tipo = 112 THEN ('OS '|| produtoservico.numero) "
                           f"WHEN m.tipo = 220 THEN 'CI' "
                           f"END AS OS_NF_CI, "
                           f"COALESCE(natop.descricao, ''), localestoque.nome, "
                           f"COALESCE(funcionarios.funcionario, ''), COALESCE(m.obs, '') "
                           f"FROM movimentacao m "
                           f"INNER JOIN produto ON (m.codigo = produto.codigo) "
                           f"INNER JOIN localestoque ON (m.localestoque = localestoque.id) "
                           f"LEFT JOIN funcionarios ON (m.funcionario = funcionarios.id) "
                           f"LEFT JOIN saidaprod ON (m.id = saidaprod.movimentacao) "
                           f"LEFT JOIN entradaprod ON (m.id = entradaprod.movimentacao) "
                           f"LEFT JOIN produtoservico ON (m.id = produtoservico.movimentacao) "
                           f"LEFT JOIN ordemservico ON (m.id = ordemservico.movimentacao) "
                           f"LEFT JOIN produtoos ON (m.id = produtoos.movimentacao) "
                           f"LEFT JOIN natop ON (( COALESCE( saidaprod.natureza, 0 ) + "
                           f"COALESCE( entradaprod.natureza, 0 ) ) = natop.ID) "
                           f"WHERE m.data >= '{data_inicio}' and m.data <= '{data_fim}' "
                           f"order by m.data, Qtde_Entrada, m.id;")
            select_movis = cursor.fetchall()

            return select_movis

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tipos_movimentos(self):
        try:
            ent_nf = 130
            ent_nf_130 = "entradaprod.nota"

            ent_op = 110
            ent_op_110 = "ordemservico.numero"

            ent_os = 112
            ent_os_112 = "produtoservico.numero"

            ent_inv = 140
            ent_inv_140 = "produtoservico.numero"

            sai_nf = 230
            sai_nf_230 = "saidaprod.numero"

            sai_devolucao = 250
            sai_devolucao_250 = "produtoservico.numero"

            sai_op = 210
            sai_op_210 = "produtoos.numero"

            sai_ci = 220
            sai_ci_220 = "produtoos.numero"

            sai_inv = 240
            sai_inv_240 = "produtoos.numero"

            return ent_nf, ent_nf_130, ent_op, ent_op_110, ent_os, ent_os_112, sai_nf, \
                   sai_nf_230, ent_inv, ent_inv_140, sai_devolucao, sai_devolucao_250, \
                   sai_op, sai_op_210, sai_ci, sai_ci_220, sai_inv, sai_inv_240

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def datas_relatorio(self):
        try:
            data_inicio = self.date_Inicio.text()
            data_inicio_certa = datetime.strptime(data_inicio, '%d/%m/%Y').date()

            dia_inicio = data_inicio_certa.strftime("%d")

            data_fim = self.date_Final.text()
            data_fim_certa = datetime.strptime(data_fim, '%d/%m/%Y').date()

            dia_fim = data_fim_certa.strftime("%d")
            mes_fim = data_fim_certa.strftime("%m")
            ano_fim = data_fim_certa.strftime("%Y")
            data_str = f"{dia_fim}/{mes_fim}/{ano_fim}"

            if mes_fim == "01":
                mes_certo = "Janeiro"
            elif mes_fim == "02":
                mes_certo = "Fevereiro"
            elif mes_fim == "03":
                mes_certo = "Marco"
            elif mes_fim == "04":
                mes_certo = "Abril"
            elif mes_fim == "05":
                mes_certo = "Maio"
            elif mes_fim == "06":
                mes_certo = "Junho"
            elif mes_fim == "07":
                mes_certo = "Julho"
            elif mes_fim == "08":
                mes_certo = "Agosto"
            elif mes_fim == "09":
                mes_certo = "Setembro"
            elif mes_fim == "10":
                mes_certo = "Outubro"
            elif mes_fim == "11":
                mes_certo = "Novembro"
            elif mes_fim == "12":
                mes_certo = "Dezembro"
            else:
                mes_certo = ""

            return data_inicio_certa, data_fim_certa, dia_fim, mes_certo, ano_fim, dia_inicio, data_str

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tititutu_almox(self, data_inicial, data_final):
        try:
            ent_nf, ent_nf_130, ent_op, ent_op_110, ent_os, ent_os_112, sai_nf, sai_nf_230, \
            ent_inv, ent_inv_140, sai_devolucao, sai_devolucao_250, sai_op, sai_op_210, sai_ci, \
            sai_ci_220, sai_inv, sai_inv_240 = self.tipos_movimentos()

            results = []
            ops_entradas = []

            total = self.select_mov_almox(data_inicial, data_final, ent_nf, ent_nf_130)
            for dados in total:
                results.append(dados)

            total = self.select_mov_almox(data_inicial, data_final, ent_op, ent_op_110)
            for dados in total:
                ops_entradas.append(dados)
                results.append(dados)

            total = self.select_mov_almox(data_inicial, data_final, ent_os, ent_os_112)
            for dados in total:
                results.append(dados)

            total = self.select_mov_almox(data_inicial, data_final, ent_inv, ent_inv_140)
            for dados in total:
                results.append(dados)

            total = self.select_mov_almox(data_inicial, data_final, sai_nf, sai_nf_230)
            for dados in total:
                data, cod, descr, ref, um, ent, saida, saldo, os_nf_ci, cfop, local, solcic, obs = dados
                if not os_nf_ci:
                    pass
                else:
                    results.append(dados)

            total = self.select_mov_almox(data_inicial, data_final, sai_devolucao, sai_devolucao_250)
            for dados in total:
                results.append(dados)

            total = self.select_mov_almox(data_inicial, data_final, sai_op, sai_op_210)
            for dados in total:
                results.append(dados)

            total = self.select_mov_almox(data_inicial, data_final, sai_ci, sai_ci_220)
            for dados in total:
                results.append(dados)

            total = self.select_mov_almox(data_inicial, data_final, sai_inv, sai_inv_240)
            for dados in total:
                results.append(dados)

            return results, ops_entradas

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tititutu_total(self, data_inicial, data_final):
        try:
            ent_nf, ent_nf_130, ent_op, ent_op_110, ent_os, ent_os_112, sai_nf, sai_nf_230, \
            ent_inv, ent_inv_140, sai_devolucao, sai_devolucao_250, sai_op, sai_op_210, sai_ci, \
            sai_ci_220, sai_inv, sai_inv_240 = self.tipos_movimentos()

            results = []
            ops_entradas = []

            total = self.select_mov_total(data_inicial, data_final, ent_nf, ent_nf_130)
            for dados in total:
                results.append(dados)

            total = self.select_mov_total(data_inicial, data_final, ent_op, ent_op_110)
            for dados in total:
                ops_entradas.append(dados)
                results.append(dados)

            total = self.select_mov_total(data_inicial, data_final, ent_os, ent_os_112)
            for dados in total:
                results.append(dados)

            total = self.select_mov_total(data_inicial, data_final, ent_inv, ent_inv_140)
            for dados in total:
                results.append(dados)

            total = self.select_mov_total(data_inicial, data_final, sai_nf, sai_nf_230)
            for dados in total:
                data, cod, descr, ref, um, ent, saida, saldo, os_nf_ci, cfop, local, solcic, obs = dados
                if not os_nf_ci:
                    pass
                else:
                    results.append(dados)

            total = self.select_mov_total(data_inicial, data_final, sai_devolucao, sai_devolucao_250)
            for dados in total:
                results.append(dados)

            total = self.select_mov_total(data_inicial, data_final, sai_op, sai_op_210)
            for dados in total:
                results.append(dados)

            total = self.select_mov_total(data_inicial, data_final, sai_ci, sai_ci_220)
            for dados in total:
                results.append(dados)

            total = self.select_mov_total(data_inicial, data_final, sai_inv, sai_inv_240)
            for dados in total:
                results.append(dados)

            return results, ops_entradas

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def select_mistura_op(self, cod, num_op):
        try:
            dados_para_tabela = []
            campo_br = ""

            cur = conecta.cursor()
            cur.execute(f"SELECT id, descricao, id_versao FROM produto where codigo = {cod};")
            detalhes_produtos = cur.fetchall()

            id_prod, descricao, id_estrutura = detalhes_produtos[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, ' ') as obs, prod.unidade, "
                           f"((SELECT quantidade FROM ordemservico where numero = {num_op}) * "
                           f"(estprod.quantidade)) AS Qtde, "
                           f"prod.localizacao, prod.quantidade "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id_estrutura = {id_estrutura} ORDER BY prod.descricao;")
            select_estrut = cursor.fetchall()

            for dados_estrut in select_estrut:
                id_mat_e, cod_e, descr_e, ref_e, um_e, qtde_e, local_e, saldo_e = dados_estrut

                cursor = conecta.cursor()
                cursor.execute(f"SELECT max(estprod.id), max(prod.codigo), max(prod.descricao), "
                               f"sum(prodser.QTDE_ESTRUT_PROD)as total "
                               f"FROM estrutura_produto as estprod "
                               f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                               f"INNER JOIN produtoos as prodser ON estprod.id = prodser.id_estrut_prod "
                               f"where estprod.id_estrutura = {id_estrutura} "
                               f"and prodser.numero = {num_op} and estprod.id = {id_mat_e} "
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
                                   f"prodser.quantidade, prodser.QTDE_ESTRUT_PROD "
                                   f"from produtoos as prodser "
                                   f"INNER JOIN produto as prod ON prodser.produto = prod.id "
                                   f"where prodser.numero = {num_op} and prodser.id_estrut_prod = {id_mat_e};")
                    select_os = cursor.fetchall()

                    for dados_os in select_os:
                        id_mat_os, data_os, cod_os, descr_os, ref_os, um_os, qtde_os, qtde_mat_os = dados_os

                        dados2 = (id_mat_e, cod_e, descr_e, ref_e, um_e, qtde_mat_os, local_e, saldo_e,
                                  data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                        dados_para_tabela.append(dados2)

            itens_manipula_total = dados_para_tabela

            tabela_estrutura = []
            tabela_consumo_os = []

            for itens in itens_manipula_total:
                id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo, \
                data_os, cod_os, descr_os, ref_os, um_os, qtde_os = itens

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

            return tabela_consumo_os, tabela_estrutura

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excel_op(self, dados_para_op):
        try:
            emissao, num_op_int, cod, descr_op, ref_op, um_op, qtde_op_float, obs_op, consumo_os, estrutura, \
            caminho, aba_sheet = dados_para_op

            cod_op_int = int(cod)

            book = carregar_workbook(caminho)

            ws = book["Sheet2"]

            mp_copy = book.copy_worksheet(ws)
            mp_copy.title = aba_sheet

            camino = os.path.join('..', 'arquivos', "logo_pequeno.png")
            caminho_arquivo = definir_caminho_arquivo(camino)

            adiciona_imagem(mp_copy, caminho_arquivo, 'C3')

            dados_estrut = []
            dados_os_a = []
            dados_os_b = []
            dados_os_c = []
            total_qtde_mov = 0.00

            for tabi in estrutura:
                id_mat_est, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo = tabi
                qtde_e_float = float(qtde_est)
                dados = (cod_est, descr_est, ref_est, um_est, qtde_e_float)
                dados_estrut.append(dados)

            df = dataframe_pandas(dados_estrut, ['Cód.', 'Descrição', 'Referência', 'UM', 'Qtde'])

            codigo_int = {'Cód.': int}
            df = df.astype(codigo_int)
            qtde_float = {'Qtde': float}
            df = df.astype(qtde_float)

            for tabi2 in consumo_os:
                id_mat_os, data_os, cod_os, descr_os, ref_os, um_os, qtde_os = tabi2

                qtde_os_float = float(qtde_os)

                total_qtde_mov = total_qtde_mov + qtde_os_float

                dados1 = (emissao, cod_os)
                dados2 = descr_os
                dados3 = (ref_os, um_os, qtde_os)
                dados_os_a.append(dados1)
                dados_os_b.append(dados2)
                dados_os_c.append(dados3)

            df1 = dataframe_pandas(dados_os_a, ['Data', 'Cód.'])
            df2 = dataframe_pandas(dados_os_b, ['Descrição'])
            df3 = dataframe_pandas(dados_os_c, ['Referência', 'UM', 'Qtde'])

            codigo_int1 = {'Cód.': int}
            df1 = df1.astype(codigo_int1)

            qtde_float3 = {'Qtde': float}
            df3 = df3.astype(qtde_float3)

            writer = escritor_dataframe(caminho)

            writer.book = book

            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

            linhas_frame = df.shape[0]
            colunas_frame = df.shape[1]

            linhas_frame1 = df1.shape[0]
            colunas_frame1 = df1.shape[1]

            linhas_frame2 = df2.shape[0]
            colunas_frame2 = df2.shape[1]

            linhas_frame3 = df3.shape[0]
            colunas_frame3 = df3.shape[1]

            inicia = 14
            rows = range(inicia, inicia + linhas_frame)
            columns = range(2, colunas_frame + 2)

            for row in rows:
                for col in columns:
                    edita_alinhamento(mp_copy.cell(row, col), quebra_linha=True)
                    edita_bordas(mp_copy.cell(row, col))

            inicia1 = 14
            rows1 = range(inicia1, inicia1 + linhas_frame1)
            columns1 = range(8, colunas_frame1 + 8)

            for row1 in rows1:
                for col1 in columns1:
                    edita_alinhamento(mp_copy.cell(row1, col1), quebra_linha=True)
                    edita_bordas(mp_copy.cell(row1, col1))

            inicia2 = 14
            rows2 = range(inicia2, inicia2 + linhas_frame2)
            columns2 = range(10, colunas_frame2 + 11)

            for row2 in rows2:
                mp_copy.merge_cells(f"J{row2}:K{row2}")
                for col2 in columns2:
                    edita_alinhamento(mp_copy.cell(row2, col2), quebra_linha=True)
                    edita_bordas(mp_copy.cell(row2, col2))

            inicia3 = 14
            rows3 = range(inicia3, inicia3 + linhas_frame3)
            columns3 = range(12, colunas_frame3 + 12)

            for row3 in rows3:
                for col3 in columns3:
                    edita_alinhamento(mp_copy.cell(row3, col3), quebra_linha=True)
                    edita_bordas(mp_copy.cell(row3, col3))

            linhas_certas3 = linhas_frame3 + 14

            lanca_dados_coluna(mp_copy, "D6", emissao, 16, True)
            lanca_dados_mesclado(mp_copy, 'M4:N4', 'M4', num_op_int, 18, True)
            lanca_dados_coluna(mp_copy, "B9", cod_op_int, 12, False)
            lanca_dados_mesclado(mp_copy, 'C9:D9', 'C9', descr_op, 12, False)
            lanca_dados_mesclado(mp_copy, 'E9:H9', 'E9', ref_op, 12, False)
            lanca_dados_coluna(mp_copy, "I9", um_op, 12, False)
            lanca_dados_coluna(mp_copy, "J9", qtde_op_float, 12, False)
            if obs_op:
                lanca_dados_mesclado(mp_copy, 'K9:N9', 'K9', obs_op, 12, False)
            else:
                lanca_dados_mesclado(mp_copy, 'K9:N9', 'K9', "", 12, False)

            lanca_dados_coluna(mp_copy, f'N{linhas_certas3}', total_qtde_mov, 12, True)
            mp_copy[f'N{linhas_certas3}'].number_format = '0.00'

            lanca_dados_mesclado(mp_copy, f'L{linhas_certas3}:M{linhas_certas3}', f'l{linhas_certas3}',
                                 "Total Mov.", 12, True)

            escritor_direto_dataframe(df, writer, aba_sheet, 13, 1, False, False)
            escritor_direto_dataframe(df1, writer, aba_sheet, 13, 7, False, False)
            escritor_direto_dataframe(df2, writer, aba_sheet, 13, 9, False, False)
            escritor_direto_dataframe(df3, writer, aba_sheet, 13, 11, False, False)

            writer.save()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def remove_modelo_op(self, caminho):
        try:
            wb = carregar_workbook(caminho)
            if 'Sheet2' in wb.sheetnames:
                wb.remove(wb['Sheet2'])
            wb.save(caminho)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excel_mov(self, dados_tabela, arquivo_modelo, caminho, aba_sheet):
        try:
            df = dataframe_pandas(dados_tabela, ['Data', 'Cód.', 'Descrição', 'Referência', 'UM',
                                                 'Entrada', 'Saída', 'Saldo', 'OS/NF/CI', 'CFOP',
                                                 'Local', 'Solicitante', 'OBS'])

            codigo_int = {'Cód.': int}
            df = df.astype(codigo_int)
            entrada_float = {'Entrada': float}
            df = df.astype(entrada_float)
            saida_float = {'Saída': float}
            df = df.astype(saida_float)
            saldo_float = {'Saldo': float}
            df = df.astype(saldo_float)

            book = carregar_workbook(arquivo_modelo)

            ws = book["Sheet1"]
            ws.title = aba_sheet

            writer = escritor_dataframe(caminho)

            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

            linhas_frame = df.shape[0]
            colunas_frame = df.shape[1]

            ws = book.active

            inicia = 6
            rows = range(inicia, inicia + linhas_frame)
            columns = range(1, colunas_frame + 1)

            for row in rows:
                for col in columns:
                    edita_alinhamento(ws.cell(row, col), quebra_linha=True)
                    edita_bordas(ws.cell(row, col))

            escritor_direto_dataframe(df, writer, aba_sheet, 5, 0, False, False)

            writer.save()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def final(self):
        try:
            self.label_Excel.setText("")

            self.widget_Progress.setHidden(False)

            Thread(target=self.final1).start()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def final1(self):
        try:
            data_inicio, data_fim, dia_fim, mes, ano_fim, dia_inicio, data_str = self.datas_relatorio()

            if data_inicio == data_fim:
                novo = f'\Mov {dia_inicio} de {mes}.xlsx'
            else:
                novo = f'\Mov {dia_inicio} até {dia_fim} de {mes}.xlsx'

            dados_tabela = extrair_tabela(self.table_OP)

            camino = os.path.join('..', 'arquivos', 'modelo excel', 'est_mov.xlsx')
            caminho_arquivo = definir_caminho_arquivo(camino)

            aba = 'Movimentação'

            desktop = Path.home() / "Desktop"
            desk_str = str(desktop)
            caminho = (desk_str + novo)

            self.excel_mov(dados_tabela, caminho_arquivo, caminho, aba)

            for dados in dados_tabela:
                data, cod, descr, ref, um, entrada, saida, saldo, op, cfop, local, solicitante, obs = dados

                if entrada != "0":
                    posicao = op.find("OP ")

                    if posicao == 0:
                        inicio = posicao + 3
                        escolha = op[inicio:]

                        op_int = int(escolha)

                        cur = conecta.cursor()
                        cur.execute(f"SELECT id, numero, datainicial, status, produto, quantidade, obs "
                                    f"FROM ordemservico where numero = {op_int};")
                        extrair_dados = cur.fetchall()
                        id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, obs_2 = extrair_dados[0]

                        tabela_consumo_os, tabela_estrutura = self.select_mistura_op(cod, op_int)

                        lista_para_op = [data, op_int, cod, descr, ref, um, entrada, obs_2,
                                         tabela_consumo_os, tabela_estrutura, caminho, op]

                        self.excel_op(lista_para_op)

            self.remove_modelo_op(caminho)

            self.widget_Progress.setHidden(True)

            self.label_Excel.setText("Planilha criada com Sucesso!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    relmov = TelaEstMovimentacaoV2()
    relmov.show()
    qt.exec_()

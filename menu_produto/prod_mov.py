import sys
from banco_dados.conexao import conecta
from forms.tela_prod_mov import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import lanca_tabela, layout_cabec_tab
from comandos.telas import tamanho_aplicacao, icone
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import inspect
import os
import traceback
from threading import Thread


class TelaProdutoMovimentacao(QMainWindow, Ui_MainWindow):
    def __init__(self, produto, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        if produto:
            self.verifica_codigo(produto)

        icone(self, "menu_cadastro.png")
        tamanho_aplicacao(self)

        layout_cabec_tab(self.table_Mov)

        self.definir_line_bloqueados()

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

    def definir_line_bloqueados(self):
        try:
            self.line_Codigo.setReadOnly(True)
            self.line_Descricao.setReadOnly(True)
            self.line_DescrCompl.setReadOnly(True)
            self.line_Referencia.setReadOnly(True)
            self.line_UM.setReadOnly(True)
            self.line_Conjunto.setReadOnly(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_codigo(self, codigo_produto):
        try:
            cur = conecta.cursor()
            cur.execute(f"SELECT prod.descricao, COALESCE(prod.descricaocomplementar, ''), "
                        f"COALESCE(prod.obs, ''), prod.unidade, COALESCE(prod.localizacao, ''), conj.conjunto, "
                        f"prod.quantidade "
                        f"FROM produto as prod "
                        f"LEFT JOIN conjuntos conj ON prod.conjunto = conj.id "
                        f"where prod.codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            if detalhes_produto:
                descr, compl, ref, um, local, conjunto, saldo = detalhes_produto[0]

                self.line_Codigo.setText(codigo_produto)
                self.line_Descricao.setText(descr)
                self.line_DescrCompl.setText(compl)
                self.line_Referencia.setText(ref)
                self.line_UM.setText(um)
                self.line_Conjunto.setText(conjunto)
                self.line_Saldo.setText(str(saldo))

                Thread(target=self.manipula_dados_tabela_mov(codigo_produto)).start()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_mov(self, codigo_produto):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT m.data, "
                           f"CASE WHEN m.tipo < 200 THEN m.quantidade END AS Qtde_Entrada, "
                           f"CASE WHEN m.tipo > 200 THEN m.quantidade END AS Qtde_Saida, "
                           f"(select case when sum(quantidade) is null then 0 else sum(quantidade) end "
                           f"from movimentacao where produto=m.produto "
                           f"and tipo<200 and localestoque=m.localestoque)-(select case when sum(quantidade) "
                           f"is null then 0 else sum(quantidade) end "
                           f"from movimentacao where produto=m.produto "
                           f"and tipo>200 and localestoque=m.localestoque)+(case when ((select sum(m2.quantidade) "
                           f"from movimentacao m2 where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"(((m.tipo<200) and ((m2.data>m.data) or ((m2.data=m.data) and (m2.id>m.id)))) "
                           f"or(m.tipo>200 and m2.data>m.data)) and m2.tipo<200)*-1) is null then 0 else "
                           f"((select sum(m2.quantidade) from movimentacao m2 where m2.localestoque=m.localestoque "
                           f"and m2.produto=m.produto and "
                           f"(((m.tipo<200) and ((m2.data>m.data) or((m2.data=m.data) and (m2.id>m.id)))) "
                           f"or(m.tipo>200 and m2.data>m.data)) and m2.tipo<200)*-1) end) + "
                           f"(case when (select sum(m2.quantidade) from movimentacao m2 "
                           f"where m2.localestoque=m.localestoque and m2.produto=m.produto and "
                           f"((m2.data=m.data and (m2.id>m.id  or (m.tipo<200)) )or(m2.data>m.data)) "
                           f"and m2.tipo>200) is null then 0 else (select sum(m2.quantidade) "
                           f"from movimentacao m2 where m2.localestoque=m.localestoque and m2.produto=m.produto "
                           f"and ((m2.data=m.data and (m2.id>m.id or (m.tipo<200)) )or(m2.data>m.data)) "
                           f"and m2.tipo>200) end) "
                           f"as saldo, "
                           f"CASE WHEN m.tipo = 210 THEN ('OP '|| produtoos.numero) "
                           f"WHEN m.tipo = 110 THEN ('OP '|| ordemservico.numero) "
                           f"WHEN m.tipo = 111 THEN ('DEV. OP '|| produtoos.numero) "
                           f"WHEN m.tipo = 130 THEN ('NF '|| entradaprod.nota) "
                           f"WHEN m.tipo = 140 THEN ('INVENTÁRIO') "
                           f"WHEN m.tipo = 240 THEN ('INVENTÁRIO') "
                           f"WHEN m.tipo = 230 THEN ('NF '|| saidaprod.numero) "
                           f"WHEN m.tipo = 250 THEN ('Devol. OS '|| produtoservico.numero) "
                           f"WHEN m.tipo = 112 THEN ('OS '|| produtoservico.numero) "
                           f"WHEN m.tipo = 220 THEN 'CI' "
                           f"END AS OS_NF_CI, "
                           f"natop.descricao as CFOP, localestoque.nome, "
                           f"COALESCE(m.obs, ''), m.tipo, "
                           f"CASE WHEN m.tipo = 130 THEN ('OC '|| occ.numero) "
                           f"WHEN m.tipo = 230 THEN ('OV '|| ocs.numero) "
                           f"END as op_ov, "
                           f"CASE WHEN m.tipo = 210 THEN (funcop.funcionario) "
                           f"WHEN m.tipo = 110 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 111 THEN (funcop.funcionario) "
                           f"WHEN m.tipo = 130 THEN (fornecedores.razao) "
                           f"WHEN m.tipo = 140 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 230 THEN (clientes.razao) "
                           f"WHEN m.tipo = 250 THEN (funcionarios.funcionario) "
                           f"WHEN m.tipo = 112 THEN (funcos.funcionario) "
                           f"WHEN m.tipo = 220 THEN (funcionarios.funcionario) "
                           f"END AS empresa_func "
                           f"FROM movimentacao m "
                           f"INNER JOIN produto ON (m.codigo = produto.codigo) "
                           f"INNER JOIN localestoque ON (m.localestoque = localestoque.id) "
                           f"LEFT JOIN funcionarios ON (m.funcionario = funcionarios.id) "
                           f"LEFT JOIN saidaprod ON (m.id = saidaprod.movimentacao) "
                           f"LEFT JOIN entradaprod ON (m.id = entradaprod.movimentacao) "
                           f"LEFT JOIN produtoservico ON (m.id = produtoservico.movimentacao) "
                           f"LEFT JOIN ordemservico ON (m.id = ordemservico.movimentacao) "
                           f"LEFT JOIN produtoos ON (m.id = produtoos.movimentacao) "
                           f"LEFT JOIN funcionarios as funcop ON (produtoos.funcionarios = funcop.id) "
                           f"LEFT JOIN funcionarios as funcos ON (produtoservico.funcionarios = funcos.id) "
                           f"LEFT JOIN ORDEMCOMPRA ocs ON saidaprod.ordemcompra = ocs.id "
                           f"LEFT JOIN ORDEMCOMPRA occ ON entradaprod.ordemcompra = occ.id "
                           f"LEFT JOIN fornecedores ON (entradaprod.fornecedor = fornecedores.id) "
                           f"LEFT JOIN clientes ON (saidaprod.cliente = clientes.id) "
                           f"LEFT JOIN natop ON (( COALESCE( saidaprod.natureza, 0 ) + "
                           f"COALESCE( entradaprod.natureza, 0 ) ) = natop.ID) "
                           f"WHERE m.codigo = '{codigo_produto}' "
                           f"order by m.data, (case when m.tipo >= 200 then 2 else 1 end), m.id;")
            results = cursor.fetchall()
            if results:
                for i in results:
                    data, entrada, saida, saldo, registro, cfop, local_est, obs, tipo, op_ov, empresa_func = i

                    data_final = f'{data.day}/{data.month}/{data.year}'

                    if entrada:
                        ent = entrada
                    else:
                        ent = ""

                    if saida:
                        sai = saida
                    else:
                        sai = ""

                    if registro:
                        reg = registro
                    else:
                        reg = ""

                    if cfop:
                        natur = cfop
                    else:
                        natur = ""

                    if empresa_func:
                        empresa = empresa_func
                    else:
                        empresa = ""
                    if op_ov:
                        ordens = op_ov
                    else:
                        ordens = ""

                    dados = (data_final, local_est, ent, sai, saldo, reg, ordens, natur, empresa, obs)
                    tabela_nova.append(dados)

            if tabela_nova:
                lanca_tabela(self.table_Mov, tabela_nova)

                self.table_Mov.scrollToBottom()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaProdutoMovimentacao("")
    tela.show()
    qt.exec_()

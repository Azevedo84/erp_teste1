import sys
from banco_dados.conexao import conecta
from forms.tela_pcp_produto import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import lanca_tabela, layout_cabec_tab
from comandos.telas import tamanho_aplicacao, icone
from comandos.cores import cor_verde, widgets, cor_vermelho
from comandos.lines import validador_so_numeros
from comandos.conversores import valores_para_float
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import inspect
import os
import traceback


class TelaPcpProdutoV2(QMainWindow, Ui_MainWindow):
    def __init__(self, produto, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_producao.png")
        tamanho_aplicacao(self)

        layout_cabec_tab(self.table_Producao)
        layout_cabec_tab(self.table_Compra)
        layout_cabec_tab(self.table_Consumo)
        layout_cabec_tab(self.table_Venda)
        layout_cabec_tab(self.table_Estoque)
        layout_cabec_tab(self.table_Estrutura)
        layout_cabec_tab(self.table_Usado)
        layout_cabec_tab(self.table_Mov)

        validador_so_numeros(self.line_Codigo)

        self.showMaximized()

        self.line_Codigo.editingFinished.connect(self.verifica_line_codigo_manual)

        self.processando = False

        if produto:
            self.line_Codigo.setText(produto)
            self.line_Codigo.setReadOnly(True)
            self.verifica_line_codigo_manual()

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
            self.line_Descricao.setReadOnly(True)
            self.line_DescrCompl.setReadOnly(True)
            self.line_Referencia.setReadOnly(True)
            self.line_UM.setReadOnly(True)
            self.line_Saldo_Total.setReadOnly(True)
            self.line_Local.setReadOnly(True)
            self.line_NCM.setReadOnly(True)
            self.line_Conjunto.setReadOnly(True)
            self.line_Tipo.setReadOnly(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_codigo_manual(self):
        if not self.processando:
            try:
                self.processando = True

                codigo_produto = self.line_Codigo.text()

                if not codigo_produto:
                    self.mensagem_alerta('O campo "Código" não pode estar vazio')
                    self.limpa_dados_produto()
                    self.limpa_tabelas()
                elif int(codigo_produto) == 0:
                    self.mensagem_alerta('O campo "Código" não pode ser "0"')
                    self.limpa_dados_produto()
                    self.limpa_tabelas()
                else:
                    self.verifica_sql_produto_manual()
                    self.table_Estoque.setFocus()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_sql_produto_manual(self):
        try:
            codigo_produto = self.line_Codigo.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade, localizacao, quantidade "
                           f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cursor.fetchall()
            if not detalhes_produto:
                self.mensagem_alerta('Este código de produto não existe!')
                self.limpa_dados_produto()
                self.limpa_tabelas()
            else:
                self.lanca_dados_produto_manual()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_produto_manual(self):
        try:
            codigo_produto = self.line_Codigo.text()
            cur = conecta.cursor()
            cur.execute(f"SELECT prod.id, prod.descricao, "
                        f"COALESCE(prod.descricaocomplementar, '') as compl, "
                        f"COALESCE(prod.obs, '') as obs, "
                        f"prod.unidade, "
                        f"COALESCE(prod.localizacao, '') as local, "
                        f"prod.ncm, "
                        f"prod.quantidade, prod.embalagem, COALESCE(prod.kilosmetro, '') as kilos, conj.conjunto, "
                        f"COALESCE(tip.tipomaterial, '') as tips, prod.id_versao "
                        f"FROM produto as prod "
                        f"LEFT JOIN conjuntos conj ON prod.conjunto = conj.id "
                        f"LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id "
                        f"where prod.codigo = {codigo_produto};")
            detal = cur.fetchall()
            if detal:
                id_prod, descr, compl, ref, um, local, ncm, saldo, embal, kg_mt, conjunto, tipo, id_versao = detal[0]

                self.line_Descricao.setText(descr)
                self.line_DescrCompl.setText(compl)
                self.line_Referencia.setText(ref)
                self.line_UM.setText(um)
                self.line_Saldo_Total.setText(str(saldo))
                self.line_Local.setText(local)
                self.line_Conjunto.setText(conjunto)
                self.line_Tipo.setText(tipo)

                if not ncm:
                    self.line_NCM.setText("")
                    self.line_NCM.setStyleSheet("QLineEdit { background-color: yellow; }")
                else:
                    self.line_NCM.setText(ncm)
                    self.line_NCM.setStyleSheet("QLineEdit { background-color: white; }")

                self.limpa_tabelas()

                self.manipula_dados_tabela_estoque(id_prod)
                self.manipula_dados_tabela_producao(codigo_produto)
                self.manipula_dados_tabela_mov(id_prod)
                self.manipula_dados_tabela_venda(codigo_produto)
                self.manipula_dados_tabela_compra(codigo_produto)
                self.manipula_dados_tabela_consumo(codigo_produto)
                self.manipula_dados_tabela_usado(codigo_produto)

                if id_versao:
                    self.manipula_dados_tabela_estrutura(codigo_produto)
            else:
                self.mensagem_alerta("Este cadastro de produto não existe!")
                self.line_Codigo.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tabelas(self):
        try:
            self.table_Estoque.setRowCount(0)
            self.table_Producao.setRowCount(0)
            self.table_Estrutura.setRowCount(0)
            self.table_Mov.setRowCount(0)
            self.table_Venda.setRowCount(0)
            self.table_Consumo.setRowCount(0)
            self.table_Compra.setRowCount(0)
            self.table_Usado.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_produto(self):
        try:
            self.line_Codigo.clear()

            self.line_Descricao.clear()
            self.line_Referencia.clear()
            self.line_UM.clear()
            self.line_Saldo_Total.clear()
            self.line_Local.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_estoque(self, id_prod):
        try:
            cur = conecta.cursor()
            cur.execute(f"SELECT loc.nome, sald.saldo FROM SALDO_ESTOQUE as sald "
                        f"INNER JOIN LOCALESTOQUE loc ON sald.local_estoque = loc.id "
                        f"where sald.produto_id = {id_prod} order by loc.nome;")
            detalhes_saldos = cur.fetchall()
            if detalhes_saldos:
                lanca_tabela(self.table_Estoque, detalhes_saldos)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_producao(self, cod_prod):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.codigo, "
                           f"prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.id_estrutura "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where ordser.status = 'A' and prod.codigo = {cod_prod} "
                           f"order by ordser.numero;")
            op_abertas = cursor.fetchall()
            if op_abertas:
                op_ab_editado = []
                for dados_op in op_abertas:
                    emissao, previsao, op, cod, descr, ref, um, qtde, id_estrut = dados_op

                    if id_estrut:
                        data = f'{emissao.day}/{emissao.month}/{emissao.year}'

                        total_estrut = 0
                        total_consumo = 0

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT estprod.id, "
                                       f"((SELECT quantidade FROM ordemservico where numero = {op}) * "
                                       f"(estprod.quantidade)) AS Qtde "
                                       f"FROM estrutura_produto as estprod "
                                       f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                                       f"where estprod.id_estrutura = {id_estrut};")
                        itens_estrutura = cursor.fetchall()

                        for dads in itens_estrutura:
                            ides, quantidade = dads
                            total_estrut += 1

                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT max(prodser.ID_ESTRUT_PROD), "
                                           f"sum(prodser.QTDE_ESTRUT_PROD) as total "
                                           f"FROM estrutura_produto as estprod "
                                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                                           f"INNER JOIN produtoos as prodser ON estprod.id = prodser.ID_ESTRUT_PROD "
                                           f"where prodser.numero = {op} and estprod.id = {ides} "
                                           f"group by prodser.ID_ESTRUT_PROD;")
                            itens_consumo = cursor.fetchall()
                            for duds in itens_consumo:
                                id_mats, qtde_mats = duds
                                if ides == id_mats and quantidade == qtde_mats:
                                    total_consumo += 1

                        msg = f"{total_estrut}/{total_consumo}"

                        dados = (data, op, qtde, msg)
                        op_ab_editado.append(dados)

                if op_ab_editado:
                    lanca_tabela(self.table_Producao, op_ab_editado)
                    self.widget_Producao.setStyleSheet(f"background-color: {cor_verde};")
                else:
                    self.widget_Producao.setStyleSheet(f"background-color: {widgets};")

            else:
                self.widget_Producao.setStyleSheet(f"background-color: {widgets};")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_estrutura(self, cod_prod):
        try:
            nova_tabela = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {cod_prod};")
            select_prod = cursor.fetchall()
            idez, cod, id_estrut = select_prod[0]

            if id_estrut:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                               f"conj.conjunto, prod.unidade, (estprod.quantidade * 1) as qtde, "
                               f"COALESCE(prod.ncm, '') as ncm "
                               f"from estrutura_produto as estprod "
                               f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                               f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                               f"where estprod.id_estrutura = {id_estrut} "
                               f"order by conj.conjunto DESC, prod.descricao ASC;")
                tabela_estrutura = cursor.fetchall()

                if tabela_estrutura:
                    for i in tabela_estrutura:
                        cod, descr, ref, conjunto, um, qtde, ncm = i

                        qtde_float = float(qtde)

                        dados = (cod, descr, ref, um, qtde_float)
                        nova_tabela.append(dados)

            if nova_tabela:
                lanca_tabela(self.table_Estrutura, nova_tabela)
                self.widget_Estrutura.setStyleSheet(f"background-color: {widgets};")

            else:
                conjunto = self.line_Conjunto.text()

                if conjunto == "Produtos Acabados":
                    self.widget_Estrutura.setStyleSheet(f"background-color: {cor_vermelho};")
                else:
                    self.widget_Estrutura.setStyleSheet(f"background-color: {widgets};")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_usado(self, cod_prod):
        try:
            planilha_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, estprod.id_estrutura, estprod.quantidade "
                           f"from estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f" where prod.codigo = {cod_prod};")
            tabela_estrutura = cursor.fetchall()
            for i in tabela_estrutura:
                ides_mat, id_estrutura, qtde = i

                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, ''), prod.unidade, "
                               f"COALESCE(prod.obs2, '') "
                               f"from estrutura as est "
                               f"INNER JOIN produto prod ON est.id_produto = prod.id "
                               f"where est.id = {id_estrutura};")
                select_prod = cursor.fetchall()

                if select_prod:
                    cod, descr, ref, um, obs = select_prod[0]
                    dados = (cod, descr, ref, um, qtde)
                    planilha_nova.append(dados)

            if planilha_nova:
                planilha_nova_ordenada = sorted(planilha_nova, key=lambda x: x[1])

                lanca_tabela(self.table_Usado, planilha_nova_ordenada)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_mov(self, id_prod):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT FIRST 35 mov.id, mov.data, mov.codigo, mov.tipo, mov.quantidade, "
                           f"COALESCE(func.funcionario, '') as fuck, loc.nome "
                           f"FROM movimentacao AS mov "
                           f"LEFT JOIN funcionarios AS func ON mov.funcionario = func.id "
                           f"INNER JOIN localestoque AS loc ON mov.localestoque = loc.id "
                           f"WHERE mov.produto = {id_prod} "
                           f"ORDER BY mov.data DESC, mov.id DESC;")
            dados_mov = cursor.fetchall()

            if dados_mov:
                for i in dados_mov:
                    id_mov, emissao, codigo, tipo, qtde, funcionario, loc_est = i

                    op_oc_ci_e = ''
                    operacao_e = ''
                    op_oc_ci_s = ''
                    operacao_s = ''
                    solicitante = ''

                    if tipo < 200:
                        qtde_entrada = str(qtde)
                        qtde_saida = ""
                    else:
                        qtde_saida = str(qtde)
                        qtde_entrada = ""

                    if tipo == 210:
                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT id, numero "
                                       f"from PRODUTOOS "
                                       f"where movimentacao = {id_mov};")
                        dados_produtoos = cursor.fetchall()
                        if dados_produtoos:
                            for prodos in dados_produtoos:
                                id_prodos, con_op = prodos
                                operacao_s = 'CONSUMO'
                                op_oc_ci_s = f'OP {con_op}'
                                solicitante = funcionario

                    elif tipo == 110:
                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT id, numero "
                                       f"from ordemservico "
                                       f"where movimentacao = {id_mov};")
                        dados_ordemservico = cursor.fetchall()
                        if dados_ordemservico:
                            for ordser in dados_ordemservico:
                                id_prodos, prod_op = ordser
                                operacao_e = 'PRODUÇÃO'
                                op_oc_ci_e = f'OP {prod_op}'
                                solicitante = funcionario

                    elif tipo == 130:
                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT ent.id, ent.nota, oc.numero, ent.natureza, nat.descricao, forn.razao "
                                       f"from ENTRADAPROD as ent "
                                       f"LEFT JOIN ORDEMCOMPRA oc ON ent.ordemcompra = oc.id "
                                       f"LEFT JOIN NATOP nat ON ent.natureza = nat.id "
                                       f"INNER JOIN FORNECEDORES forn ON ent.fornecedor = forn.id "
                                       f"where ent.movimentacao = {id_mov};")
                        dados_entrada = cursor.fetchall()
                        if dados_entrada:
                            for ent_prod in dados_entrada:
                                id_ent, nota_ent, oc_ent, id_nat_e, nat_ent, fornec = ent_prod

                                if id_nat_e == 4:
                                    operacao_e = "COMPRA"
                                    op_oc_ci_e = f"OC {oc_ent}"
                                elif id_nat_e == 6:
                                    operacao_e = "INDUSTR."
                                    op_oc_ci_e = f"OC {oc_ent}"
                                else:
                                    if nota_ent:
                                        operacao_e = f"NF ENTRADA"
                                        op_oc_ci_e = f"NF {nota_ent}"
                                    else:
                                        operacao_e = f"NF ENTRADA"

                                solicitante = fornec

                    elif tipo == 230:
                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT sai.id, sai.numero, oc.numero, sai.natureza, nat.descricao, clin.razao "
                                       f"from SAIDAPROD as sai "
                                       f"LEFT JOIN ORDEMCOMPRA oc ON sai.ordemcompra = oc.id "
                                       f"LEFT JOIN NATOP nat ON sai.natureza = nat.id "
                                       f"INNER JOIN CLIENTES clin ON sai.cliente = clin.id "
                                       f"where sai.movimentacao = {id_mov};")
                        dados_saida = cursor.fetchall()
                        if dados_saida:
                            for sai_prod in dados_saida:
                                id_sai, nota_sai, ov_sai, id_nat_s, nat_sai, cliente = sai_prod

                                if id_nat_s == 5 or id_nat_s == 11 or id_nat_s == 9 or id_nat_s == 12:
                                    operacao_s = "VENDA"
                                    op_oc_ci_s = f"OV {ov_sai}"
                                elif id_nat_s == 6 or id_nat_s == 49:
                                    operacao_s = "INDUSTR."
                                    op_oc_ci_s = f"OV {ov_sai}"
                                else:
                                    if nota_sai:
                                        operacao_s = f"NF SAÍDA"
                                        op_oc_ci_s = f"NF{nota_sai}"
                                    else:
                                        operacao_s = f"NF SAÍDA"

                                solicitante = cliente

                    elif tipo == 250:
                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT prodser.id, prodser.numero "
                                       f"from PRODUTOSERVICO as prodser "
                                       f"where prodser.movimentacao = {id_mov};")
                        dados_dev_os = cursor.fetchall()
                        if dados_dev_os:
                            for dev_os in dados_dev_os:
                                id_dev_os, dev_num = dev_os
                                operacao_s = f'DEVOLUÇÃO OS '
                                op_oc_ci_s = f'OS {dev_num}'
                                solicitante = funcionario

                    elif tipo == 112:
                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT prodser.id, prodser.numero "
                                       f"from PRODUTOSERVICO as prodser "
                                       f"where prodser.movimentacao = {id_mov};")
                        dados_con_os = cursor.fetchall()
                        if dados_con_os:
                            for con_os in dados_con_os:
                                id_con_os, os_num = con_os
                                operacao_e = f'CONSUMO OS'
                                op_oc_ci_e = f'OS {os_num}'
                                solicitante = funcionario

                    elif tipo == 220:
                        operacao_s = f'CONSUMO INTERNO'
                        op_oc_ci_s = 'CI'
                        solicitante = funcionario

                    elif tipo == 140:
                        operacao_e = f'INVENTÁRIO'
                        op_oc_ci_e = 'SCIV'
                        solicitante = funcionario

                    elif tipo == 240:
                        operacao_s = f'INVENTÁRIO'
                        op_oc_ci_s = 'SCIV'
                        solicitante = funcionario

                    else:
                        print("AINDA NÃO FOI DEFINIDO TIPO", i)

                    dados = (id_mov, emissao, operacao_e, op_oc_ci_e, qtde_entrada,
                             operacao_s, op_oc_ci_s, qtde_saida, solicitante, loc_est)
                    tabela_nova.append(dados)

            if tabela_nova:
                tabela_nova2 = []
                lista_de_listas_ordenada = sorted(tabela_nova, key=lambda x: (x[1], x[0]))

                for teste in lista_de_listas_ordenada:
                    id_muv, emis, oper_e, op_e, qtde_e, oper_s, op_s, qtde_s, soli, loc_ = teste

                    data = f'{emis.day}/{emis.month}/{emis.year}'
                    didis = (data, oper_e, op_e, qtde_e, oper_s, op_s, qtde_s, soli, loc_)
                    tabela_nova2.append(didis)
                lanca_tabela(self.table_Mov, tabela_nova2)
                self.table_Mov.scrollToBottom()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_venda(self, cod_prod):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.emissao, ped.id, cli.razao, prodint.qtde, "
                           f"prodint.data_previsao "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prodint.status = 'A' and prod.codigo = {cod_prod};")
            dados_pi = cursor.fetchall()
            if dados_pi:
                for i_pi in dados_pi:
                    emissao_pi, num_pi, clie_pi, qtde_pi, entrega_pi = i_pi

                    emi_pi = f'{emissao_pi.day}/{emissao_pi.month}/{emissao_pi.year}'
                    entreg_pi = f'{entrega_pi.day}/{entrega_pi.month}/{entrega_pi.year}'

                    dados_pi = (num_pi, "", emi_pi, clie_pi, qtde_pi, entreg_pi)
                    tabela_nova.append(dados_pi)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.data, oc.numero, cli.razao, prodoc.quantidade, prodoc.dataentrega, "
                           f"COALESCE(prodoc.id_pedido, '') as pedi "
                           f"FROM PRODUTOORDEMCOMPRA as prodoc "
                           f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                           f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                           f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                           f"LEFT JOIN pedidointerno as ped ON prodoc.id_pedido = ped.id "
                           f"where prodoc.quantidade > prodoc.produzido "
                           f"and oc.status = 'A' "
                           f"and oc.entradasaida = 'S' "
                           f"and prod.codigo = {cod_prod};")
            dados_ov = cursor.fetchall()
            if dados_ov:
                for i_ov in dados_ov:
                    emissao_ov, num_ov, clie_ov, qtde_ov, entrega_ov, num_pi_ov = i_ov

                    emi_ov = f'{emissao_ov.day}/{emissao_ov.month}/{emissao_ov.year}'
                    entreg_ov = f'{entrega_ov.day}/{entrega_ov.month}/{entrega_ov.year}'

                    dados = (num_pi_ov, num_ov, emi_ov, clie_ov, qtde_ov, entreg_ov)
                    tabela_nova.append(dados)

            if tabela_nova:
                lanca_tabela(self.table_Venda, tabela_nova)
                self.widget_Vendas.setStyleSheet(f"background-color: {cor_verde};")
            else:
                self.widget_Vendas.setStyleSheet(f"background-color: {widgets};")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_compra(self, cod_prod):
        try:
            tabela_nova = []

            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE(prodreq.mestre, ''), req.dataemissao, prodreq.quantidade "
                           f"FROM produtoordemsolicitacao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"INNER JOIN ordemsolicitacao as req ON prodreq.mestre = req.idsolicitacao "
                           f"LEFT JOIN produtoordemrequisicao as preq ON prodreq.id = preq.id_prod_sol "
                           f"WHERE prodreq.status = 'A' "
                           f"and prod.codigo = {cod_prod} "
                           f"AND preq.id_prod_sol IS NULL "
                           f"ORDER BY prodreq.mestre;")
            dados_sol = cursor.fetchall()

            if dados_sol:
                for i_sol in dados_sol:
                    num_sol, emissao_sol, qtde_sol = i_sol

                    emi_sol = f'{emissao_sol.day}/{emissao_sol.month}/{emissao_sol.year}'

                    dedos_sol = (num_sol, "", "", emi_sol, qtde_sol, "", "")
                    tabela_nova.append(dedos_sol)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT sol.idsolicitacao, prodreq.quantidade, req.data, prodreq.numero, "
                           f"prodreq.destino, prodreq.id_prod_sol "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"INNER JOIN ordemrequisicao as req ON prodreq.mestre = req.id "
                           f"INNER JOIN produtoordemsolicitacao as prodsol ON prodreq.id_prod_sol = prodsol.id "
                           f"INNER JOIN ordemsolicitacao as sol ON prodsol.mestre = sol.idsolicitacao "
                           f"where prodreq.status = 'A' "
                           f"and prod.codigo = {cod_prod};")
            dados_req = cursor.fetchall()

            if dados_req:
                for i_req in dados_req:
                    num_sol_req, qtde_req, emissao_req, num_req, destino, id_prod_sol = i_req

                    emi_req = f'{emissao_req.day}/{emissao_req.month}/{emissao_req.year}'

                    dedos_req = (num_sol_req, num_req, "", emi_req, qtde_req, "", "")
                    tabela_nova.append(dedos_req)

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT sol.idsolicitacao, prodreq.numero, oc.data, oc.numero, forn.razao, "
                f"prodoc.quantidade, prodoc.produzido, prodoc.dataentrega "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"INNER JOIN produtoordemrequisicao as prodreq ON prodoc.id_prod_req = prodreq.id "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"INNER JOIN fornecedores as forn ON oc.fornecedor = forn.id "
                f"INNER JOIN produtoordemsolicitacao as prodsol ON prodreq.id_prod_sol = prodsol.id "
                f"INNER JOIN ordemsolicitacao as sol ON prodsol.mestre = sol.idsolicitacao "
                f"where oc.entradasaida = 'E' "
                f"AND oc.STATUS = 'A' "
                f"AND prodoc.produzido < prodoc.quantidade "
                f"and prod.codigo = {cod_prod}"
                f"ORDER BY oc.numero;")
            dados_oc = cursor.fetchall()

            if dados_oc:
                for i_oc in dados_oc:
                    num_sol_oc, id_req_oc, emissao_oc, num_oc, forncec_oc, qtde_oc, prod_oc, entrega_oc = i_oc

                    emi_oc = f'{emissao_oc.day}/{emissao_oc.month}/{emissao_oc.year}'
                    ent_oc = f'{entrega_oc.day}/{entrega_oc.month}/{entrega_oc.year}'

                    dedos_oc = (num_sol_oc, id_req_oc, num_oc, emi_oc, qtde_oc, ent_oc, forncec_oc)
                    tabela_nova.append(dedos_oc)

            if tabela_nova:
                lanca_tabela(self.table_Compra, tabela_nova)
                self.widget_Compras.setStyleSheet(f"background-color: {cor_verde};")
            else:
                self.widget_Compras.setStyleSheet(f"background-color: {widgets};")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela_consumo(self, cod_prod):
        try:
            tabela_nova = []

            qtde_necessidade = 0

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, estprod.id_estrutura "
                           f"from estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where prod.codigo = {cod_prod};")
            dados_estrut = cursor.fetchall()
            for i_estrut in dados_estrut:
                ides_mat, id_estrutura = i_estrut

                cursor = conecta.cursor()
                cursor.execute(f"select id, id_produto, num_versao, data_versao, obs, data_criacao "
                               f"from estrutura where id = {id_estrutura};")
                estrutura = cursor.fetchall()

                id_produto = estrutura[0][1]

                cursor = conecta.cursor()
                cursor.execute(f"select ordser.datainicial, ordser.numero, prod.codigo, prod.descricao "
                               f"from ordemservico as ordser "
                               f"INNER JOIN produto prod ON ordser.produto = prod.id "
                               f"where ordser.status = 'A' "
                               f"and prod.id = {id_produto} "
                               f"order by ordser.numero;")
                op_abertas = cursor.fetchall()
                if op_abertas:
                    for ii in op_abertas:
                        emissao, num_op, cod_pai, descr_pai = ii

                        emis = f'{emissao.day}/{emissao.month}/{emissao.year}'

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT estprod.id, prod.codigo, "
                                       f"((SELECT quantidade FROM ordemservico where numero = {num_op}) * "
                                       f"(estprod.quantidade)) AS Qtde "
                                       f"FROM estrutura_produto as estprod "
                                       f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                                       f"where estprod.id = {ides_mat};")
                        select_estrut = cursor.fetchall()
                        if select_estrut:
                            id_mat, cod_estrut, qtde_total = select_estrut[0]

                            total_float = valores_para_float(qtde_total)

                            cursor = conecta.cursor()
                            cursor.execute(f"SELECT max(prod.codigo), max(prod.descricao), "
                                           f"sum(prodser.QTDE_ESTRUT_PROD) as total "
                                           f"FROM estrutura_produto as estprod "
                                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                                           f"INNER JOIN produtoos as prodser ON estprod.id = prodser.ID_ESTRUT_PROD "
                                           f"where estprod.id_estrutura = {id_estrutura} "
                                           f"and prodser.numero = {num_op} and estprod.id = {id_mat} "
                                           f"group by prodser.ID_ESTRUT_PROD;")
                            select_os_resumo = cursor.fetchall()
                            if select_os_resumo:
                                for os_cons in select_os_resumo:
                                    cod_cons, descr_cons, qtde_cons_total = os_cons

                                    dados = (emis, num_op, qtde_total, qtde_cons_total, cod_pai, descr_pai)
                                    tabela_nova.append(dados)

                                    cons_float = valores_para_float(qtde_cons_total)

                                    qtde_necessidade += total_float - cons_float
                            else:
                                dados = (emis, num_op, qtde_total, "0", cod_pai, descr_pai)
                                tabela_nova.append(dados)

                                qtde_necessidade += total_float

            self.label_Op_Nec.setText("")

            if tabela_nova:
                lanca_tabela(self.table_Consumo, tabela_nova)
                if qtde_necessidade > 0:
                    arred = round(qtde_necessidade, 2)
                    msg = f"Nec. {arred}"
                    self.label_Op_Nec.setText(msg)

                self.widget_Consumo.setStyleSheet(f"background-color: {cor_verde};")
            else:
                self.widget_Consumo.setStyleSheet(f"background-color: {widgets};")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaPcpProdutoV2("")
    tela.show()
    qt.exec_()

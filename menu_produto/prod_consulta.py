import sys
from banco_dados.conexao import conecta
from forms.tela_prod_consultar import *
from arquivos.chamar_arquivos import definir_caminho_arquivo
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import lanca_tabela, layout_cabec_tab
from comandos.telas import tamanho_aplicacao, icone, editar_botao
from comandos.lines import validador_inteiro
from comandos.conversores import float_para_virgula
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtGui import QIcon
import inspect
import os
import traceback
from datetime import date


class TelaProdutoConsultar(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        caminho = os.path.join('..', 'arquivos', 'icones', 'lupa.png')
        caminho_arquivo = definir_caminho_arquivo(caminho)
        icon = QIcon(caminho_arquivo)
        self.btn_Lupa_Prod.setIcon(icon)
        self.escolher_produto = []
        self.btn_Lupa_Prod.clicked.connect(self.abrir_tela_escolher_produto)

        icone(self, "menu_cadastro.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Estoque)

        self.definir_bloqueados()

        self.tela_prod_incluir = []
        self.tela_prod_alterar = []
        self.tela_movimentacao = []
        self.tela_compras = []
        self.tela_estrutura = []
        self.tela_preco_venda = []
        self.tela_ppc_prod = []
        self.tela_prod_ficha = []

        validador_inteiro(self.line_Codigo, 123456)

        self.line_Codigo.editingFinished.connect(self.verifica_line_codigo_manual)

        editar_botao(self.btn_Novo, "novo.png", 'Novo', 35)
        self.btn_Novo.clicked.connect(self.abrir_tela_novo)

        editar_botao(self.btn_Alterar, "editar.png", 'Alterar', 35)
        self.btn_Alterar.clicked.connect(self.abrir_tela_alterar)

        editar_botao(self.btn_Excluir, "excluir.png", 'Excluir', 35)
        self.btn_Excluir.clicked.connect(self.excluir_produto)

        self.btn_Mov.clicked.connect(self.abrir_movimentacao)
        self.btn_Compras.clicked.connect(self.abrir_compras)
        self.btn_Estrutura.clicked.connect(self.abrir_estrutura)
        self.btn_Preco.clicked.connect(self.abrir_preco_venda)
        self.btn_Pcp.clicked.connect(self.abrir_pcp_produto)
        self.btn_Imprimir.clicked.connect(self.abrir_prod_ficha)

        self.processando = False

        self.definir_line_bloqueados()

        self.btn_Alterar.setHidden(True)
        self.btn_Excluir.setHidden(True)
        self.btn_Imprimir.setHidden(True)
        self.btn_Mov.setHidden(True)
        self.btn_Compras.setHidden(True)
        self.btn_Estrutura.setHidden(True)
        self.btn_Preco.setHidden(True)
        self.btn_Pcp.setHidden(True)

        self.line_Codigo.setFocus()

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

    def definir_bloqueados(self):
        try:
            self.date_Emissao.setReadOnly(True)
            self.line_Barras.setReadOnly(True)
            self.line_Descricao.setReadOnly(True)
            self.line_DescrCompl.setReadOnly(True)
            self.line_Referencia.setReadOnly(True)
            self.line_Embalagem.setReadOnly(True)
            self.line_kg_mt.setReadOnly(True)
            self.line_Custo_Unit.setReadOnly(True)
            self.line_Local.setReadOnly(True)
            self.line_Qtde_Mini.setReadOnly(True)
            self.plain_Obs.setReadOnly(True)

            self.line_NCM.setReadOnly(True)
            self.line_UM.setReadOnly(True)
            self.line_Conjunto.setReadOnly(True)
            self.line_Tipo.setReadOnly(True)
            self.line_Projeto.setReadOnly(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def abrir_tela_escolher_produto(self):
        cod_prod = self.line_Codigo.text()
        from menu_cadastros.prod_pesquisar import TelaProdutoPesquisar

        self.escolher_produto = TelaProdutoPesquisar(cod_prod, True)
        self.escolher_produto.produto_escolhido.connect(self.atualizar_produto_entry)
        self.escolher_produto.show()

    def atualizar_produto_entry(self, produto):
        self.line_Codigo.setText(produto)
        self.verifica_line_codigo_manual()

    def abrir_movimentacao(self):
        try:
            codigo = self.line_Codigo.text()

            from menu_cadastros.prod_mov import TelaProdutoMovimentacao
            self.tela_movimentacao = TelaProdutoMovimentacao(codigo)
            self.tela_movimentacao.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def abrir_compras(self):
        try:
            codigo = self.line_Codigo.text()

            from menu_cadastros.prod_compras import TelaProdutoCompras
            self.tela_compras = TelaProdutoCompras(codigo)
            self.tela_compras.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def abrir_estrutura(self):
        try:
            codigo = self.line_Codigo.text()

            from menu_estrutura.estrut_incluir import TelaEstruturaIncluir
            self.tela_estrutura = TelaEstruturaIncluir(codigo)
            self.tela_estrutura.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def abrir_preco_venda(self):
        try:
            codigo = self.line_Codigo.text()

            from menu_estrutura.estrut_custo import TelaCusto
            self.tela_preco_venda = TelaCusto(codigo)
            self.tela_preco_venda.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def abrir_pcp_produto(self):
        try:
            codigo = self.line_Codigo.text()

            from menu_pcp.pcp_produto import TelaPcpProduto
            self.tela_ppc_prod = TelaPcpProduto(codigo)
            self.tela_ppc_prod.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def abrir_prod_ficha(self):
        try:
            from menu_cadastros.prod_fichas import TelaFichasProdutos
            self.tela_prod_ficha = TelaFichasProdutos()
            self.tela_prod_ficha.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def abrir_tela_novo(self):
        try:
            from menu_cadastros.pro_incluir import TelaProdutoIncluir
            self.tela_prod_incluir = TelaProdutoIncluir()
            self.tela_prod_incluir.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def abrir_tela_alterar(self):
        try:
            codigo = self.line_Codigo.text()
            emissao = self.date_Emissao.text()
            barra = self.line_Barras.text()
            descr = self.line_Descricao.text()
            compl = self.line_DescrCompl.text()
            ref = self.line_Referencia.text()
            um = self.line_UM.text()
            embalagem = self.line_Embalagem.text()
            kg_mt = self.line_kg_mt.text()
            custo = self.line_Custo_Unit.text()
            local = self.line_Local.text()
            conjunto = self.line_Conjunto.text()
            tipo = self.line_Tipo.text()
            projeto = self.line_Projeto.text()
            qtde_mini = self.line_Qtde_Mini.text()
            ncm = self.line_NCM.text()
            obs = self.plain_Obs.toPlainText()

            dados = (codigo, emissao, barra, descr, compl, ref, um, embalagem, kg_mt, custo, local,
                     conjunto, tipo, projeto, qtde_mini, ncm, obs)

            from menu_cadastros.prod_alterar import TelaProdutoAlterar
            self.tela_prod_alterar = TelaProdutoAlterar(dados)
            self.tela_prod_alterar.alteracao.connect(self.atualizar_dados_produto)
            self.tela_prod_alterar.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_produto(self):
        try:
            codigo = self.line_Codigo.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT COUNT(*) "
                           f"FROM movimentacao AS mov "
                           f"INNER JOIN produto AS prod ON mov.produto = prod.id "
                           f"WHERE prod.codigo = '{codigo}';")
            quantidade_mov = cursor.fetchone()[0]

            if quantidade_mov > 0:
                self.mensagem_alerta(f"O código {codigo} tem movimentos no estoque e não pode ser excluído!")
            else:
                self.verifica_onde_usa_excluir()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_onde_usa_excluir(self):
        try:
            codigo = self.line_Codigo.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT mat.id, mat.mestre, prod.codigo, prod.descricao "
                           f"from materiaprima as mat "
                           f"INNER JOIN produto AS prod ON mat.mestre = prod.id "
                           f"where mat.codigo = '{codigo}';")
            tabela_onde_usa = cursor.fetchall()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao "
                           f"from estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where prod.codigo = '{codigo}';")
            tabela_onde_usa_v2 = cursor.fetchall()

            if tabela_onde_usa:
                cod_mestre = tabela_onde_usa[0][1]
                descr_mestre = tabela_onde_usa[0][2]

                self.mensagem_alerta(f"O código {codigo} está sendo usado na estrutura do produto "
                                     f"{cod_mestre} - {descr_mestre} e não pode ser excluído!")

            elif tabela_onde_usa_v2:
                cod_mestre = tabela_onde_usa[0][1]
                descr_mestre = tabela_onde_usa[0][2]

                self.mensagem_alerta(f"O código {codigo} está sendo usado na estrutura do produto "
                                     f"{cod_mestre} - {descr_mestre} e não pode ser excluído!")

            else:
                self.verifica_estrutura_excluir()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_estrutura_excluir(self):
        try:
            codigo = self.line_Codigo.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT mat.id, mat.mestre, prod.codigo, prod.descricao "
                           f"from materiaprima as mat "
                           f"INNER JOIN produto AS prod ON mat.mestre = prod.id "
                           f"where prod.codigo = '{codigo}';")
            tabela_estrutura = cursor.fetchall()

            codigo_produto = self.line_Codigo.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {codigo_produto};")
            select_prod = cursor.fetchall()
            id_pai, cod, id_versao = select_prod[0]

            if tabela_estrutura:
                self.mensagem_alerta(f"O código {codigo} tem estrutura cadastrada e não pode ser excluído!")
            elif id_versao:
                self.mensagem_alerta(f"O código {codigo} tem estrutura cadastrada e não pode ser excluído!")
            else:
                self.verifica_compras_excluir()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_compras_excluir(self):
        try:
            codigo = self.line_Codigo.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT COALESCE(prodreq.mestre, ''), prodreq.quantidade "
                           f"FROM produtoordemsolicitacao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"WHERE prod.codigo = '{codigo}';")
            dados_sol = cursor.fetchall()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prodreq.quantidade, prodreq.numero, "
                           f"prodreq.destino, prodreq.id_prod_sol "
                           f"FROM produtoordemrequisicao as prodreq "
                           f"INNER JOIN produto as prod ON prodreq.produto = prod.ID "
                           f"where prod.codigo = '{codigo}';")
            dados_req = cursor.fetchall()

            cursor = conecta.cursor()
            cursor.execute(
                f"SELECT oc.data, oc.numero, prodoc.quantidade, prodoc.produzido, prodoc.dataentrega "
                f"FROM ordemcompra as oc "
                f"INNER JOIN produtoordemcompra as prodoc ON oc.id = prodoc.mestre "
                f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                f"where oc.entradasaida = 'E' "
                f"and prod.codigo = '{codigo}';")
            dados_oc = cursor.fetchall()

            if dados_sol:
                self.mensagem_alerta(f"O Código {codigo} tem Solicitações de Compra e "
                                     f"não pode ser excluído!")
            elif dados_req:
                self.mensagem_alerta(f"O Código {codigo} tem Requisições de Compra e "
                                     f"não pode ser excluído!")
            elif dados_oc:
                self.mensagem_alerta(f"O Código {codigo} tem Ordens de Compra e "
                                     f"não pode ser excluído!")
            else:
                self.verifica_vendas_excluir()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_vendas_excluir(self):
        try:
            codigo = self.line_Codigo.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ped.emissao, ped.id, cli.razao, prodint.qtde, "
                           f"prodint.data_previsao "
                           f"FROM PRODUTOPEDIDOINTERNO as prodint "
                           f"INNER JOIN produto as prod ON prodint.id_produto = prod.id "
                           f"INNER JOIN pedidointerno as ped ON prodint.id_pedidointerno = ped.id "
                           f"INNER JOIN clientes as cli ON ped.id_cliente = cli.id "
                           f"where prod.codigo = {codigo};")
            dados_pi = cursor.fetchall()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT oc.data, oc.numero, cli.razao, prodoc.quantidade, prodoc.dataentrega, "
                           f"COALESCE(prodoc.id_pedido, '') as pedi "
                           f"FROM PRODUTOORDEMCOMPRA as prodoc "
                           f"INNER JOIN produto as prod ON prodoc.produto = prod.id "
                           f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                           f"INNER JOIN clientes as cli ON oc.cliente = cli.id "
                           f"LEFT JOIN pedidointerno as ped ON prodoc.id_pedido = ped.id "
                           f"where prodoc.quantidade > prodoc.produzido "
                           f"and oc.entradasaida = 'S' "
                           f"and prod.codigo = {codigo};")
            dados_ov = cursor.fetchall()

            if dados_pi:
                self.mensagem_alerta(f"O Código {codigo} tem Pedido Interno de Venda e "
                                     f"não pode ser excluído!")
            elif dados_ov:
                self.mensagem_alerta(f"O Código {codigo} tem Ordens de Venda e "
                                     f"não pode ser excluído!")
            else:
                self.verifica_ordem_servico_excluir()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_ordem_servico_excluir(self):
        try:
            codigo = self.line_Codigo.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT os.id, prod.codigo "
                           f"FROM servico as os "
                           f"INNER JOIN produto as prod ON os.produto = prod.id "
                           f"where prod.codigo = {codigo};")
            dados_produto = cursor.fetchall()

            if dados_produto:
                self.mensagem_alerta(f"O Código {codigo} tem Ordem de Serviço e "
                                     f"não pode ser excluído!")
            else:
                msg = f'Tem certeza que deseja excluir este produto?'
                if self.pergunta_confirmacao(msg):
                    cursor = conecta.cursor()
                    cursor.execute(f"DELETE FROM produto WHERE codigo = '{codigo}';")

                    conecta.commit()

                    self.mensagem_alerta(f"Cadastro do produto {codigo} foi excluído com Sucesso!")
                    self.limpa_dados_produto()
                    self.line_Codigo.clear()
                    self.line_Codigo.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def atualizar_dados_produto(self, alterado):
        try:
            if alterado:
                self.verifica_line_codigo_manual()

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

                if int(codigo_produto) == 0:
                    self.mensagem_alerta('O campo "Código" não pode ser "0"')
                    self.limpa_dados_produto()
                    self.line_Codigo.clear()
                    self.line_Codigo.setFocus()
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
                self.line_Codigo.clear()
                self.line_Codigo.setFocus()
            else:
                self.limpa_dados_produto()
                self.table_Estoque.clearContents()
                self.lanca_dados_produto_manual()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_produto_manual(self):
        try:
            codigo_produto = self.line_Codigo.text()
            cur = conecta.cursor()
            cur.execute(f"SELECT prod.id, prod.codbarras, prod.descricao, COALESCE(prod.descricaocomplementar, ''), "
                        f"COALESCE(prod.obs, ''), prod.unidade, COALESCE(prod.localizacao, ''), prod.ncm, "
                        f"prod.quantidade, prod.embalagem, COALESCE(prod.kilosmetro, ''), conj.conjunto, "
                        f"COALESCE(tip.tipomaterial, ''), prod.DATA_CRIACAO, COALESCE(prod.embalagem, ''), "
                        f"prod.custounitario, prod.quantidademin, proj.projeto, conj.id, prod.obs2 "
                        f"FROM produto as prod "
                        f"LEFT JOIN conjuntos conj ON prod.conjunto = conj.id "
                        f"LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id "
                        f"LEFT JOIN projeto proj ON prod.projeto = proj.id "
                        f"where prod.codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            if detalhes_produto:
                id_prod, barras, descr, compl, ref, um, local, ncm, saldo, embal, kg_mt, conjunto, tipo, \
                data, embalagem, custo, minima, projeto, id_conj, obs = detalhes_produto[0]

                barras_sem_espacos = barras.strip()

                custo_unit = float_para_virgula(custo)

                if data:
                    self.date_Emissao.setDate(data)

                self.line_Barras.setText(barras_sem_espacos)
                self.line_Descricao.setText(descr)
                self.line_DescrCompl.setText(compl)
                self.line_Referencia.setText(ref)
                self.line_UM.setText(um)
                self.line_Embalagem.setText(embalagem)
                self.line_kg_mt.setText(kg_mt)
                self.line_Custo_Unit.setText(custo_unit)
                self.line_Saldo_Total.setText(str(saldo))
                self.line_Local.setText(local)
                self.line_Conjunto.setText(conjunto)
                self.line_Tipo.setText(tipo)
                self.line_Projeto.setText(projeto)
                self.plain_Obs.setPlainText(obs)

                if minima:
                    min_str = str(minima)
                else:
                    min_str = "0"

                self.line_Qtde_Mini.setText(min_str)

                if not ncm:
                    self.line_NCM.setText("")
                    self.line_NCM.setStyleSheet("QLineEdit { background-color: yellow; }")
                else:
                    self.line_NCM.setText(ncm)
                    self.line_NCM.setStyleSheet("QLineEdit { background-color: white; }")

                self.manipula_dados_tabela_estoque(id_prod)
                self.btn_Alterar.setHidden(False)
                self.btn_Excluir.setHidden(False)
                self.btn_Imprimir.setHidden(False)
                self.btn_Mov.setHidden(False)
                self.btn_Compras.setHidden(False)
                self.btn_Preco.setHidden(False)
                self.btn_Pcp.setHidden(False)

                if id_conj == 10:
                    self.btn_Estrutura.setHidden(False)
                else:
                    self.btn_Estrutura.setHidden(True)
            else:
                self.mensagem_alerta("Este cadastro de produto não existe!")
                self.line_Codigo.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_produto(self):
        try:
            data_hoje = date.today()
            self.date_Emissao.setDate(data_hoje)

            self.line_Referencia.clear()
            self.line_Descricao.clear()
            self.line_DescrCompl.clear()
            self.line_Embalagem.clear()
            self.line_NCM.clear()
            self.line_kg_mt.clear()
            self.line_Qtde_Mini.clear()
            self.line_Custo_Unit.clear()
            self.line_Local.clear()
            self.line_Barras.clear()
            self.plain_Obs.clear()

            self.line_UM.clear()
            self.line_Conjunto.clear()
            self.line_Tipo.clear()
            self.line_Projeto.clear()

            self.btn_Alterar.setHidden(True)
            self.btn_Excluir.setHidden(True)
            self.btn_Imprimir.setHidden(True)
            self.btn_Mov.setHidden(True)
            self.btn_Compras.setHidden(True)
            self.btn_Estrutura.setHidden(True)
            self.btn_Preco.setHidden(True)
            self.btn_Pcp.setHidden(True)

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


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaProdutoConsultar()
    tela.show()
    qt.exec_()

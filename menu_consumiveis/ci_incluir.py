import sys
from banco_dados.conexao import conecta
from forms.tela_ci_incluir import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.telas import tamanho_aplicacao, icone
from comandos.conversores import valores_para_float
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import inspect
import os
from datetime import date, datetime
import traceback


class TelaCiIncluir(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_consumiveis.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Recomendacao)
        layout_cabec_tab(self.table_Consumo)
        layout_cabec_tab(self.table_Estoque)

        self.line_Obs.setText("PRODUÇÃO INTERNA")

        self.table_Recomendacao.viewport().installEventFilter(self)

        self.btn_Consulta_OC.clicked.connect(self.consultar_oc)
        self.line_Num_OC.editingFinished.connect(self.verifica_num_oc)

        self.btn_Consulta_Func.clicked.connect(self.consultar_funcionario)

        self.btn_ExcluirTudo_Rec.clicked.connect(self.excluir_tudo_rec)
        self.btn_ExcluirItem_Rec.clicked.connect(self.excluir_item_rec)
        self.btn_Adicionar_Todos_Rec.clicked.connect(self.lanca_tudo_rec)

        self.btn_ExcluirTudo_Con.clicked.connect(self.excluir_tudo_con)
        self.btn_ExcluirItem_Con.clicked.connect(self.excluir_item_con)

        self.line_Codigo.editingFinished.connect(self.verifica_line_codigo_manual)
        self.line_Qtde.editingFinished.connect(self.verifica_line_qtde_manual)

        self.btn_Adicionar.clicked.connect(self.verifica_line_qtde_manual)

        self.btn_Buscar.clicked.connect(self.procura_produtos)

        self.btn_Limpar.clicked.connect(self.limpa_tudo)

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)

        self.processando = False

        self.definir_line_ano_consumo()
        self.definir_combo_funcionario()
        self.definir_combo_consulta_funcionario()
        self.definir_combo_localestoque()
        self.combo_Funcionario.setFocus()

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

    def definir_line_ano_consumo(self):
        try:
            data_hoje = date.today()
            ano_atual = data_hoje.strftime("%Y")
            ano_menos_dois = str(int(ano_atual) - 1)
            self.date_Emissao.setDate(data_hoje)
            self.line_Ano.setText(ano_menos_dois)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_combo_funcionario(self):
        try:
            tabela = []

            self.combo_Funcionario.clear()
            tabela.append("")

            cur = conecta.cursor()
            cur.execute(f"SELECT id, funcionario FROM funcionarios where ativo = 'S' order by funcionario;")
            detalhes_func = cur.fetchall()

            for dadus in detalhes_func:
                ides, func = dadus
                tabela.append(func)

            self.combo_Funcionario.addItems(tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_combo_consulta_funcionario(self):
        try:
            tabela = []

            self.combo_Func_Consulta.clear()
            tabela.append("")

            cur = conecta.cursor()
            cur.execute(f"SELECT id, funcionario FROM funcionarios where ativo = 'S' order by funcionario;")
            detalhes_func = cur.fetchall()

            for dadus in detalhes_func:
                ides, func = dadus
                tabela.append(func)

            self.combo_Func_Consulta.addItems(tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_combo_localestoque(self):
        try:
            tabela = []
            cur = conecta.cursor()
            cur.execute(f"SELECT id, nome FROM LOCALESTOQUE order by nome;")
            detalhes = cur.fetchall()

            for dadus in detalhes:
                ides, local = dadus
                tabela.append(local)

            self.combo_Local_Estoque.addItems(tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_num_oc(self):
        if not self.processando:
            try:
                self.processando = True

                self.excluir_tudo_rec()

                numero_oc = self.line_Num_OC.text()

                if numero_oc:
                    self.consultar_oc()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def consultar_oc(self):
        try:
            self.excluir_tudo_rec()

            tabela = []
            numero_oc = self.line_Num_OC.text()

            if numero_oc:
                cur = conecta.cursor()
                cur.execute(f"SELECT prod.id, oc.data, prod.codigo, prod.descricao, prod.unidade, "
                            f"prodoc.quantidade, prod.quantidade, COALESCE(prod.localizacao, '') as loc "
                            f"FROM PRODUTOORDEMCOMPRA as prodoc "
                            f"INNER JOIN produto prod ON prodoc.produto = prod.id "
                            f"INNER JOIN ORDEMCOMPRA oc ON prodoc.mestre = oc.id "
                            f"INNER JOIN LOCALESTOQUE loc ON oc.localestoque = loc.id "
                            f"where prodoc.numero = {numero_oc} and oc.entradasaida = 'E' "
                            f"order by prod.codigo;")
                detalhes = cur.fetchall()
                if detalhes:
                    for dadus in detalhes:
                        id_prod, data, codigo, descr, um, qtde, saldo, local = dadus

                        data1 = f'{data.day}/{data.month}/{data.year}'

                        selecao = (data1, codigo, descr, um, qtde, saldo, local)
                        tabela.append(selecao)
                else:
                    self.mensagem_alerta("Não foi encontrado itens dessa Ordem de Compra!")
                    self.excluir_tudo_rec()

            if tabela:
                lanca_tabela(self.table_Recomendacao, tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consultar_funcionario(self):
        try:
            self.excluir_tudo_rec()

            tabela = []

            func = self.combo_Func_Consulta.currentText()

            cur = conecta.cursor()
            cur.execute(f"SELECT id, funcionario FROM funcionarios where funcionario = '{func}';")
            detalhes_func = cur.fetchall()
            id_func, nome_func = detalhes_func[0]

            ano = self.line_Ano.text()

            if func:
                if ano:
                    cur = conecta.cursor()
                    cur.execute(f"SELECT prod.id, mov.data, prod.codigo, prod.descricao, prod.unidade, "
                                f"mov.quantidade, prod.quantidade, COALESCE(prod.localizacao, '') as loc "
                                f"FROM movimentacao as mov "
                                f"INNER JOIN produto prod ON mov.produto = prod.id "
                                f"INNER JOIN LOCALESTOQUE loc ON mov.localestoque = loc.id "
                                f"where mov.tipo = 220 "
                                f"and mov.funcionario = {id_func} "
                                f"and mov.data > '{ano}-01-01';")
                    detalhes = cur.fetchall()

                    if detalhes:
                        for dadus in detalhes:
                            id_prod, data, codigo, descr, um, qtde, saldo, local = dadus

                            data1 = f'{data.day}/{data.month}/{data.year}'

                            selecao = (data1, codigo, descr, um, qtde, saldo, local)
                            tabela.append(selecao)

            if tabela:
                lanca_tabela(self.table_Recomendacao, tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def procura_produtos(self):
        try:
            notifica = 0

            tabela = []

            descricao1 = self.line_Descricao1_Busca.text().upper()
            descricao2 = self.line_Descricao2_Busca.text().upper()
            estoque = self.check_Estoque_Busca.isChecked()
            movimentacao = self.check_Mov_Busca.isChecked()

            if descricao1 and descricao2 and estoque and movimentacao:
                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT DISTINCT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                    f"prod.unidade, "
                    f"COALESCE(prod.localizacao, '') as loc, COALESCE(prod.quantidade, '') as qti "
                    f"FROM produto as prod "
                    f"INNER JOIN movimentacao as mov ON prod.id = mov.produto "
                    f"WHERE (prod.descricao LIKE '%{descricao1}%' OR "
                    f"prod.descricaocomplementar LIKE '%{descricao1}%' OR "
                    f"prod.obs LIKE '%{descricao1}%') "
                    f"AND (prod.descricao LIKE '%{descricao2}%' OR "
                    f"prod.descricaocomplementar LIKE '%{descricao2}%' OR "
                    f"prod.obs LIKE '%{descricao2}%') AND prod.quantidade > 0 "
                    f"ORDER BY prod.descricao;")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif descricao1 and estoque and movimentacao:
                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT DISTINCT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                    f"prod.unidade, "
                    f"COALESCE(prod.localizacao, '') as loc, COALESCE(prod.quantidade, '') as qti "
                    f"FROM produto as prod "
                    f"INNER JOIN movimentacao as mov ON prod.id = mov.produto "
                    f"WHERE (prod.descricao LIKE '%{descricao1}%' OR "
                    f"prod.descricaocomplementar LIKE '%{descricao1}%' OR "
                    f"prod.obs LIKE '%{descricao1}%') "
                    f"AND prod.quantidade > 0 "
                    f"ORDER BY prod.descricao;")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif descricao2 and estoque and movimentacao:
                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT DISTINCT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                    f"prod.unidade, "
                    f"COALESCE(prod.localizacao, '') as loc, COALESCE(prod.quantidade, '') as qti "
                    f"FROM produto as prod "
                    f"INNER JOIN movimentacao as mov ON prod.id = mov.produto "
                    f"WHERE (prod.descricao LIKE '%{descricao2}%' OR "
                    f"prod.descricaocomplementar LIKE '%{descricao2}%' OR "
                    f"prod.obs LIKE '%{descricao2}%') AND prod.quantidade > 0 "
                    f"ORDER BY prod.descricao;")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif descricao1 and descricao2 and movimentacao:
                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT DISTINCT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                    f"prod.unidade, "
                    f"COALESCE(prod.localizacao, '') as loc, COALESCE(prod.quantidade, '') as qti "
                    f"FROM produto as prod "
                    f"INNER JOIN movimentacao as mov ON prod.id = mov.produto "
                    f"WHERE (prod.descricao LIKE '%{descricao1}%' OR "
                    f"prod.descricaocomplementar LIKE '%{descricao1}%' OR "
                    f"prod.obs LIKE '%{descricao1}%') "
                    f"AND (prod.descricao LIKE '%{descricao2}%' OR "
                    f"prod.descricaocomplementar LIKE '%{descricao2}%' OR "
                    f"prod.obs LIKE '%{descricao2}%') "
                    f"ORDER BY prod.descricao;")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif descricao1 and descricao2 and estoque:
                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT DISTINCT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                    f"prod.unidade, "
                    f"COALESCE(prod.localizacao, '') as loc, COALESCE(prod.quantidade, '') as qti "
                    f"FROM produto as prod "
                    f"WHERE (prod.descricao LIKE '%{descricao1}%' OR "
                    f"prod.descricaocomplementar LIKE '%{descricao1}%' OR "
                    f"prod.obs LIKE '%{descricao1}%') "
                    f"AND (prod.descricao LIKE '%{descricao2}%' OR "
                    f"prod.descricaocomplementar LIKE '%{descricao2}%' OR "
                    f"prod.obs LIKE '%{descricao2}%') AND prod.quantidade > 0 "
                    f"ORDER BY prod.descricao;")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif descricao1 and descricao2:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, codigo, descricao, COALESCE(obs, '') as obs, unidade, "
                               f"COALESCE(localizacao, '') as loc, COALESCE(quantidade, '') as qti "
                               f"FROM produto "
                               f"WHERE (descricao LIKE '%{descricao1}%' OR "
                               f"descricaocomplementar LIKE '%{descricao1}%' OR "
                               f"obs LIKE '%{descricao1}%') "
                               f"AND (descricao LIKE '%{descricao2}%' OR "
                               f"descricaocomplementar LIKE '%{descricao2}%' OR "
                               f"obs LIKE '%{descricao2}%') "
                               f"order by descricao;")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif descricao1 and estoque:
                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT id, codigo, descricao, COALESCE(obs, '') as obs, unidade, "
                    f"COALESCE(localizacao, '') as loc, COALESCE(quantidade, '') as qti "
                    f"FROM produto "
                    f"WHERE (descricao LIKE '%{descricao1}%' OR "
                    f"descricaocomplementar LIKE '%{descricao1}%' OR "
                    f"obs LIKE '%{descricao1}%') "
                    f"AND quantidade > 0 "
                    f"ORDER BY descricao;")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif descricao2 and estoque:
                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                    f"prod.unidade, "
                    f"COALESCE(prod.localizacao, '') as loc, COALESCE(prod.quantidade, '') as qti "
                    f"FROM produto as prod "
                    f"INNER JOIN movimentacao as mov ON prod.id = mov.produto "
                    f"WHERE (prod.descricao LIKE '%{descricao2}%' OR "
                    f"prod.descricaocomplementar LIKE '%{descricao2}%' OR "
                    f"prod.obs LIKE '%{descricao2}%') AND prod.quantidade > 0 "
                    f"ORDER BY prod.descricao;")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif descricao1 and movimentacao:
                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT DISTINCT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                    f"prod.unidade, "
                    f"COALESCE(prod.localizacao, '') as loc, COALESCE(prod.quantidade, '') as qti "
                    f"FROM produto as prod "
                    f"INNER JOIN movimentacao as mov ON prod.id = mov.produto "
                    f"WHERE (prod.descricao LIKE '%{descricao1}%' OR "
                    f"prod.descricaocomplementar LIKE '%{descricao1}%' OR "
                    f"prod.obs LIKE '%{descricao1}%') "
                    f"ORDER BY prod.descricao;")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif descricao2 and movimentacao:
                cursor = conecta.cursor()
                cursor.execute(
                    f"SELECT DISTINCT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                    f"prod.unidade, "
                    f"COALESCE(prod.localizacao, '') as loc, COALESCE(prod.quantidade, '') as qti "
                    f"FROM produto as prod "
                    f"INNER JOIN movimentacao as mov ON prod.id = mov.produto "
                    f"WHERE (prod.descricao LIKE '%{descricao2}%' OR "
                    f"prod.descricaocomplementar LIKE '%{descricao2}%' OR "
                    f"prod.obs LIKE '%{descricao2}%') "
                    f"ORDER BY prod.descricao;")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif estoque and movimentacao:
                self.mensagem_alerta("Preencha no mínimo uma Descrição, descr. Complementar ou Referência do produto!")

            elif descricao1:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, codigo, descricao, COALESCE(obs, '') as obs, unidade, "
                               f"COALESCE(localizacao, '') as loc, COALESCE(quantidade, '') as qti "
                               f"FROM produto "
                               f"WHERE descricao LIKE '%{descricao1}%' OR "
                               f"descricaocomplementar LIKE '%{descricao1}%' OR "
                               f"obs LIKE '%{descricao1}%';")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif descricao2:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, codigo, descricao, COALESCE(obs, '') as obs, unidade, "
                               f"COALESCE(localizacao, '') as loc, COALESCE(quantidade, '') as qti "
                               f"FROM produto "
                               f"WHERE descricao LIKE '%{descricao2}%' OR "
                               f"descricaocomplementar LIKE '%{descricao2}%' OR "
                               f"obs LIKE '%{descricao2}%';")
                detalhes_produto = cursor.fetchall()
                for tudo in detalhes_produto:
                    id_prod, cod, descr, ref, um, local, saldo = tudo

                    dados = ("", cod, descr, um, "1", saldo, local)
                    tabela.append(dados)

                notifica += 1

            elif estoque:
                self.mensagem_alerta("Preencha no mínimo uma Descrição, descr. Complementar ou Referência do produto!")

            elif movimentacao:
                self.mensagem_alerta("Preencha no mínimo uma Descrição, descr. Complementar ou Referência do produto!")

            else:
                self.mensagem_alerta("Preencha no mínimo uma Descrição, descr. Complementar ou Referência do produto!")

            if notifica:
                if tabela:
                    lanca_tabela(self.table_Recomendacao, tabela)
                else:
                    self.mensagem_alerta("Não foi encontrado nenhum registro com essas condições!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def eventFilter(self, source, event):
        try:
            if (event.type() == QtCore.QEvent.MouseButtonDblClick and
                    event.buttons() == QtCore.Qt.LeftButton and
                    source is self.table_Recomendacao.viewport()):

                funcionario = self.combo_Funcionario.currentText()
                local_est = self.combo_Local_Estoque.currentText()

                if funcionario and local_est:
                    item = self.table_Recomendacao.currentItem()

                    extrai_recomendados = extrair_tabela(self.table_Recomendacao)
                    item_selecionado = extrai_recomendados[item.row()]

                    data_rec, c_rec, d_rec, um_rec, qtde_rec, saldo, local = item_selecionado

                    qtde_rec_float = valores_para_float(qtde_rec)
                    if qtde_rec_float > 0:
                        qtde_final = qtde_rec_float
                    else:
                        qtde_final = 1

                    extrai_consumo = extrair_tabela(self.table_Consumo)

                    didis = [c_rec, d_rec, um_rec, qtde_final, funcionario, local_est]
                    extrai_consumo.append(didis)

                    if extrai_consumo:
                        lanca_tabela(self.table_Consumo, extrai_consumo)
                else:
                    self.mensagem_alerta("Defina o Funcionário e o Local de Estoque!")

            return super(QMainWindow, self).eventFilter(source, event)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_tudo_rec(self):
        try:
            extrai_recomendados = extrair_tabela(self.table_Recomendacao)
            if extrai_recomendados:
                self.table_Recomendacao.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_tudo_con(self):
        try:
            extrai_recomendados = extrair_tabela(self.table_Consumo)
            if extrai_recomendados:
                self.table_Consumo.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_tudo_est(self):
        try:
            extrai_recomendados = extrair_tabela(self.table_Estoque)
            if extrai_recomendados:
                self.table_Estoque.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_item_rec(self):
        try:
            extrai_recomendados = extrair_tabela(self.table_Recomendacao)
            if not extrai_recomendados:
                self.mensagem_alerta(f'A tabela "Lista de Recomendações" está vazia!')
            else:
                linha_selecao = self.table_Recomendacao.currentRow()
                if linha_selecao >= 0:
                    self.table_Recomendacao.removeRow(linha_selecao)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_item_con(self):
        try:
            extrai_recomendados = extrair_tabela(self.table_Consumo)
            if not extrai_recomendados:
                self.mensagem_alerta(f'A tabela "Lista de Consumo" está vazia!')
            else:
                linha_selecao = self.table_Consumo.currentRow()
                if linha_selecao >= 0:
                    self.table_Consumo.removeRow(linha_selecao)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_manual(self):
        self.line_Codigo.clear()
        self.line_Referencia.clear()
        self.line_Descricao.clear()
        self.line_Local.clear()
        self.line_UM.clear()
        self.line_Qtde.clear()
        self.line_Saldo_Total.clear()
        self.excluir_tudo_est()

    def limpa_consulta_oc_func(self):
        try:
            self.line_Num_OC.clear()
            self.definir_line_ano_consumo()
            self.definir_combo_consulta_funcionario()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_consulta_prod(self):
        try:
            self.line_Descricao1_Busca.clear()
            self.line_Descricao2_Busca.clear()
            self.check_Estoque_Busca.isChecked()
            self.check_Mov_Busca.isChecked()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tudo(self):
        try:
            self.limpa_manual()
            self.definir_combo_funcionario()
            self.definir_combo_localestoque()
            self.limpa_consulta_oc_func()
            self.limpa_consulta_prod()

            self.excluir_tudo_con()
            self.excluir_tudo_rec()
            self.line_Obs.clear()

            self.combo_Funcionario.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_tudo_rec(self):
        try:
            funcionario = self.combo_Funcionario.currentText()
            local_est = self.combo_Local_Estoque.currentText()

            if funcionario and local_est:
                extrai_consumo = extrair_tabela(self.table_Consumo)
                extrai_recomendados = extrair_tabela(self.table_Recomendacao)

                if not extrai_recomendados:
                    self.mensagem_alerta(f'A tabela "Lista de Recomendações" está vazia!')
                else:
                    for i in extrai_recomendados:
                        data_rec, cod_rec, desc_rec, um_rec, qtde_rec, saldo, local = i

                        qtde_rec_float = valores_para_float(qtde_rec)
                        if qtde_rec_float > 0:
                            qtde_final = qtde_rec_float
                        else:
                            qtde_final = 1

                        didis = [cod_rec, desc_rec, um_rec, qtde_final, funcionario, local_est]
                        extrai_consumo.append(didis)

                    if extrai_consumo:
                        lanca_tabela(self.table_Consumo, extrai_consumo)
            else:
                self.mensagem_alerta("Defina o Funcionário e o Local de Estoque!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_codigo_manual(self):
        if not self.processando:
            try:
                self.processando = True

                codigo_produto = self.line_Codigo.text()
                funcionario = self.combo_Funcionario.currentText()
                local_est = self.combo_Local_Estoque.currentText()

                if codigo_produto:
                    if int(codigo_produto) == 0:
                        self.mensagem_alerta('O campo "Código" não pode ser "0"')
                        self.limpa_manual()
                    else:
                        if funcionario and local_est:
                            self.verifica_sql_produto_manual()
                        else:
                            self.mensagem_alerta("Defina o Funcionário e o Local de Estoque!")

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
            cursor.execute(f"SELECT id, descricao, COALESCE(obs, ' ') as obs, unidade, localizacao, quantidade "
                           f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cursor.fetchall()
            if not detalhes_produto:
                self.mensagem_alerta('Este código de produto não existe!')
                self.limpa_manual()
            else:
                ides, descr, ref, um, local, saldo = detalhes_produto[0]
                saldo_float = valores_para_float(saldo)
                if saldo_float > 0:
                    self.lanca_dados_produto_manual()
                else:
                    self.mensagem_alerta('Este produto não possui saldo em estoque!')
                    self.limpa_manual()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_produto_manual(self):
        try:
            codigo_produto = self.line_Codigo.text()

            cur = conecta.cursor()
            cur.execute(f"SELECT id, descricao, COALESCE(obs, '') as obs, unidade, COALESCE(localizacao, '') as local, "
                        f"quantidade, embalagem, kilosmetro FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            ides, descr, ref, um, local, qtde, embalagem, kg_mt = detalhes_produto[0]

            self.line_Descricao.setText(descr)
            self.line_UM.setText(um)
            self.line_Local.setText(local)
            self.line_Qtde.setFocus()

            self.lanca_saldos_local_manual(ides)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_saldos_local_manual(self, id_prod):
        try:
            cur = conecta.cursor()
            cur.execute(f"SELECT loc.nome, sald.saldo FROM SALDO_ESTOQUE as sald "
                        f"INNER JOIN LOCALESTOQUE loc ON sald.local_estoque = loc.id "
                        f"where sald.produto_id = {id_prod} order by loc.nome;")
            detalhes_saldos = cur.fetchall()
            if detalhes_saldos:
                lanca_tabela(self.table_Estoque, detalhes_saldos)

            self.soma_saldos_manual()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def soma_saldos_manual(self):
        try:
            saldo_somado = 0.00
            dados_local_estoque = extrair_tabela(self.table_Estoque)
            if dados_local_estoque:
                for i in dados_local_estoque:
                    local, qtde = i

                    qtde_float = valores_para_float(qtde)

                    saldo_somado += qtde_float

            saldo_somado_2casas = ("%.3f" % saldo_somado)
            self.line_Saldo_Total.setText(saldo_somado_2casas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_qtde_manual(self):
        if not self.processando:
            try:
                self.processando = True

                qtdezinha = self.line_Qtde.text()

                if qtdezinha:
                    if qtdezinha == "0":
                        self.mensagem_alerta('O campo "Qtde:" não pode ser "0"')
                        self.line_Qtde.clear()
                        self.line_Qtde.setFocus()
                    else:
                        self.item_produto_manual()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def item_produto_manual(self):
        try:
            funcionario = self.combo_Funcionario.currentText()
            local_est = self.combo_Local_Estoque.currentText()

            um = self.line_UM.text()
            cod = self.line_Codigo.text()
            descr = self.line_Descricao.text()

            qtde_manu = self.line_Qtde.text()
            if "," in qtde_manu:
                qtde_manu_com_ponto = qtde_manu.replace(',', '.')
                qtde_float = float(qtde_manu_com_ponto)
            else:
                qtde_float = float(qtde_manu)

            extrai_consumo = extrair_tabela(self.table_Consumo)

            dados = [cod, descr, um, qtde_float, funcionario, local_est]
            extrai_consumo.append(dados)
            lanca_tabela(self.table_Consumo, extrai_consumo)

            self.limpa_manual()
            self.line_Descricao.clear()
            self.line_Referencia.clear()
            self.line_UM.clear()
            self.line_Local.clear()
            self.line_Qtde.clear()
            self.line_Codigo.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_salvamento(self):
        try:
            extrai_consumo = extrair_tabela(self.table_Consumo)

            if not extrai_consumo:
                self.mensagem_alerta(f'A tabela "Lista de Consumo" está vazia!')
            else:
                self.verifica_saldos()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_saldos(self):
        try:
            extrai_consumo = extrair_tabela(self.table_Consumo)

            soma_qtde_dict = {}

            for itens in extrai_consumo:
                codigo, descricao, um, qtde, funcionario, local_estoque = itens

                qtde_float = valores_para_float(qtde)

                chave = (codigo, local_estoque)
                if chave in soma_qtde_dict:
                    soma_qtde_dict[chave] += qtde_float
                else:
                    soma_qtde_dict[chave] = qtde_float

            nova_lista = [(codigo, local_estoque, qtde_somada) for (codigo, local_estoque), qtde_somada in
                          soma_qtde_dict.items()]

            soma_qtde_dict1 = {}
            for i in nova_lista:
                codis, locis, qtids = i

                acumula = 0

                qtids_float = valores_para_float(qtids)

                acumula -= qtids_float

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, codigo, quantidade FROM produto where codigo = '{codis}';")
                dados_produto = cursor.fetchall()
                id_prod, codigo, saldo_total = dados_produto[0]

                saldo_total_float = valores_para_float(saldo_total)

                acumula += saldo_total_float

                if codigo in soma_qtde_dict1:
                    soma_qtde_dict1[codigo] += acumula
                else:
                    soma_qtde_dict1[codigo] = acumula

            nova_lista1 = [(codigo1, qtde_somada1) for (codigo1), qtde_somada1 in
                           soma_qtde_dict1.items()]

            self.define_lista_para_salvar(nova_lista1)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def define_lista_para_salvar(self, lista_acumulada):
        try:
            lista_sem_saldo = []
            lista_com_saldo = []

            datamov = self.date_Emissao.text()
            date_mov = datetime.strptime(datamov, '%d/%m/%Y').date()
            data_emissao = str(date_mov)

            obs = self.line_Obs.text().upper()

            extrai_consumo = extrair_tabela(self.table_Consumo)

            for itens in extrai_consumo:
                codigo, descricao, um, qtde, funcionario, local_estoque = itens

                cur = conecta.cursor()
                cur.execute(f"SELECT id, funcionario FROM funcionarios where funcionario = '{funcionario}';")
                detalhes_func = cur.fetchall()
                id_func, nome_func = detalhes_func[0]

                qtde_float = valores_para_float(qtde)

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, codigo, quantidade FROM produto where codigo = '{codigo}';")
                dados_produto = cursor.fetchall()
                id_prod, codigo, saldo_total = dados_produto[0]

                cur = conecta.cursor()
                cur.execute(f"SELECT id, nome FROM LOCALESTOQUE where nome = '{local_estoque}';")
                detalhes_local = cur.fetchall()

                id_local, nomezinho = detalhes_local[0]

                for acum in lista_acumulada:
                    codis, qtids = acum

                    qtids_float = valores_para_float(qtids)

                    if codis == codigo:
                        if qtids_float < 0:
                            dados_sem = (codigo, descricao, qtde, saldo_total)
                            lista_sem_saldo.append(dados_sem)

                        else:
                            dados_com = (id_prod, obs, qtde_float, data_emissao, codigo, id_func, id_local)
                            lista_com_saldo.append(dados_com)

            self.salvar_consumo_interno(lista_com_saldo, lista_sem_saldo)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_consumo_interno(self, lista_com_saldo, lista_sem_saldo):
        try:
            obss = self.line_Obs.text()
            print(obss)

            if lista_sem_saldo:
                msg = "Não foi possível salvar esta movimentação pois temos produtos sem saldo na lista:\n\n"
                for nao_tem in lista_sem_saldo:
                    codigo, descricao, qtde, saldo_total = nao_tem
                    msg += f" - Cód: {codigo} - {descricao}: Saldo Estoque: {saldo_total}\n\n"

                self.mensagem_alerta(msg)

            else:
                if lista_com_saldo:
                    for tem in lista_com_saldo:
                        id_prod, obs, qtde_float, data_emissao, codigo, id_func, id_local = tem

                        cursor = conecta.cursor()
                        cursor.execute(f"Insert into MOVIMENTACAO (ID, PRODUTO, OBS, TIPO, QUANTIDADE, "
                                       f"DATA, CODIGO, FUNCIONARIO, LOCALESTOQUE) values "
                                       f"(GEN_ID(GEN_MOVIMENTACAO_ID,1), {id_prod}, '{obs}', '220', "
                                       f"{qtde_float}, '{data_emissao}', {codigo}, {id_func}, {id_local});")

                    conecta.commit()
                    print("salvado")

                    self.mensagem_alerta(f'Consumo Interno foi lançada com sucesso!')
                    self.limpa_tudo()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaCiIncluir()
    tela.show()
    qt.exec_()

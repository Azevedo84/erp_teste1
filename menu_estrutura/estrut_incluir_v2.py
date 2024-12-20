import sys
from banco_dados.conexao import conecta
from forms.tela_estrut_incluir_v2 import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.lines import validador_decimal
from comandos.telas import tamanho_aplicacao, icone
from comandos.cores import cor_cinza_claro
from comandos.conversores import valores_para_float, valores_para_virgula
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtGui import QColor
from datetime import date
import inspect
import os
import traceback


class TelaEstruturaIncluirV2(QMainWindow, Ui_MainWindow):
    def __init__(self, produto, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_estrutura.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Estrutura)

        self.definir_line_bloqueados()

        self.processando = False

        if produto:
            self.line_Codigo_Estrut.setText(produto)
            self.line_Codigo_Estrut.setReadOnly(True)
            self.verifica_line_codigo_acabado()

        self.line_Codigo_Estrut.setFocus()

        self.line_Codigo_Estrut.editingFinished.connect(self.verifica_line_codigo_acabado)
        self.line_Codigo_Manu.editingFinished.connect(self.verifica_line_codigo_manual)
        self.line_Qtde_Manu.editingFinished.connect(self.verifica_line_qtde_manual)
        self.line_Medida_Manu.editingFinished.connect(self.verifica_line_medida_manual)
        self.line_Tempo_Mao.editingFinished.connect(self.mascara_tempo_mao_de_obra)

        self.combo_Versao.activated.connect(self.manipula_versao_escolhida)

        self.btn_ExcluirTudo.clicked.connect(self.excluir_tudo_tab_produtos)
        self.btn_ExcluirItem.clicked.connect(self.excluir_produto_tab)
        self.btn_Limpar.clicked.connect(self.limpar)

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)

        self.check_Converte_Manu.stateChanged.connect(self.verifica_check_converter_kilos)

        validator = QtGui.QRegExpValidator(QtCore.QRegExp(r'\d+'), self.line_Codigo_Estrut)
        self.line_Codigo_Estrut.setValidator(validator)

        validator = QtGui.QRegExpValidator(QtCore.QRegExp(r'\d+'), self.line_Codigo_Manu)
        self.line_Codigo_Manu.setValidator(validator)

        validador_decimal(self.line_Qtde_Manu, 9999999.000)
        validador_decimal(self.line_Medida_Manu, 9999999.000)

        self.lista_original = []

        self.widget_MaoObra.setHidden(True)
        self.widget_Terceiros.setHidden(True)
        self.widget_medida_peca.setHidden(True)

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

    def limpa_tabela(self):
        try:
            self.table_Estrutura.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_check_converter_kilos(self):
        try:
            if self.check_Converte_Manu.isChecked():
                self.widget_medida_peca.setHidden(False)
            else:
                self.widget_medida_peca.setHidden(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_line_bloqueados(self):
        try:
            self.line_Descricao_Estrut.setReadOnly(True)
            self.line_Referencia_Estrut.setReadOnly(True)
            self.line_Tipo_Estrut.setReadOnly(True)
            self.line_NCM_Estrut.setReadOnly(True)
            self.line_UM_Estrut.setReadOnly(True)

            self.line_Descricao_Manu.setReadOnly(True)
            self.line_DescrCompl_Manu.setReadOnly(True)
            self.line_Referencia_Manu.setReadOnly(True)
            self.line_UM_Manu.setReadOnly(True)
            self.line_NCM_Manu.setReadOnly(True)
            self.line_Kg_Manu.setReadOnly(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_codigo_acabado(self):
        if not self.processando:
            try:
                self.processando = True

                self.limpa_tudo()

                codigo_produto = self.line_Codigo_Estrut.text()

                if not codigo_produto:
                    self.mensagem_alerta('O campo "Código" não pode estar vazio!')
                    self.limpa_dados_produto_estrutura()
                    self.limpa_tabela()
                elif int(codigo_produto) == 0:
                    self.mensagem_alerta('O campo "Código" não pode ser "0"!')
                    self.limpa_dados_produto_estrutura()
                    self.limpa_tabela()
                else:
                    self.verifica_sql_acabado()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_sql_acabado(self):
        try:
            codigo_produto = self.line_Codigo_Estrut.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade, conjunto, quantidade "
                           f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cursor.fetchall()
            if not detalhes_produto:
                self.mensagem_alerta('Este código de produto não existe!')
                self.limpa_dados_produto_estrutura()
                self.limpa_tabela()
                self.line_Codigo_Estrut.clear()
            else:
                conjunto = detalhes_produto[0][3]
                if conjunto == 10:
                    self.lanca_dados_acabado()
                else:
                    self.mensagem_alerta('Este produto não tem o conjunto classificado como "Produtos Acabados"!')
                    self.limpa_dados_produto_estrutura()
                    self.limpa_tabela()
                    self.line_Codigo_Estrut.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_acabado(self):
        try:
            codigo_produto = self.line_Codigo_Estrut.text()
            cur = conecta.cursor()
            cur.execute(f"SELECT prod.descricao, COALESCE(tip.tipomaterial, '') as tipus, "
                        f"COALESCE(prod.obs, '') as ref, prod.unidade, "
                        f"COALESCE(prod.ncm, '') as ncm, COALESCE(prod.obs2, '') as obs "
                        f"FROM produto as prod "
                        f"LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id "
                        f"where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            descr, tipo, ref, um, ncm, obs = detalhes_produto[0]

            self.line_Descricao_Estrut.setText(descr)
            self.line_Referencia_Estrut.setText(ref)
            self.line_Tipo_Estrut.setText(tipo)
            self.line_NCM_Estrut.setText(ncm)
            self.line_UM_Estrut.setText(um)

            tipo_material = self.line_Tipo_Estrut.text()

            if not tipo_material:
                self.mensagem_alerta('O campo "Tipo de Material" não pode estar vazio!\n\n'
                                     'Entre no cadastro de produtos e defina o Tipo de Material:\n'
                                     'Exemplos: CONJUNTO, USINAGEM, INDUSTRIALIZACAO')
                self.limpa_dados_produto_estrutura()
                self.limpa_tabela()
                self.line_Codigo_Estrut.clear()
            else:
                self.line_Obs.setText(obs)

                if tipo == "INDUSTRIALIZACAO":
                    self.widget_Terceiros.setHidden(False)
                    self.widget_MaoObra.setHidden(True)
                else:
                    self.widget_MaoObra.setHidden(False)
                    self.widget_Terceiros.setHidden(True)

                self.combo_Versao.setFocus()

                self.lanca_descricao_tempo_mao_de_obra(codigo_produto)
                self.lanca_descricao_custo_servico(codigo_produto)

                self.lanca_versoes()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_descricao_tempo_mao_de_obra(self, codigo):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT etapas, tempo FROM produto WHERE codigo = {codigo};")
            dados_produto = cursor.fetchall()
            if dados_produto:
                for i in dados_produto:
                    etapas, tempo = i

                    if etapas:
                        self.line_Descricao_Mao.setText(etapas)
                    if tempo:
                        self.line_Tempo_Mao.setText(str(tempo))

                self.mascara_tempo_mao_de_obra()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_descricao_custo_servico(self, codigo):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT terceirizadoobs, terceirizado FROM produto WHERE codigo = {codigo};")
            dados_produto = cursor.fetchall()
            if dados_produto:
                for i in dados_produto:
                    descr_servico, custo = i

                    if descr_servico:
                        self.line_Descricao_Servico.setText(descr_servico)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_versoes(self):
        try:
            codigo_produto = self.line_Codigo_Estrut.text()

            tabela = []

            self.combo_Versao.clear()
            tabela.append("")

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo, id_versao FROM produto where codigo = {codigo_produto};")
            select_prod = cursor.fetchall()
            id_pai, cod, id_versao = select_prod[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, num_versao, data_versao, obs, data_criacao "
                           f"from estrutura "
                           f"where id_produto = {id_pai} order by data_versao;")
            tabela_versoes = cursor.fetchall()

            if tabela_versoes:
                for i in tabela_versoes:
                    id_estrut, num_versao, data, obs, criacao = i

                    data_versao = data.strftime("%d/%m/%Y")

                    if id_versao == id_estrut:
                        status_txt = "ATIVO"
                    else:
                        status_txt = "INATIVO"

                    msg = f"VERSÃO: {num_versao} - {data_versao} - {status_txt}"
                    tabela.append(msg)
            else:
                data_hoje = date.today()
                data_atual = data_hoje.strftime("%d/%m/%Y")

                msg = f"VERSÃO: 0 - {data_atual} - ATIVO"
                tabela.append(msg)

            self.combo_Versao.addItems(tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_versao_escolhida(self):
        try:
            self.limpa_tabela()
            self.lista_original = []

            codigo_produto = self.line_Codigo_Estrut.text()
            versao = self.combo_Versao.currentText()

            if codigo_produto and versao:
                self.definir_status()
                self.lanca_estrutura()

                self.line_Codigo_Manu.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_estrutura(self):
        try:
            nova_tabela = []
            codigo_produto = self.line_Codigo_Estrut.text()

            versao = self.combo_Versao.currentText()
            versaotete = versao.find(" - ")
            num_versao = versao[8:versaotete]

            status = self.line_Status.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo FROM produto where codigo = {codigo_produto};")
            select_prod = cursor.fetchall()
            id_pai, cod = select_prod[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, num_versao, data_versao, obs, data_criacao "
                           f"from estrutura "
                           f"where id_produto = {id_pai} and num_versao = {num_versao};")
            tabela_versoes = cursor.fetchall()

            if tabela_versoes:
                id_estrutura = tabela_versoes[0][0]

                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                               f"conj.conjunto, prod.unidade, (estprod.quantidade * 1) as qtde, "
                               f"COALESCE(prod.ncm, '') as ncm "
                               f"from estrutura_produto as estprod "
                               f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                               f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                               f"where estprod.id_estrutura = {id_estrutura} "
                               f"order by conj.conjunto DESC, prod.descricao ASC;")
                tabela_estrutura = cursor.fetchall()

                if tabela_estrutura:
                    for i in tabela_estrutura:
                        cod_filho, descr, ref, conjunto, um, qtde, ncm = i

                        qtde_float = float(qtde)
                        tem_consumo = "NÃO"

                        if status == "A":
                            cursor = conecta.cursor()
                            cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.id, "
                                           f"prod.descricao, "
                                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                                           f"ordser.quantidade "
                                           f"from ordemservico as ordser "
                                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                                           f"where ordser.status = 'A' "
                                           f"AND prod.id = {id_pai};")
                            op_abertas = cursor.fetchall()

                            if op_abertas:
                                for ii in op_abertas:
                                    num_op = ii[2]

                                    cursor = conecta.cursor()
                                    cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                                                   f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                                                   f"((SELECT quantidade FROM ordemservico where numero = {num_op}) * "
                                                   f"(estprod.quantidade)) AS Qtde, "
                                                   f"COALESCE(prod.localizacao, ''), prod.quantidade "
                                                   f"FROM estrutura_produto as estprod "
                                                   f"INNER JOIN produto as prod ON estprod.id_prod_filho = prod.id "
                                                   f"where estprod.id_estrutura = {id_estrutura} "
                                                   f"and prod.codigo = {cod_filho} "
                                                   f"ORDER BY prod.descricao;")
                                    select_estrut = cursor.fetchall()
                                    if select_estrut:
                                        for dads_estrut in select_estrut:
                                            id_est_e, cod_e, des_e, ref_e, um_e, qtde_e, local_e, saldo_e = dads_estrut

                                            cursor = conecta.cursor()
                                            cursor.execute(f"SELECT max(estprod.id), max(prod.codigo), "
                                                           f"max(prod.descricao), "
                                                           f"sum(p_op.QTDE_ESTRUT_PROD)as total "
                                                           f"FROM estrutura_produto as estprod "
                                                           f"INNER JOIN produto as prod ON "
                                                           f"estprod.id_prod_filho = prod.id "
                                                           f"INNER JOIN produtoos as p_op ON "
                                                           f"estprod.id = p_op.ID_ESTRUT_PROD "
                                                           f"where estprod.id_estrutura = {id_estrutura} "
                                                           f"and p_op.numero = {num_op} "
                                                           f"and estprod.id = {id_est_e} "
                                                           f"group by p_op.ID_ESTRUT_PROD;")
                                            select_os_resumo = cursor.fetchall()

                                            if select_os_resumo:
                                                tem_consumo = "SIM"

                        else:
                            tem_consumo = "SIM"

                        dados = (cod_filho, descr, ref, um, qtde_float, ncm, conjunto, tem_consumo)
                        nova_tabela.append(dados)

            if nova_tabela:
                lanca_tabela(self.table_Estrutura, nova_tabela)
                self.pinta_tabela()
                self.lista_original = nova_tabela

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_status(self):
        try:
            cod_pai = self.line_Codigo_Estrut.text()

            versao = self.combo_Versao.currentText()
            versaotete = versao.find(" - ")
            num_versao = versao[8:versaotete]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo FROM produto WHERE codigo = '{cod_pai}';")
            dados_prod = cursor.fetchall()
            id_pai = dados_prod[0][0]

            if id_pai:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, num_versao, data_versao, obs, data_criacao "
                               f"from estrutura "
                               f"where id_produto = {id_pai} and num_versao = {num_versao};")
                tabela_versoes = cursor.fetchall()
                if tabela_versoes:
                    id_estrutura = tabela_versoes[0][0]

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT mov.id, mov.data, mov.codigo, mov.tipo, mov.quantidade, "
                                   f"COALESCE(func.funcionario, '') as fuck, loc.nome, os.numero, ent.natureza, "
                                   f"ent.nota "
                                   f"FROM movimentacao AS mov "
                                   f"LEFT JOIN funcionarios AS func ON mov.funcionario = func.id "
                                   f"INNER JOIN localestoque AS loc ON mov.localestoque = loc.id "
                                   f"LEFT JOIN ordemservico AS os ON mov.id = os.movimentacao "
                                   f"LEFT JOIN ENTRADAPROD as ent ON mov.id = ent.movimentacao "
                                   f"LEFT JOIN produtoos as prodos ON os.id = prodos.mestre "
                                   f"WHERE mov.produto = {id_pai} and os.id_estrutura = {id_estrutura} "
                                   f"and (os.numero IS NOT NULL or ent.natureza IS not NULL);")
                    dados_mov = cursor.fetchall()
                    if dados_mov:
                        for i in dados_mov:
                            id_mov, data, cod, tipo, qtde, func, local, num_op, natur, num_nf = i

                            if natur == 4:
                                self.line_Status.setText("B")
                            if num_op:
                                self.line_Status.setText("B")
                    else:
                        self.line_Status.setText("A")
                else:
                    self.line_Status.setText("A")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pinta_tabela(self):
        try:
            nome_tab = self.table_Estrutura

            status = self.line_Status.text()

            dados_tabela = extrair_tabela(nome_tab)

            if dados_tabela:
                for index, dados in enumerate(dados_tabela):
                    cod_filho, desc, ref, um, qtde, ncm, classe, tem_consumo = dados

                    if tem_consumo == "SIM" and status == "A":
                        num_colunas = len(dados_tabela[0])
                        for i in range(num_colunas):
                            nome_tab.item(index, i).setBackground(QColor(cor_cinza_claro))
                    elif status == "B":
                        num_colunas = len(dados_tabela[0])
                        for i in range(num_colunas):
                            nome_tab.item(index, i).setBackground(QColor(cor_cinza_claro))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_produto_estrutura(self):
        try:
            self.line_Descricao_Estrut.clear()
            self.line_Tipo_Estrut.clear()
            self.line_Referencia_Estrut.clear()
            self.line_NCM_Estrut.clear()
            self.line_UM_Estrut.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_mao_de_obra_servico(self):
        try:
            self.line_Descricao_Mao.clear()
            self.line_Tempo_Mao.clear()

            self.line_Descricao_Servico.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_manu(self):
        try:
            self.line_Codigo_Manu.clear()
            self.line_Descricao_Manu.clear()
            self.line_DescrCompl_Manu.clear()
            self.line_Referencia_Manu.clear()
            self.line_UM_Manu.clear()
            self.line_NCM_Manu.clear()
            self.line_Qtde_Manu.clear()
            self.line_Medida_Manu.clear()
            self.check_Converte_Manu.setChecked(False)
            self.line_Codigo_Manu.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tudo(self):
        self.limpa_tabela()
        self.limpa_dados_produto_estrutura()

        self.combo_Versao.clear()
        self.line_Status.clear()

        self.limpa_dados_mao_de_obra_servico()
        self.line_Obs.clear()

    def mascara_tempo_mao_de_obra(self):
        try:
            tempo_mao = self.line_Tempo_Mao.text()

            if tempo_mao:
                tempo_mao_sem_espacos = tempo_mao.strip()
                if "." in tempo_mao_sem_espacos:
                    string_com_virgula = tempo_mao_sem_espacos.replace('.', ',')
                elif tempo_mao_sem_espacos.startswith(','):
                    string_com_virgula = '0' + tempo_mao_sem_espacos
                else:
                    string_com_virgula = tempo_mao_sem_espacos

                self.line_Tempo_Mao.setText(string_com_virgula)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_codigo_manual(self):
        if not self.processando:
            try:
                self.processando = True

                codigo_produto = self.line_Codigo_Manu.text()
                codigo_pai = self.line_Codigo_Estrut.text()

                status = self.line_Status.text()

                if codigo_pai:
                    if status:
                        if status == "A":
                            if not codigo_produto:
                                self.mensagem_alerta('O campo "Código" não pode estar vazio!')
                                self.line_Codigo_Manu.clear()
                            elif int(codigo_produto) == 0:
                                self.mensagem_alerta('O campo "Código" não pode ser "0"!')
                                self.line_Codigo_Manu.clear()
                            elif codigo_pai == codigo_produto:
                                self.mensagem_alerta('O campo "Código" não pode ser igual ao código da estrutura!')
                                self.line_Codigo_Manu.clear()
                            else:
                                self.verifica_sql_produto_manual()
                        else:
                            self.mensagem_alerta(f'O produto {codigo_pai} tem movimentação e não pode ser alterado!')
                            self.line_Codigo_Manu.clear()

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
                self.lanca_dados_produto_manual()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_produto_manual(self):
        try:
            codigo_produto = self.line_Codigo_Manu.text()
            cur = conecta.cursor()
            cur.execute(f"SELECT descricao, COALESCE(descricaocomplementar, '') as compl, "
                        f"COALESCE(obs, '') as obs, unidade, COALESCE(ncm, '') as local, "
                        f"quantidade, embalagem, kilosmetro "
                        f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            descr, compl, ref, um, ncm, qtde, embalagem, kg_mt = detalhes_produto[0]

            self.line_Descricao_Manu.setText(descr)
            self.line_DescrCompl_Manu.setText(compl)
            self.line_Referencia_Manu.setText(ref)
            self.line_NCM_Manu.setText(ncm)
            self.line_UM_Manu.setText(um)

            if kg_mt == 0:
                kg_mt_string = ''
            else:
                kg_mt_string = str(kg_mt)
            self.line_Kg_Manu.setText(kg_mt_string)

            if um == "KG":
                self.check_Converte_Manu.setChecked(True)
                self.line_Medida_Manu.setFocus()
            else:
                self.check_Converte_Manu.setChecked(False)
                self.line_Qtde_Manu.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_medida_manual(self):
        if not self.processando:
            try:
                self.processando = True

                medida = self.line_Medida_Manu.text()

                if medida:
                    medida_float = valores_para_float(medida)

                    kg_mt = self.line_Kg_Manu.text()
                    if kg_mt:
                        kg_mt_float = valores_para_float(kg_mt)

                        qtde_peso = medida_float * (kg_mt_float/1000)
                        qtde_peso_2casas = ("%.2f" % qtde_peso)

                        qtde_final = valores_para_virgula(qtde_peso_2casas)

                        self.line_Qtde_Manu.setText(qtde_final)
                        self.line_Qtde_Manu.setFocus()
                    else:
                        self.mensagem_alerta('O cadastro do produto precisa ter a informação "KG/MT"!')
                        self.limpa_dados_manu()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_line_qtde_manual(self):
        if not self.processando:
            try:
                self.processando = True

                qtdezinha = self.line_Qtde_Manu.text()

                if len(qtdezinha) == 0:
                    self.mensagem_alerta('O campo "Qtde:" não pode estar vazio')
                    self.line_Qtde_Manu.clear()
                    self.line_Qtde_Manu.setFocus()
                elif qtdezinha == "0":
                    self.mensagem_alerta('O campo "Qtde:" não pode ser "0"')
                    self.line_Qtde_Manu.clear()
                    self.line_Qtde_Manu.setFocus()
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
            cod = self.line_Codigo_Manu.text()

            qtde_manu = self.line_Qtde_Manu.text()
            if "," in qtde_manu:
                qtde_manu_com_ponto = qtde_manu.replace(',', '.')
                qtdezinha_float = float(qtde_manu_com_ponto)
            else:
                qtdezinha_float = float(qtde_manu)

            extrai_estrutura = extrair_tabela(self.table_Estrutura)

            ja_existe = False
            for itens in extrai_estrutura:
                cod_con = itens[0]
                if cod_con == cod:
                    ja_existe = True
                    break

            if not ja_existe:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                               f"COALESCE(prod.ncm, '') as ncm, conj.conjunto, "
                               f"prod.unidade "
                               f"FROM produto as prod "
                               f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                               f"where codigo = {cod};")
                detalhes_produto = cursor.fetchall()
                cod, descr, ref, ncm, conjunto, um = detalhes_produto[0]

                dados1 = [cod, descr, ref, um, qtdezinha_float, ncm, conjunto, "NÃO"]
                extrai_estrutura.append(dados1)

                if extrai_estrutura:
                    lanca_tabela(self.table_Estrutura, extrai_estrutura)
                    self.pinta_tabela()

            else:
                self.mensagem_alerta("Este produto já foi adicionado a estrutura!")

            self.limpa_dados_manu()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_produto_tab(self):
        try:
            nome_tabela = self.table_Estrutura

            cod_pai = self.line_Codigo_Estrut.text()

            status = self.line_Status.text()

            if cod_pai:
                dados_tab = extrair_tabela(nome_tabela)
                if not dados_tab:
                    self.mensagem_alerta(f'A tabela "Produtos Ordem de Compra" está vazia!')
                else:
                    if status:
                        if status == "A":
                            linha = nome_tabela.currentRow()
                            if linha >= 0:
                                cod_filho, desc, ref, um, qtde, ncm, classe, tem_consumo = dados_tab[linha]

                                if tem_consumo == "NÃO":
                                    nome_tabela.removeRow(linha)
                                else:
                                    self.mensagem_alerta(f"O produto {cod_filho} está sendo consumido "
                                                         f"e não pode ser excluído!")
                        else:
                            self.mensagem_alerta(f'O produto {cod_pai} tem movimentação e não pode ser alterado!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_tudo_tab_produtos(self):
        try:
            nome_tabela = self.table_Estrutura

            pode_excluir = True

            cod_pai = self.line_Codigo_Estrut.text()

            status = self.line_Status.text()

            if cod_pai:
                if status:
                    if status == "A":
                        dados_tab = extrair_tabela(nome_tabela)
                        if not dados_tab:
                            self.mensagem_alerta(f'A tabela "Estrutura" está vazia!')
                        else:
                            for i in dados_tab:
                                cod_filho, desc, ref, um, qtde, ncm, classe, tem_consumo = i

                                if tem_consumo == "SIM":
                                    pode_excluir = False
                                    break

                            if pode_excluir:
                                nome_tabela.setRowCount(0)
                            else:
                                self.mensagem_alerta(f"Tem produtos sendo consumidos "
                                                     f"e não podem ser excluídos!")

                    else:
                        self.mensagem_alerta(f'O produto {cod_pai} tem movimentação e não pode ser alterado!')

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpar(self):
        self.limpa_tudo()
        self.line_Codigo_Estrut.clear()
        self.line_Codigo_Estrut.setFocus()

    def verifica_salvamento(self):
        try:
            codigo_produto = self.line_Codigo_Estrut.text()
            status = self.line_Status.text()

            versao = self.combo_Versao.currentText()

            if not codigo_produto:
                self.mensagem_alerta('O campo "Código" não pode estar vazio!')
                self.limpa_dados_produto_estrutura()
                self.limpa_tabela()
            elif int(codigo_produto) == 0:
                self.mensagem_alerta('O campo "Código" não pode ser "0"!')
                self.limpa_dados_produto_estrutura()
                self.limpa_tabela()
            elif not versao:
                self.mensagem_alerta('O campo "Nº Versão" não pode estar vazio!')
                self.limpa_dados_produto_estrutura()
                self.limpa_tabela()
            else:
                if status == "A":
                    extrai_tabela = extrair_tabela(self.table_Estrutura)
                    if not extrai_tabela:
                        if self.pergunta_confirmacao(f'A tabela "Estrutura" está vazia! Deseja mesmo continuar?'):
                            teve_lancado = self.salvar_produtos_estrutura(extrai_tabela)
                            self.salvar_dados_adicionais(teve_lancado)
                    else:
                        tipo_material = self.line_Tipo_Estrut.text()

                        if tipo_material == "INDUSTRIALIZACAO":
                            descr_servico = self.line_Descricao_Servico.text()

                            if not descr_servico:
                                self.mensagem_alerta('O campo "Descrição do Serviço" não pode estar vazio!')
                            else:
                                teve_lancado = self.salvar_produtos_estrutura(extrai_tabela)
                                self.salvar_dados_adicionais(teve_lancado)
                        else:
                            descr_mao = self.line_Descricao_Mao.text()
                            tempo = self.line_Tempo_Mao.text()

                            if not descr_mao or not tempo:
                                self.mensagem_alerta('Os campos "Descrição e Tempo de Serviço" não podem estar vazio!')
                            else:
                                teve_lancado = self.salvar_produtos_estrutura(extrai_tabela)
                                self.salvar_dados_adicionais(teve_lancado)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_produtos_estrutura(self, produtos_modificados):
        try:
            teve_lancado = 0

            versao = self.combo_Versao.currentText()
            versaotete = versao.find(" - ")
            num_versao = versao[8:versaotete]

            cod_prod = self.line_Codigo_Estrut.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, etapas, tempo, obs2 FROM produto where codigo = '{cod_prod}';")
            selects = cursor.fetchall()
            id_prod = selects[0][0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, num_versao, data_versao, obs, data_criacao "
                           f"from estrutura "
                           f"where id_produto = {id_prod} and num_versao = {num_versao};")
            tabela_versoes = cursor.fetchall()

            if tabela_versoes:
                id_estrutura = tabela_versoes[0][0]
            else:
                cursor = conecta.cursor()
                cursor.execute("select GEN_ID(GEN_ESTRUTURA_ID,0) from rdb$database;")
                ultimo_oc0 = cursor.fetchall()
                ultimo_oc1 = ultimo_oc0[0]
                id_estrutura = int(ultimo_oc1[0]) + 1

                data_hoje = date.today()

                obs_versao = "PRIMEIRA VERSAO CRIADA"

                cursor = conecta.cursor()
                cursor.execute(f"Insert into estrutura (id, id_produto, num_versao, data_versao, obs) "
                               f"values (GEN_ID(GEN_ESTRUTURA_ID,1), {id_prod}, '{num_versao}', '{data_hoje}', "
                               f"'{obs_versao}');")

                conecta.commit()

                cursor.execute(f"UPDATE produto SET id_versao = {id_estrutura} "
                               f"WHERE id = {id_prod};")
                conecta.commit()

            produtos_originais_tuple = []
            produtos_modificados_tuple = []

            for i in self.lista_original:
                cod_t, desc_t, ref_t, um_t, qtde_t, ncm_t, classe_t, consumo_t = i

                qtde_float_t = valores_para_float(qtde_t)

                dados = (cod_t, qtde_float_t)
                produtos_originais_tuple.append(dados)

            for ii in produtos_modificados:
                cod_m, desc_m, ref_m, um_m, qtde_m, ncm_m, classe_m, consumo_m = ii

                qtde_float_m = valores_para_float(qtde_m)

                dados = (cod_m, qtde_float_m)
                produtos_modificados_tuple.append(dados)

            lista_saidas = set(produtos_originais_tuple).difference(set(produtos_modificados_tuple))
            if lista_saidas:
                for iii in lista_saidas:
                    cod_sai, qtde_sai = iii

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, etapas, tempo, obs2 FROM produto where codigo = '{cod_sai}';")
                    selects_sai = cursor.fetchall()
                    id_prod_sai = selects_sai[0][0]

                    cursor = conecta.cursor()
                    cursor.execute(f"DELETE FROM estrutura_produto "
                                   f"where id_estrutura = {id_estrutura} "
                                   f"and id_prod_filho = {id_prod_sai};")
                    conecta.commit()

                    teve_lancado += 1

            lista_entradas = set(produtos_modificados_tuple).difference(set(produtos_originais_tuple))
            if lista_entradas:
                for iiii in lista_entradas:
                    cod_ent, qtde_ent = iiii

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, etapas, tempo, obs2 FROM produto where codigo = '{cod_ent}';")
                    selects_ent = cursor.fetchall()
                    id_prod_ent = selects_ent[0][0]

                    cursor = conecta.cursor()
                    cursor.execute(f"Insert into estrutura_produto "
                                   f"(ID, ID_ESTRUTURA, ID_PROD_FILHO, QUANTIDADE) "
                                   f"values (GEN_ID(GEN_ESTRUTURA_PRODUTO_ID,1), "
                                   f"{id_estrutura}, {id_prod_ent}, {qtde_ent});")
                    conecta.commit()

                    teve_lancado += 1

            return teve_lancado

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_dados_adicionais(self, teve_lancado):
        try:
            cod_prod = self.line_Codigo_Estrut.text()

            descr = self.line_Descricao_Estrut.text()

            tipo_material = self.line_Tipo_Estrut.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT terceirizadoobs, etapas, tempo, obs2 FROM produto where codigo = '{cod_prod}';")
            selects = cursor.fetchall()
            terc_servico, etapa_mao, tempo_mao, obs_banco = selects[0]

            tempo_mao_float = valores_para_float(tempo_mao)

            tips_mat = 0
            if tipo_material == "INDUSTRIALIZACAO":
                descr_servico = self.line_Descricao_Servico.text().upper()

                if terc_servico != descr_servico:
                    tips_mat += 1

                    cursor = conecta.cursor()
                    cursor.execute(f"UPDATE produto SET terceirizadoobs = '{descr_servico}' "
                                   f"where codigo = '{cod_prod}';")
                    conecta.commit()

            else:
                descr_mao = self.line_Descricao_Mao.text().upper()
                tempo = self.line_Tempo_Mao.text()
                tempo_float = valores_para_float(tempo)

                if etapa_mao != descr_mao or tempo_mao_float != tempo_float:
                    tips_mat += 1

                    cursor = conecta.cursor()
                    cursor.execute(f"UPDATE produto SET etapas = '{descr_mao}', tempo = '{tempo_float}' "
                                   f"where codigo = '{cod_prod}';")
                    conecta.commit()

            obis = 0
            obs = self.line_Obs.text().upper()
            if obs_banco != obs:
                obis += 1
                cursor = conecta.cursor()
                cursor.execute(f"UPDATE produto SET obs2 = '{obs}' "
                               f"where codigo = '{cod_prod}';")
                conecta.commit()

            if tips_mat or obis or teve_lancado:
                self.mensagem_alerta(f"Cadastro do produto {descr} foi atualizado com Sucesso!!")

                self.limpa_tudo()
                self.line_Codigo_Estrut.clear()
                self.line_Obs.clear()
                self.line_Codigo_Estrut.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaEstruturaIncluirV2("")
    tela.show()
    qt.exec_()

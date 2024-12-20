import sys
from banco_dados.conexao import conecta
from forms.tela_op_status import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.cores import cor_verde_claro, cor_vermelho, cor_cinza_claro, cor_vermelho_claro
from comandos.telas import tamanho_aplicacao, icone
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtGui import QColor
from datetime import date, timedelta
import inspect
import os
import traceback


class TelaOpStatusV2(QMainWindow, Ui_ConsultaOP):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_producao.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_OP)
        layout_cabec_tab(self.table_Estrutura)
        layout_cabec_tab(self.table_Consumo)

        self.btn_Consultar.clicked.connect(self.define_filtros)
        self.line_Codigo.editingFinished.connect(self.manual_verifica_line_codigo)

        self.line_Num_OP.editingFinished.connect(self.tela2_verifica_num_op)
        self.btn_Consultar1.clicked.connect(self.tela2_verifica_num_op)

        self.processando = False

        self.qtde_vezes_select = 0

        self.definir_datas()

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

    def limpa_tabela_op(self):
        try:
            self.table_OP.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tabela_estrut(self):
        try:
            self.table_Estrutura.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tabela_consumo(self):
        try:
            self.table_Consumo.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_datas(self):
        try:
            data_hoje = date.today()
            self.date_Emissao2.setDate(data_hoje)
            self.date_Entrega2.setDate(data_hoje)

            data_entrega = data_hoje - timedelta(days=15)
            self.date_Emissao1.setDate(data_entrega)
            self.date_Entrega1.setDate(data_entrega)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pintar_tabela(self):
        try:
            qtable_widget = self.table_OP

            extrai_tabela = extrair_tabela(qtable_widget)
            if extrai_tabela:
                for index, itens in enumerate(extrai_tabela):
                    emissao, previsao, num_op, cod, descr, ref, um, qtde, msg, status = itens

                    if status == "B":
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
                    else:
                        posicao = msg.find("/")
                        qtde_estrut = msg[:posicao]
                        qtde_consumo = msg[posicao + 1:]

                        if qtde_estrut == "0":
                            qtable_widget.item(index, 0).setBackground(QColor(cor_vermelho))
                            qtable_widget.item(index, 1).setBackground(QColor(cor_vermelho))
                            qtable_widget.item(index, 2).setBackground(QColor(cor_vermelho))
                            qtable_widget.item(index, 3).setBackground(QColor(cor_vermelho))
                            qtable_widget.item(index, 4).setBackground(QColor(cor_vermelho))
                            qtable_widget.item(index, 5).setBackground(QColor(cor_vermelho))
                            qtable_widget.item(index, 6).setBackground(QColor(cor_vermelho))
                            qtable_widget.item(index, 7).setBackground(QColor(cor_vermelho))
                            qtable_widget.item(index, 8).setBackground(QColor(cor_vermelho))
                            qtable_widget.item(index, 9).setBackground(QColor(cor_vermelho))
                        elif qtde_consumo == "0":
                            qtable_widget.item(index, 0).setBackground(QColor(cor_vermelho_claro))
                            qtable_widget.item(index, 1).setBackground(QColor(cor_vermelho_claro))
                            qtable_widget.item(index, 2).setBackground(QColor(cor_vermelho_claro))
                            qtable_widget.item(index, 3).setBackground(QColor(cor_vermelho_claro))
                            qtable_widget.item(index, 4).setBackground(QColor(cor_vermelho_claro))
                            qtable_widget.item(index, 5).setBackground(QColor(cor_vermelho_claro))
                            qtable_widget.item(index, 6).setBackground(QColor(cor_vermelho_claro))
                            qtable_widget.item(index, 7).setBackground(QColor(cor_vermelho_claro))
                            qtable_widget.item(index, 8).setBackground(QColor(cor_vermelho_claro))
                            qtable_widget.item(index, 9).setBackground(QColor(cor_vermelho_claro))
                        elif qtde_estrut == qtde_consumo:
                            qtable_widget.item(index, 0).setBackground(QColor(cor_verde_claro))
                            qtable_widget.item(index, 1).setBackground(QColor(cor_verde_claro))
                            qtable_widget.item(index, 2).setBackground(QColor(cor_verde_claro))
                            qtable_widget.item(index, 3).setBackground(QColor(cor_verde_claro))
                            qtable_widget.item(index, 4).setBackground(QColor(cor_verde_claro))
                            qtable_widget.item(index, 5).setBackground(QColor(cor_verde_claro))
                            qtable_widget.item(index, 6).setBackground(QColor(cor_verde_claro))
                            qtable_widget.item(index, 7).setBackground(QColor(cor_verde_claro))
                            qtable_widget.item(index, 8).setBackground(QColor(cor_verde_claro))
                            qtable_widget.item(index, 9).setBackground(QColor(cor_verde_claro))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manual_verifica_line_codigo(self):
        if not self.processando:
            try:
                self.processando = True

                codigo_produto = self.line_Codigo.text()
                if len(codigo_produto) == 0:
                    self.mensagem_alerta('O campo "Código" não pode estar vazio')
                    self.line_Codigo.clear()
                elif int(codigo_produto) == 0:
                    self.mensagem_alerta('O campo "Código" não pode ser "0"')
                    self.line_Codigo.clear()
                else:
                    self.manual_verifica_sql_produto()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def manual_verifica_sql_produto(self):
        try:
            codigo_produto = self.line_Codigo.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade, localizacao, quantidade "
                           f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cursor.fetchall()
            if not detalhes_produto:
                self.mensagem_alerta('Este código de produto não existe!')
                self.line_Codigo.clear()
            else:
                self.manual_verifica_materia_prima()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manual_verifica_materia_prima(self):
        try:
            codigo_produto = self.line_Codigo.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, descricao, COALESCE(obs, '') as obs, unidade, terceirizado, "
                           f"conjunto, tipomaterial "
                           f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cursor.fetchall()

            id_prod, descricao, ref, um, terceirizado, conjunto, tipo = detalhes_produto[0]

            if conjunto == 10:
                self.manual_lanca_dados_produto()
            else:
                self.mensagem_alerta("Este produto não está definido como produto acabado!")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manual_lanca_dados_produto(self):
        try:
            codigo_produto = self.line_Codigo.text()

            cur = conecta.cursor()
            cur.execute(f"SELECT descricao, COALESCE(descricaocomplementar, '') as compl, "
                        f"COALESCE(obs, '') as obs, unidade, COALESCE(ncm, '') as local, "
                        f"quantidade, embalagem FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            descr, compl, ref, um, ncm, saldo, embalagem = detalhes_produto[0]

            self.line_Descricao.setText(descr)
            self.line_Referencia.setText(ref)
            self.line_UM.setText(um)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tudo_tela1(self):
        try:
            self.limpa_tabela_op()

            self.line_Codigo.clear()
            self.line_UM.clear()
            self.line_Referencia.clear()
            self.line_Descricao.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def conversao_manipula_dados(self, op_abertas):
        try:
            self.limpa_tudo_tela1()

            if op_abertas:
                op_ab_editado = []
                for dados_op in op_abertas:
                    emissao, previsao, op, cod, descr, ref, um, qtde, status, id_estrut = dados_op

                    if id_estrut:
                        data_em_texto = '{}/{}/{}'.format(emissao.day, emissao.month, emissao.year)

                        if previsao:
                            data_prev = '{}/{}/{}'.format(previsao.day, previsao.month, previsao.year)
                        else:
                            data_prev = ''

                        cursor = conecta.cursor()
                        cursor.execute(f"SELECT id, codigo FROM produto where codigo = {cod};")
                        select_prod = cursor.fetchall()

                        idez, cod = select_prod[0]

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

                        dados = (data_em_texto, data_prev, op, cod, descr, ref, um, qtde, msg, status)
                        op_ab_editado.append(dados)

                if op_ab_editado:
                    lanca_tabela(self.table_OP, op_ab_editado)

            else:
                self.mensagem_alerta("Não foi encontrado nenhum resultado para esta pesquisa")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def define_filtros(self):
        try:
            cod = self.line_Codigo.text()
            aber = False
            baix = False
            emi = False
            ent = False
            cem = False
            zero = False

            if self.check_Aberto.isChecked():
                aber = True

            if self.check_Baixado.isChecked():
                baix = True

            if self.check_Emissao.isChecked():
                emi = True

            if self.check_Entrega.isChecked():
                ent = True

            if self.check_Todo_Consumo.isChecked():
                cem = True

            if self.check_0_Consumo.isChecked():
                zero = True

            if emi and not ent and not aber and not baix and not cod and not cem and not zero:
                print("só emissao")
                self.manipula_emissao()
            elif emi and not ent and aber and not baix and not cod and not cem and not zero:
                print("só emissao e aberto")
                self.manipula_emissao_e_aberto()
            elif emi and not ent and not aber and baix and not cod and not cem and not zero:
                print("só emissao e baixado")
                self.manipula_emissao_e_baixado()
            elif not emi and ent and not aber and not baix and not cod and not cem and not zero:
                print("só entrega")
                self.manipula_entrega()
            elif not emi and ent and aber and not baix and not cod and not cem and not zero:
                print("só entrega e aberto")
                self.manipula_entrega_e_aberto()
            elif not emi and ent and not aber and baix and not cod and not cem and not zero:
                print("só entrega e baixado")
                self.manipula_entrega_e_baixado()
            elif not emi and not ent and aber and not baix and not cod and not cem and not zero:
                print("só aberto")
                self.manipula_aberto()
            elif not emi and not ent and not aber and not baix and cod and not cem and not zero:
                print("só código")
                self.manipula_codigo()
            elif not emi and not ent and aber and not baix and cod and not cem and not zero:
                print("só código e aberto")
                self.manipula_codigo_e_aberto()
            elif not emi and not ent and not aber and baix and cod and not cem and not zero:
                print("só código e baixado")
                self.manipula_codigo_e_fechado()
            elif not emi and not ent and aber and baix and cod and not cem and not zero:
                print("só código e aberto e baixado")
                self.manipula_codigo_e_aberto_e_fechado()
            elif not emi and not ent and aber and not baix and not cod and cem and not zero:
                print("só 100 e aberto")
                self.manipula_cem_e_aberto()
            elif not emi and not ent and aber and not baix and not cod and not cem and zero:
                print("só zero e aberto")
                self.manipula_zero_e_aberto()
            else:
                self.check_Aberto.setChecked(True)

                self.check_Baixado.setChecked(False)
                self.check_Emissao.setChecked(False)
                self.check_Entrega.setChecked(False)
                self.check_Todo_Consumo.setChecked(False)
                self.check_0_Consumo.setChecked(False)

            self.pintar_tabela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_aberto(self):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, "
                           f"prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where ordser.status = 'A' order by ordser.numero;")
            op_abertas = cursor.fetchall()

            self.conversao_manipula_dados(op_abertas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_emissao(self):
        try:
            data_ini = self.date_Emissao1.date().toString("yyyy-MM-dd")
            data_fim = self.date_Emissao2.date().toString("yyyy-MM-dd")

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.codigo, "
                           f"prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"FROM ordemservico AS ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"WHERE ordser.datainicial BETWEEN '{data_ini}' AND '{data_fim}' "
                           f"ORDER BY ordser.numero;")
            op_abertas = cursor.fetchall()

            self.conversao_manipula_dados(op_abertas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_emissao_e_aberto(self):
        try:
            data_ini = self.date_Emissao1.date().toString("yyyy-MM-dd")
            data_fim = self.date_Emissao2.date().toString("yyyy-MM-dd")

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.codigo, "
                           f"prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"FROM ordemservico AS ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"WHERE ordser.datainicial BETWEEN '{data_ini}' AND '{data_fim}' and ordser.status = 'A' "
                           f"ORDER BY ordser.numero;")
            op_abertas = cursor.fetchall()

            self.conversao_manipula_dados(op_abertas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_emissao_e_baixado(self):
        try:
            data_ini = self.date_Emissao1.date().toString("yyyy-MM-dd")
            data_fim = self.date_Emissao2.date().toString("yyyy-MM-dd")

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.codigo, "
                           f"prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"FROM ordemservico AS ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"WHERE ordser.datainicial BETWEEN '{data_ini}' AND '{data_fim}' and ordser.status = 'B' "
                           f"ORDER BY ordser.numero;")
            op_abertas = cursor.fetchall()

            self.conversao_manipula_dados(op_abertas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_entrega(self):
        try:
            data_ini = self.date_Entrega1.date().toString("yyyy-MM-dd")
            data_fim = self.date_Entrega2.date().toString("yyyy-MM-dd")

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.codigo, "
                           f"prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"FROM ordemservico AS ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"WHERE ordser.dataprevisao BETWEEN '{data_ini}' AND '{data_fim}' "
                           f"ORDER BY ordser.numero;")
            op_abertas = cursor.fetchall()

            self.conversao_manipula_dados(op_abertas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_entrega_e_aberto(self):
        try:
            data_ini = self.date_Entrega1.date().toString("yyyy-MM-dd")
            data_fim = self.date_Entrega2.date().toString("yyyy-MM-dd")

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.codigo, "
                           f"prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"FROM ordemservico AS ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"WHERE ordser.dataprevisao BETWEEN '{data_ini}' AND '{data_fim}' and ordser.status = 'A' "
                           f"ORDER BY ordser.numero;")
            op_abertas = cursor.fetchall()

            self.conversao_manipula_dados(op_abertas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_entrega_e_baixado(self):
        try:
            data_ini = self.date_Entrega1.date().toString("yyyy-MM-dd")
            data_fim = self.date_Entrega2.date().toString("yyyy-MM-dd")

            cursor = conecta.cursor()
            cursor.execute(f"SELECT ordser.datainicial, ordser.dataprevisao, ordser.numero, prod.codigo, "
                           f"prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"FROM ordemservico AS ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"WHERE ordser.dataprevisao BETWEEN '{data_ini}' AND '{data_fim}' and ordser.status = 'B' "
                           f"ORDER BY ordser.numero;")
            op_abertas = cursor.fetchall()

            self.conversao_manipula_dados(op_abertas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_codigo(self):
        try:
            cod = self.line_Codigo.text()

            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, "
                           f"prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where prod.codigo = {cod} order by ordser.numero;")
            op_abertas = cursor.fetchall()

            self.conversao_manipula_dados(op_abertas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_codigo_e_aberto(self):
        try:
            cod = self.line_Codigo.text()

            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, "
                           f"prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where prod.codigo = {cod} and ordser.status = 'A' order by ordser.numero;")
            op_abertas = cursor.fetchall()

            self.conversao_manipula_dados(op_abertas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_codigo_e_fechado(self):
        try:
            cod = self.line_Codigo.text()

            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, "
                           f"prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where prod.codigo = {cod} and ordser.status = 'B' order by ordser.numero;")
            op_abertas = cursor.fetchall()

            self.conversao_manipula_dados(op_abertas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_codigo_e_aberto_e_fechado(self):
        try:
            cod = self.line_Codigo.text()

            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, "
                           f"prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where prod.codigo = {cod} "
                           f"order by ordser.numero;")
            op_abertas = cursor.fetchall()

            self.conversao_manipula_dados(op_abertas)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_cem_e_aberto(self):
        try:
            self.limpa_tudo_tela1()

            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, "
                           f"prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where ordser.status = 'A' order by ordser.numero;")
            op_abertas = cursor.fetchall()

            if op_abertas:
                op_ab_editado = []
                for dados_op in op_abertas:
                    emissao, previsao, op, cod, descr, ref, um, qtde, status, id_estrut = dados_op

                    data_em_texto = '{}/{}/{}'.format(emissao.day, emissao.month, emissao.year)

                    if previsao:
                        data_prev = '{}/{}/{}'.format(previsao.day, previsao.month, previsao.year)
                    else:
                        data_prev = ''

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, codigo FROM produto where codigo = {cod};")
                    select_prod = cursor.fetchall()

                    idez, cod = select_prod[0]

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
                        cursor.execute(f"SELECT max(prodser.ID_ESTRUT_PROD), sum(prodser.QTDE_ESTRUT_PROD) as total "
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

                    if total_estrut > 0:
                        if total_estrut == total_consumo:
                            dados = (data_em_texto, data_prev, op, cod, descr, ref, um, qtde, msg, status)
                            op_ab_editado.append(dados)

                if op_ab_editado:
                    lanca_tabela(self.table_OP, op_ab_editado)
                else:
                    self.mensagem_alerta("Não foi encontrado nenhum resultado para esta pesquisa")

            else:
                self.mensagem_alerta("Não foi encontrado nenhum resultado para esta pesquisa")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_zero_e_aberto(self):
        try:
            self.limpa_tudo_tela1()

            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.dataprevisao, ordser.numero, "
                           f"prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"ordser.quantidade, ordser.status, ordser.ID_ESTRUTURA "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where ordser.status = 'A' order by ordser.numero;")
            op_abertas = cursor.fetchall()

            if op_abertas:
                op_ab_editado = []
                for dados_op in op_abertas:
                    emissao, previsao, op, cod, descr, ref, um, qtde, status, id_estrut = dados_op

                    data_em_texto = '{}/{}/{}'.format(emissao.day, emissao.month, emissao.year)

                    if previsao:
                        data_prev = '{}/{}/{}'.format(previsao.day, previsao.month, previsao.year)
                    else:
                        data_prev = ''

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT id, codigo FROM produto where codigo = {cod};")
                    select_prod = cursor.fetchall()

                    idez, cod = select_prod[0]

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
                        cursor.execute(f"SELECT max(prodser.ID_ESTRUT_PROD), sum(prodser.qtde_materia)as total "
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

                    if total_consumo == 0:
                        dados = (data_em_texto, data_prev, op, cod, descr, ref, um, qtde, msg, status)
                        op_ab_editado.append(dados)

                if op_ab_editado:
                    lanca_tabela(self.table_OP, op_ab_editado)
                else:
                    self.mensagem_alerta("Não foi encontrado nenhum resultado para esta pesquisa")

            else:
                self.mensagem_alerta("Não foi encontrado nenhum resultado para esta pesquisa")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tudo_tela2(self):
        try:
            self.limpa_tabela_estrut()
            self.limpa_tabela_consumo()

            self.line_Codigo_Manu.clear()
            self.line_UM_Manu.clear()
            self.line_Referencia_Manu.clear()
            self.line_Descricao_Manu.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tela2_verifica_num_op(self):
        if not self.processando:
            try:
                self.processando = True

                id_os, num_op, data_emissao, status_os, produto_os, qtde_os, obs, id_estrut = self.tela2_dados_op()

                if not num_op:
                    self.mensagem_alerta('O campo "Nº OP" não pode estar vazio')
                    self.reiniciar()
                elif int(num_op) == 0:
                    self.mensagem_alerta('O campo "Nº OP" não pode ser "0"')
                    self.reiniciar()
                else:
                    self.tela2_verifica_sql_op()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def tela2_verifica_sql_op(self):
        try:
            id_os, num_op, data_emissao, status_os, produto_os, qtde_os, obs, id_estrut = self.tela2_dados_op()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT numero, datainicial, status, produto, quantidade "
                           f"FROM ordemservico where numero = {num_op};")
            extrair_dados = cursor.fetchall()
            if not extrair_dados:
                self.mensagem_alerta('Este número de "OP" não existe!')

            else:
                self.tela2_verifica_vinculo_materia()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tela2_verifica_vinculo_materia(self):
        try:
            id_os, num_op, data_emissao, status_os, produto_os, qtde_os, obs, id_estrut = self.tela2_dados_op()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT codigo, ID_ESTRUT_PROD, QTDE_ESTRUT_PROD FROM produtoos where numero = {num_op};")
            itens_os = cursor.fetchall()

            verifica_cadastro = 0
            for item in itens_os:
                codigo, id_materia, qtde_materia = item
                if not id_materia and not qtde_materia:
                    verifica_cadastro = verifica_cadastro + 1

            if verifica_cadastro > 0:
                self.mensagem_alerta('O material consumido não está vinculado com a estrutura!')
                self.reiniciar()
            else:
                self.tela2_verifica_dados_op()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tela2_verifica_dados_op(self):
        try:
            id_os, num_op, data_emissao, status_os, produto_os, qtde_os, obs, id_estrut = self.tela2_dados_op()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, ' ') as obs, prod.unidade, "
                           f"((SELECT quantidade FROM ordemservico where numero = {num_op}) * "
                           f"(estprod.quantidade)) AS Qtde, "
                           f"prod.localizacao, prod.quantidade "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id_estrutura = {id_estrut} ORDER BY prod.descricao;")
            itens_select_estrut = cursor.fetchall()
            if not itens_select_estrut:
                self.mensagem_alerta('Este material não tem estrutura cadastrada!')
                self.reiniciar()
            elif produto_os is None:
                self.mensagem_alerta('Esta "OP" está sem código de produto!')
                self.reiniciar()
            elif qtde_os is None:
                self.mensagem_alerta('A quantidade da "OP" deve ser maior que "0"!')
                self.reiniciar()
            elif num_op is None:
                self.mensagem_alerta('O número da "OP" deve ser maior que "0"!')
                self.reiniciar()
            else:
                self.tela2_lanca_dados_op()
                self.tela2_separar_dados_select()
                self.tela2_pintar_tabelas()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tela2_lanca_dados_op(self):
        try:
            id_os, num_op, data_emissao, status_os, produto_os, qtde_os, obs, id_estrut = self.tela2_dados_op()

            cur = conecta.cursor()
            cur.execute("SELECT codigo, descricao, COALESCE(obs, ' ') as obs, unidade "
                        "FROM produto where id = '{}';".format(produto_os))
            detalhes_produtos = cur.fetchall()
            codigo_id, descricao_id, referencia_id, unidade_id = detalhes_produtos[0]

            self.date_Emissao3.setDate(data_emissao)

            self.line_Codigo_Manu.setText(codigo_id)
            self.line_Descricao_Manu.setText(descricao_id)
            self.line_Referencia_Manu.setText(referencia_id)
            self.line_UM_Manu.setText(unidade_id)
            numero = str(qtde_os).replace('.', ',')
            self.line_Qtde_Manu.setText(numero)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tela2_dados_op(self):
        try:
            numero_os_line = self.line_Num_OP.text()
            cur = conecta.cursor()
            cur.execute(f"SELECT id, numero, datainicial, status, produto, quantidade, obs, ID_ESTRUTURA "
                        f"FROM ordemservico where numero = {numero_os_line};")
            extrair_dados = cur.fetchall()
            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, obs, id_estrut = extrair_dados[0]

            return id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, obs, id_estrut

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tela2_separar_dados_select(self):
        try:
            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, obs, id_estrut = self.tela2_dados_op()

            itens_manipula_total = self.tela2_select_mistura()

            self.qtde_vezes_select = self.qtde_vezes_select + 1

            tabela_estrutura = []
            tabela_consumo_os = []

            for itens in itens_manipula_total:
                id_mat, cod_est, descr_est, ref_est, um_est, qtde_est, local, saldo, \
                data_os, cod_os, descr_os, ref_os, um_os, qtde_os = itens

                qtde_est_str = str(qtde_est)
                qtde_est_float = float(qtde_est_str)
                qtde_est_red = "%.3f" % qtde_est_float

                if qtde_os == "":
                    qtde_os_red = qtde_os
                else:
                    qtde_os_str = str(qtde_os)
                    qtde_os_float = float(qtde_os_str)
                    qtde_os_red = "%.3f" % qtde_os_float

                lista_est = (id_mat, cod_est, descr_est, um_est, qtde_est_red)
                tabela_estrutura.append(lista_est)

                lista_os = (id_mat, data_os, cod_os, descr_os, um_os, qtde_os_red)
                tabela_consumo_os.append(lista_os)

            lanca_tabela(self.table_Consumo, tabela_consumo_os)
            lanca_tabela(self.table_Estrutura, tabela_estrutura)

            if status_os == "B":
                msg = "Ordem de Produção Encerrada"
                self.label_Status.setText(msg)
            else:
                msg = "Ordem de Produção Aberta"
                self.label_Status.setText(msg)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tela2_select_mistura(self):
        try:
            id_os, numero_os, data_emissao, status_os, produto_os, qtde_os, obs, id_estrut = self.tela2_dados_op()

            dados_para_tabela = []
            campo_br = ""

            cursor = conecta.cursor()
            cursor.execute(f"SELECT estprod.id, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, ' ') as obs, prod.unidade, "
                           f"((SELECT quantidade FROM ordemservico where numero = {numero_os}) * "
                           f"(estprod.quantidade)) AS Qtde, "
                           f"prod.localizacao, prod.quantidade "
                           f"FROM estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"where estprod.id_estrutura = {id_estrut} ORDER BY prod.descricao;")
            select_estrut = cursor.fetchall()

            for dados_estrut in select_estrut:
                id_mat_e, cod_e, descr_e, ref_e, um_e, qtde_e, local_e, saldo_e = dados_estrut

                cursor = conecta.cursor()
                cursor.execute(f"SELECT max(estprod.id), max(prod.codigo), max(prod.descricao), "
                               f"sum(prodser.QTDE_ESTRUT_PROD) as total "
                               f"FROM estrutura_produto as estprod "
                               f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                               f"INNER JOIN produtoos as prodser ON estprod.id = prodser.ID_ESTRUT_PROD "
                               f"where estprod.id_estrutura = {id_estrut} "
                               f"and prodser.numero = {numero_os} and estprod.id = {id_mat_e} "
                               f"group by prodser.ID_ESTRUT_PROD;")
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
                    cursor.execute(f"select prodser.ID_ESTRUT_PROD, "
                                   f"COALESCE((extract(day from prodser.data)||'/'||"
                                   f"extract(month from prodser.data)||'/'||"
                                   f"extract(year from prodser.data)), '') AS DATA, prod.codigo, prod.descricao, "
                                   f"COALESCE(prod.obs, '') as obs, prod.unidade, "
                                   f"prodser.quantidade, prodser.QTDE_ESTRUT_PROD "
                                   f"from produtoos as prodser "
                                   f"INNER JOIN produto as prod ON prodser.produto = prod.id "
                                   f"where prodser.numero = {numero_os} and prodser.ID_ESTRUT_PROD = {id_mat_e};")
                    select_os = cursor.fetchall()

                    for dados_os in select_os:
                        id_mat_os, data_os, cod_os, descr_os, ref_os, um_os, qtde_os, qtde_mat_os = dados_os

                        dados2 = (id_mat_e, cod_e, descr_e, ref_e, um_e, qtde_mat_os, local_e, saldo_e,
                                  data_os, cod_os, descr_os, ref_os, um_os, qtde_os)
                        dados_para_tabela.append(dados2)

            return dados_para_tabela

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tela2_pintar_tabelas(self):
        try:
            extrai_total = self.tela2_jutando_tabelas_extraidas()

            for index, itens in enumerate(extrai_total):
                id_mat, cod_est, descr_est, um_est, qtde_est, \
                data_os, cod_os, descr_os, um_os, qtde_os = itens

                if cod_os:
                    self.table_Estrutura.item(index, 0).setBackground(QColor(cor_cinza_claro))
                    self.table_Estrutura.item(index, 1).setBackground(QColor(cor_cinza_claro))
                    self.table_Estrutura.item(index, 2).setBackground(QColor(cor_cinza_claro))
                    self.table_Estrutura.item(index, 3).setBackground(QColor(cor_cinza_claro))
                    self.table_Estrutura.item(index, 4).setBackground(QColor(cor_cinza_claro))

                    self.table_Consumo.item(index, 0).setBackground(QColor(cor_cinza_claro))
                    self.table_Consumo.item(index, 1).setBackground(QColor(cor_cinza_claro))
                    self.table_Consumo.item(index, 2).setBackground(QColor(cor_cinza_claro))
                    self.table_Consumo.item(index, 3).setBackground(QColor(cor_cinza_claro))
                    self.table_Consumo.item(index, 4).setBackground(QColor(cor_cinza_claro))
                    self.table_Consumo.item(index, 5).setBackground(QColor(cor_cinza_claro))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def tela2_jutando_tabelas_extraidas(self):
        try:
            estrutura = extrair_tabela(self.table_Estrutura)
            consumo_os = extrair_tabela(self.table_Consumo)

            linhas_est = len(estrutura)
            extrai_total = []

            for linha_est in range(linhas_est):
                id_mat_est, cod_est, descr_est, um_est, qtde_est = estrutura[linha_est]

                id_mat_os, data_os, cod_os, descr_os, um_os, qtde_os = consumo_os[linha_est]

                extrai_todos = (id_mat_est, cod_est, descr_est, um_est, qtde_est,
                                data_os, cod_os, descr_os, um_os, qtde_os)
                extrai_total.append(extrai_todos)
            return extrai_total

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def reiniciar(self):
        try:
            self.line_Num_OP.clear()
            self.line_Codigo_Manu.clear()
            self.line_Referencia_Manu.clear()
            self.line_Descricao_Manu.clear()
            self.line_UM_Manu.clear()
            self.line_Qtde_Manu.clear()
            self.line_UM_Manu.clear()
            self.table_Estrutura.setRowCount(0)
            self.table_Consumo.setRowCount(0)

            self.line_Num_OP.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    opconsulta = TelaOpStatusV2()
    opconsulta.show()
    qt.exec_()

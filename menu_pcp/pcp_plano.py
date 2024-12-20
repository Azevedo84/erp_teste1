import sys
from banco_dados.conexao import conecta
from forms.tela_pcp_plano import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.telas import tamanho_aplicacao, icone
from PyQt5.QtWidgets import QApplication, QMainWindow, QShortcut, QMessageBox
from PyQt5.QtGui import QKeySequence, QFont, QColor
from PyQt5.QtCore import Qt
import inspect
import os
import traceback


class TelaPcpPlano(QMainWindow, Ui_ConsultaPcp_Prod):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)
        
        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_producao.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Estrutura)

        self.tab_shortcut = QShortcut(QKeySequence(Qt.Key_Tab), self)
        self.tab_shortcut.activated.connect(self.manipula_tab)

        self.btn_Consultar_Cod.clicked.connect(self.verifica_line_codigo)
        self.line_Codigo.returnPressed.connect(lambda: self.verifica_line_codigo())

        self.btn_Consulta_Estrut.clicked.connect(self.verifica_line_qtde)
        self.line_Qtde.returnPressed.connect(lambda: self.verifica_line_qtde())

        self.configura_label()
        self.line_Codigo.setFocus()

        self.vermelho = "#ff0000"
        self.branco = "#ffffff"

        cursor = conecta.cursor()
        cursor.execute(f"SELECT tip.id||' - '||tip.tipomaterial FROM tipomaterial as tip order by tip.tipomaterial;")
        setores = cursor.fetchall()

        self.checkBox_Crescente.setChecked(True)

        geek_list = [setor[0] for setor in setores]
        self.combo_Tipo.addItems(geek_list)

        self.combo_Classifica.setCurrentIndex(1)

        self.combo_Tipo.activated.connect(self.verifica_line_qtde)
        self.combo_Classifica.activated.connect(self.verifica_line_qtde)

        self.checkBox_Crescente.stateChanged.connect(self.verifica_line_qtde)
        self.checkBox_Comprados.stateChanged.connect(self.verifica_line_qtde)
        self.checkBox_Sem_Saldo.stateChanged.connect(self.verifica_line_qtde)
        self.checkBox_Compra_Pend.stateChanged.connect(self.verifica_line_qtde)
        self.checkBox_OP_Pend.stateChanged.connect(self.verifica_line_qtde)
        self.checkBox_Saldo_Acabado.stateChanged.connect(self.verifica_line_qtde)

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

    def manipula_tab(self):
        try:
            if self.line_Codigo.hasFocus():
                self.verifica_line_codigo()

            elif self.line_Qtde.hasFocus():
                self.verifica_line_qtde()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def configura_label(self):
        try:
            validator = QtGui.QIntValidator(0, 123456, self.line_Codigo)
            locale = QtCore.QLocale("pt_BR")
            validator.setLocale(locale)
            self.line_Codigo.setValidator(validator)

            validator = QtGui.QIntValidator(0, 123456, self.line_Qtde)
            locale = QtCore.QLocale("pt_BR")
            validator.setLocale(locale)
            self.line_Qtde.setValidator(validator)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pintar_tabela(self):
        try:
            extrai_tabela = extrair_tabela(self.table_Estrutura)

            testinho = 0

            for itens in extrai_tabela:
                cod, descr, ref, um, qtde, conj, saldo = itens
                qtde_float = float(qtde)
                saldo_float = float(saldo)

                if saldo_float < qtde_float:
                    testinho = testinho + 1
                    testinho2 = testinho - 1

                    font = QFont()
                    font.setBold(True)

                    self.table_Estrutura.item(testinho2, 4).setBackground(QColor(self.vermelho))
                    self.table_Estrutura.item(testinho2, 4).setFont(font)
                    self.table_Estrutura.item(testinho2, 4).setForeground(QColor(self.branco))

                else:
                    testinho = testinho + 1

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_codigo(self):
        try:
            codigo_produto = self.line_Codigo.text()
            if len(codigo_produto) == 0:
                self.mensagem_alerta('O campo "Código" não pode estar vazio')
                self.line_Codigo.clear()
            elif int(codigo_produto) == 0:
                self.mensagem_alerta('O campo "Código" não pode ser "0"')
                self.line_Codigo.clear()
            else:
                self.verifica_sql_codigo()
        
        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_sql_codigo(self):
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
                self.lanca_dados_codigo()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_codigo(self):
        try:
            codigo_produto = self.line_Codigo.text()
            cur = conecta.cursor()
            cur.execute(f"SELECT prod.descricao, COALESCE(prod.obs, ' ') as obs, prod.unidade, prod.localizacao, "
                        f"prod.quantidade, conj.conjunto "
                        f"FROM produto as prod "
                        f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                        f"where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            descricao_id, referencia_id, unidade_id, local_id, quantidade_id, conj = detalhes_produto[0]

            quantidade_id_float = float(quantidade_id)
            numero = str(quantidade_id).replace('.', ',')

            if quantidade_id_float < 0:
                self.mensagem_alerta(f'Este produto está com saldo negativo!\n'
                                                            f'Saldo Total = {quantidade_id_float}')
                self.line_Codigo.clear()
            else:
                self.line_Descricao.setText(descricao_id)
                self.line_Ref.setText(referencia_id)
                self.line_UM.setText(unidade_id)
                self.line_Saldo.setText(numero)
                self.line_Conjunto.setText(conj)
                self.line_Qtde.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_qtde(self):
        try:
            codigo_produto = self.line_Codigo.text()
            if len(codigo_produto) == 0:
                self.mensagem_alerta('O campo "Código" não pode estar vazio')
                self.line_Codigo.clear()
            elif int(codigo_produto) == 0:
                self.mensagem_alerta('O campo "Código" não pode ser "0"')
                self.line_Codigo.clear()
            else:
                qtdezinha = self.line_Qtde.text()
                if not qtdezinha:
                    self.mensagem_alerta('O campo "Qtde:" não pode estar vazio')
                    self.line_Qtde.clear()
                    self.line_Qtde.setFocus()
                else:
                    if "," in qtdezinha:
                        qtdezinha_com_ponto = qtdezinha.replace(',', '.')
                        qtdezinha_float = float(qtdezinha_com_ponto)
                    else:
                        qtdezinha_float = float(qtdezinha)

                    if qtdezinha_float == 0:
                        self.mensagem_alerta('O campo "Qtde:" não pode ser "0"')
                        self.line_Qtde.clear()
                        self.line_Qtde.setFocus()
                    else:
                        self.classifica_inicio()
                        self.pintar_tabela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def sql_estrutura(self, codigo, quantidade):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.id, prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, ({quantidade}) as qtde, "
                           f"prod.quantidade "
                           f"FROM produto as prod where prod.codigo = {codigo};")
            select_prod = cursor.fetchall()

            idez, cod0, descr0, ref0, um0, qtde0, saldo0 = select_prod[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"prod.unidade, (mat.quantidade * {quantidade}) as qtde, "
                           f"COALESCE(tip.tipomaterial, '') as tip, "
                           f"prod.quantidade "
                           f"from materiaprima as mat "
                           f"INNER JOIN produto prod ON mat.produto = prod.id "
                           f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                           f"LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id "
                           f"where mat.mestre = {idez} "
                           f"order by conj.conjunto DESC, prod.localizacao ASC;")
            tabela_estrutura = cursor.fetchall()

            return tabela_estrutura

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_saldo_acabado(self, codigo, quantidade):
        try:
            tabela_codigus = []

            tabela = self.sql_estrutura(codigo, quantidade)

            for dados in tabela:
                cod, descr, ref, um, qtde, tipo, saldo = dados

                folha00 = (cod, descr, ref, um, qtde, tipo, saldo)
                tabela_codigus.append(folha00)

                qtde_float = float(qtde)
                saldo_float = float(saldo)

                sobra = qtde_float - saldo_float

                if sobra > 0:
                    tabela1 = self.sql_estrutura(cod, sobra)
                    for dados1 in tabela1:
                        cod1, descr1, ref1, um1, qtde1, tipo1, saldo1 = dados1

                        folha_ = (cod1, descr1, ref1, um1, qtde1, tipo1, saldo1)
                        tabela_codigus.append(folha_)

                        qtde_float1 = float(qtde1)
                        saldo_float1 = float(saldo1)

                        sobra1 = qtde_float1 - saldo_float1

                        if sobra1 > 0:
                            tabela2 = self.sql_estrutura(cod1, sobra1)
                            for dados2 in tabela2:
                                cod2, descr2, ref2, um2, qtde2, tipo2, saldo2 = dados2

                                folha11 = (cod2, descr2, ref2, um2, qtde2, tipo2, saldo2)
                                tabela_codigus.append(folha11)

                                qtde_float2 = float(qtde2)
                                saldo_float2 = float(saldo2)

                                sobra2 = qtde_float2 - saldo_float2

                                if sobra2 > 0:
                                    tabela3 = self.sql_estrutura(cod2, sobra2)
                                    for dados3 in tabela3:
                                        cod3, descr3, ref3, um3, qtde3, tipo3, saldo3 = dados3

                                        folha22 = (cod3, descr3, ref3, um3, qtde3, tipo3, saldo3)
                                        tabela_codigus.append(folha22)

                                        qtde_float3 = float(qtde3)
                                        saldo_float3 = float(saldo3)

                                        sobra3 = qtde_float3 - saldo_float3

                                        if sobra3 > 0:

                                            tabela4 = self.sql_estrutura(cod3, sobra3)
                                            for dados4 in tabela4:
                                                cod4, descr4, ref4, um4, qtde4, tipo4, saldo4 = dados4

                                                folha33 = (cod4, descr4, ref4, um4, qtde4, tipo4, saldo4)
                                                tabela_codigus.append(folha33)

                                                qtde_float4 = float(qtde4)
                                                saldo_float4 = float(saldo4)

                                                sobra4 = qtde_float4 - saldo_float4

                                                if sobra4 > 0:

                                                    tabela5 = self.sql_estrutura(cod4, sobra4)
                                                    for dados5 in tabela5:
                                                        cod5, descr5, ref5, um5, qtde5, tipo5, saldo5 = dados5

                                                        folha44 = (cod5, descr5, ref5, um5, qtde5, tipo5, saldo5)
                                                        tabela_codigus.append(folha44)

            return tabela_codigus

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calculo_sem_saldo_acabado(self, codigo, quantidade):
        try:
            tabela_codigus = []

            tabela = self.sql_estrutura(codigo, quantidade)

            for dados in tabela:
                cod, descr, ref, um, qtde, tipo, saldo = dados

                folha00 = (cod, descr, ref, um, qtde, tipo, saldo)
                tabela_codigus.append(folha00)

                tabela1 = self.sql_estrutura(cod, qtde)
                for dados1 in tabela1:
                    cod1, descr1, ref1, um1, qtde1, tipo1, saldo1 = dados1

                    folha_ = (cod1, descr1, ref1, um1, qtde1, tipo1, saldo1)
                    tabela_codigus.append(folha_)

                    tabela2 = self.sql_estrutura(cod1, qtde1)
                    for dados2 in tabela2:
                        cod2, descr2, ref2, um2, qtde2, tipo2, saldo2 = dados2

                        folha11 = (cod2, descr2, ref2, um2, qtde2, tipo2, saldo2)
                        tabela_codigus.append(folha11)

                        tabela3 = self.sql_estrutura(cod2, qtde2)
                        for dados3 in tabela3:
                            cod3, descr3, ref3, um3, qtde3, tipo3, saldo3 = dados3

                            folha22 = (cod3, descr3, ref3, um3, qtde3, tipo3, saldo3)
                            tabela_codigus.append(folha22)

                            tabela4 = self.sql_estrutura(cod3, qtde3)
                            for dados4 in tabela4:
                                cod4, descr4, ref4, um4, qtde4, tipo4, saldo4 = dados4

                                folha33 = (cod4, descr4, ref4, um4, qtde4, tipo4, saldo4)
                                tabela_codigus.append(folha33)

                                tabela5 = self.sql_estrutura(cod4, qtde4)
                                for dados5 in tabela5:
                                    cod5, descr5, ref5, um5, qtde5, tipo5, saldo5 = dados5

                                    folha44 = (cod5, descr5, ref5, um5, qtde5, tipo5, saldo5)
                                    tabela_codigus.append(folha44)

            return tabela_codigus

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def consulta_dados(self):
        try:
            codigo_item = self.line_Codigo.text()

            qtdezinha = self.line_Qtde.text()

            if "," in qtdezinha:
                qtdezinha_com_ponto = qtdezinha.replace(',', '.')
                quantidade_item = float(qtdezinha_com_ponto)
            else:
                quantidade_item = float(qtdezinha)

            if self.checkBox_Saldo_Acabado.isChecked():
                tabela_codigos = self.calculo_saldo_acabado(codigo_item, quantidade_item)
            else:
                tabela_codigos = self.calculo_sem_saldo_acabado(codigo_item, quantidade_item)

            for teta in tabela_codigos:
                cod_f, descr_f, ref_f, um_f, qtde_f, tipo_f, saldo_f = teta

                qtde_item_repetido = sum(map(lambda lista_n: lista_n.count(cod_f), tabela_codigos))
                if qtde_item_repetido > 1:

                    list_item_encontrado = [s for s in tabela_codigos if cod_f in s]

                    soma = 0.00
                    for item_encontrado in list_item_encontrado:
                        cod_e, descr_e, ref_e, um_e, qtde_e, tipo_e, saldo_e = item_encontrado
                        soma = float(qtde_e) + soma

                        indice_encon = tabela_codigos.index(item_encontrado)
                        del tabela_codigos[indice_encon]
                    soma_duas_casas = "{:.2f}".format(soma)
                    masquecoisa = (cod_f, descr_f, ref_f, um_f, soma_duas_casas, tipo_f, saldo_f)
                    tabela_codigos.append(masquecoisa)

            return tabela_codigos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def classifica_inicio(self):
        try:
            tabela_codigos = self.consulta_dados()

            extrai_tabela = list(set(tabela_codigos))

            if self.checkBox_Crescente.isChecked():
                decrescente = False
            else:
                decrescente = True

            origem = self.combo_Classifica.currentText()
            if origem == "CÓD.":
                tabela_nova = sorted(extrai_tabela, key=lambda unica: unica[0], reverse=decrescente)
            elif origem == "DESCRIÇÃO":
                tabela_nova = sorted(extrai_tabela, key=lambda unica: unica[1], reverse=decrescente)
            elif origem == "REFERÊNCIA":
                tabela_nova = sorted(extrai_tabela, key=lambda unica: unica[2], reverse=decrescente)
            elif origem == "UM":
                tabela_nova = sorted(extrai_tabela, key=lambda unica: unica[3], reverse=decrescente)
            elif origem == "QTDE":
                tabela_nova = sorted(extrai_tabela, key=lambda unica: unica[4], reverse=decrescente)
            elif origem == "TIPO DE MATERIAL":
                tabela_nova = sorted(extrai_tabela, key=lambda unica: unica[5], reverse=decrescente)
            else:
                tabela_nova = sorted(extrai_tabela, key=lambda unica: unica[6], reverse=decrescente)

            self.procura_ops(tabela_nova)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def procura_ops(self, tabela_nova):
        try:
            tabela_nova1 = []
            tabela_op_exclui = []
            tabela_op_altera = []

            if self.checkBox_OP_Pend.isChecked():
                for dados in tabela_nova:
                    cod, descr, ref, um, qtde, conj, saldo = dados

                    cursor = conecta.cursor()
                    cursor.execute(f"select id, numero, quantidade from ordemservico "
                                   f"where codigo = {cod} and status = 'A';")
                    select_numero = cursor.fetchall()
                    soma_qtde = 0
                    if select_numero:
                        for dados_op in select_numero:
                            ide_op, num_op, qtde_op = dados_op
                            soma_qtde = soma_qtde + qtde_op

                        saldo_atualizado = soma_qtde + saldo

                        if saldo_atualizado < qtde:
                            tutu2 = (cod, descr, ref, um, qtde, conj, saldo_atualizado)
                            tabela_nova1.append(tutu2)

                            deduzinho = (cod, (qtde - saldo_atualizado))
                            tabela_op_altera.append(deduzinho)
                        else:
                            dadinhos = (cod, soma_qtde)
                            tabela_op_exclui.append(dadinhos)
                    else:
                        tutuzin = (cod, descr, ref, um, qtde, conj, saldo)
                        tabela_nova1.append(tutuzin)

            else:
                tabela_nova1 = tabela_nova

            if tabela_op_exclui:
                for itens_com_op in tabela_op_exclui:
                    codigozinho, qtdezinha = itens_com_op

                    tabela = self.sql_estrutura(codigozinho, qtdezinha)
                    for materia_prima in tabela:
                        cod_mat, descr_mat, ref_mat, um_mat, qtde_mat, conj_mat, saldo_mat = materia_prima

                        list_item_encontrado = [s for s in tabela_nova1 if cod_mat in s]

                        soma = 0.00
                        for item_encontrado in list_item_encontrado:
                            cod_e, descr_e, ref_e, um_e, qtde_e, tipo_e, saldo_e = item_encontrado
                            soma = float(qtde_e) + soma

                            indice_encon = tabela_nova1.index(item_encontrado)
                            del tabela_nova1[indice_encon]

            if tabela_op_altera:
                for itens_op_altera in tabela_op_altera:
                    cod_alt, qtd_alt = itens_op_altera

                    tabela_alt = self.sql_estrutura(cod_alt, qtd_alt)
                    for mat_pr_alt in tabela_alt:
                        cod_m_alt, des_m_alt, ref_m_alt, um_m_alt, qtde_m_alt, con_m_alt, sal_m_alt = mat_pr_alt

                        list_item_encontrado = [s for s in tabela_nova1 if cod_m_alt in s]

                        soma = 0.00
                        for item_encontrado in list_item_encontrado:
                            cod_en, descr_en, ref_en, um_en, qtde_en, tipo_en, saldo_en = item_encontrado
                            soma = float(qtde_en) + soma

                            indice_encon = tabela_nova1.index(item_encontrado)
                            tabela_nova1[indice_encon] = cod_en, descr_en, ref_en, um_en, qtde_m_alt, tipo_en, saldo_en

            self.classifica_tipo(tabela_nova1)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def classifica_tipo(self, tabela_nova1):
        try:
            if self.checkBox_Comprados.isChecked():
                self.combo_Tipo.setCurrentIndex(0)

                origem = self.combo_Tipo.currentText()
                origem_nome = str(origem.split("- ")[-1])
                origem_id = int(origem.split(" -")[0])

                self.combo_Tipo.setEnabled(False)
            else:
                self.combo_Tipo.setEnabled(True)
                origem = self.combo_Tipo.currentText()
                origem_nome = str(origem.split("- ")[-1])
                origem_id = int(origem.split(" -")[0])

            tabela_nova2 = []
            soma_itens = 0

            if origem_nome == "TODOS":
                tabela_nova2 = tabela_nova1
                soma_itens = 1
            else:
                for dados in tabela_nova1:
                    cod1, descr1, ref1, um1, qtde1, tipo1, saldo1 = dados

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT tipomaterial FROM produto "
                                   f"where tipomaterial = {origem_id} and codigo = {cod1};")
                    detalhes_tipos1 = cursor.fetchall()
                    if detalhes_tipos1:
                        soma_itens = soma_itens + 1

                        dadus1 = (cod1, descr1, ref1, um1, qtde1, tipo1, saldo1)
                        tabela_nova2.append(dadus1)
            if soma_itens == 0:
                self.mensagem_alerta(f'Na lista de produtos não existe esse tipo de material!')
                self.combo_Tipo.setCurrentIndex(0)
                self.verifica_line_qtde()
            else:
                self.classifica_mat_comprados(tabela_nova2)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def classifica_mat_comprados(self, tabela_nova2):
        try:
            tabela_nova3 = []
            soma_itens = 0

            if self.checkBox_Comprados.isChecked():
                for dados in tabela_nova2:
                    cod2, descr2, ref2, um2, qtde2, tipo2, saldo2 = dados

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT conjunto FROM produto "
                                   f"where conjunto != '10' and codigo = {cod2};")
                    detalhes_tipos2 = cursor.fetchall()

                    if detalhes_tipos2:
                        soma_itens = soma_itens + 1

                        dadus2 = (cod2, descr2, ref2, um2, qtde2, tipo2, saldo2)
                        tabela_nova3.append(dadus2)
            else:
                tabela_nova3 = tabela_nova2

            self.procura_compras(tabela_nova3)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def procura_compras(self, tabela_nova3):
        try:
            tabela_nova4 = []

            if self.checkBox_Compra_Pend.isChecked():
                for dados in tabela_nova3:
                    cod2, descr2, ref2, um2, qtde2, tipo2, saldo2 = dados

                    soma_qtde = 0

                    cur = conecta.cursor()
                    cur.execute(f"SELECT id, codigo, descricao, unidade "
                                f"FROM produto where codigo = {cod2};")
                    detalhes_produto = cur.fetchall()
                    idez, codigo, descr, um = detalhes_produto[0]

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT mestre, produto, quantidade, data "
                                   f"from produtoordemsolicitacao "
                                   f"where produto = {idez} and status = 'A';")
                    select_solicitacao = cursor.fetchall()

                    if select_solicitacao:
                        for dado_sol in select_solicitacao:
                            mestre_sol, produto_sol, qtde_sol, data_sol = dado_sol
                            soma_qtde = soma_qtde + qtde_sol

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT numero, produto, quantidade, data "
                                   f"from produtoordemrequisicao "
                                   f"where produto = {idez} and status = 'A';")
                    select_requisicao = cursor.fetchall()

                    if select_requisicao:
                        for dado_req in select_requisicao:
                            mestre_req, produto_req, qtde_req, data_req = dado_req
                            soma_qtde = soma_qtde + qtde_req

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT prodoc.numero, prodoc.produto, prodoc.quantidade, oc.data "
                                   f"from produtoordemcompra as prodoc "
                                   f"INNER JOIN ordemcompra as oc ON prodoc.mestre = oc.id "
                                   f"where prodoc.produto = {idez} and (prodoc.quantidade - prodoc.produzido) != '0' "
                                   f"and oc.status = 'A' and oc.entradasaida = 'E' and oc.data > '01-01-2021';")
                    select_oc = cursor.fetchall()

                    if select_oc:
                        for dado_oc in select_oc:
                            mestre_oc, produto_oc, qtde_oc, data_oc = dado_oc
                            soma_qtde = soma_qtde + qtde_oc

                    saldo_atualizado = saldo2 + soma_qtde

                    tutu2 = (cod2, descr2, ref2, um2, qtde2, tipo2, saldo_atualizado)
                    tabela_nova4.append(tutu2)
            else:
                tabela_nova4 = tabela_nova3

            self.classifica_sem_saldo(tabela_nova4)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def classifica_sem_saldo(self, tabela_nova4):
        try:
            tabela_nova5 = []

            if self.checkBox_Sem_Saldo.isChecked():
                testinho = 0

                for dados in tabela_nova4:
                    cod, descr, ref, um, qtde, conj, saldo = dados

                    qtde_float = float(qtde)
                    saldo_float = float(saldo)

                    if saldo_float < qtde_float:
                        tutu = (cod, descr, ref, um, qtde, conj, saldo)
                        tabela_nova5.append(tutu)
                    else:
                        testinho = testinho + 1
            else:
                tabela_nova5 = tabela_nova4

            if not tabela_nova5:
                self.mensagem_alerta(f'Com essa classificação, não existem materiais sem saldo!')
                tabela_nova5 = tabela_nova4

            lanca_tabela(self.table_Estrutura, tabela_nova5)
            self.pintar_tabela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    pcpprod = TelaPcpPlano()
    pcpprod.show()
    qt.exec_()

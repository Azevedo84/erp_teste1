import sys
from banco_dados.conexao import conecta
from forms.tela_op_incluir import *
from banco_dados.controle_erros import grava_erro_banco
from banco_dados.bc_consultas import definir_proximo_registro
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.lines import definir_data_atual
from comandos.cores import cor_branco, cor_vermelho
from comandos.telas import tamanho_aplicacao, icone
from PyQt5.QtWidgets import QApplication, QMainWindow, QShortcut, QMessageBox
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtGui import QKeySequence
from PyQt5.QtCore import Qt
from datetime import date, datetime, timedelta
from unidecode import unidecode
import inspect
import os
import traceback


class TelaOpIncluir(QMainWindow, Ui_MainWindow):
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

        self.line_Codigo.returnPressed.connect(lambda: self.verifica_line_codigo())

        self.btn_Consulta_Estrut.clicked.connect(self.verifica_line_qtde)
        self.line_Qtde.returnPressed.connect(lambda: self.verifica_line_qtde())

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)

        self.btn_Limpar.clicked.connect(self.reiniciando_tela)

        self.configura_label()
        definir_data_atual(self.date_Emissao)
        definir_proximo_registro(self.line_Num_OP, "numero", "ordemservico")
        self.lanca_data_previsao()
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
            validator = QtGui.QIntValidator(0, 1234567, self.line_Codigo)
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

    def lanca_data_previsao(self):
        try:
            data_hoje = date.today()
            data_previsao = data_hoje + timedelta(weeks=4)

            self.date_Previsao.setDate(data_previsao)

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

            cursor = conecta.cursor()
            cursor.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade, localizacao, quantidade "
                           f"FROM produto where codigo = {codigo_produto} AND conjunto = 10;")
            produto_acabado = cursor.fetchall()

            if not detalhes_produto:
                self.mensagem_alerta('Este código de produto não existe!')
                self.line_Codigo.clear()
            elif not produto_acabado:
                self.mensagem_alerta('Este código não está classificado como "Produto Acabado"!')
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
                    self.lanca_estrutura()
                    self.pintar_tabela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_estrutura(self):
        try:
            codigo_produto = self.line_Codigo.text()
            qtdezinha = self.line_Qtde.text()

            if "," in qtdezinha:
                qtdezinha_com_ponto = qtdezinha.replace(',', '.')
                qtdezinha_float = float(qtdezinha_com_ponto)
            else:
                qtdezinha_float = float(qtdezinha)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo FROM produto where codigo = {codigo_produto};")
            select_prod = cursor.fetchall()

            idez, cod = select_prod[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT mat.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"conj.conjunto, prod.unidade, "
                           f"(mat.quantidade * {qtdezinha_float}) as qtde, prod.localizacao, "
                           f"prod.quantidade "
                           f"from materiaprima as mat "
                           f"INNER JOIN produto prod ON mat.codigo = prod.codigo "
                           f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                           f"where mat.mestre = {idez} order by conj.conjunto DESC, prod.descricao ASC;")
            tabela_estrutura = cursor.fetchall()

            if not tabela_estrutura:
                self.mensagem_alerta(f'Este produto não possui estrutura cadastrada!\n'
                                                            f'Antes de criar a Ordem de Produção, defina a estrutura.')
                self.reiniciando_tela()

            else:
                if tabela_estrutura:
                    lanca_tabela(self.table_Estrutura, tabela_estrutura)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def pintar_tabela(self):
        try:
            extrai_tabela = extrair_tabela(self.table_Estrutura)

            for index, itens in enumerate(extrai_tabela):
                cod, descr, ref, conj, um, qtde, loca, saldo = itens
                qtde_float = float(qtde)
                saldo_float = float(saldo)

                if saldo_float < qtde_float:
                    font = QFont()
                    font.setBold(True)

                    self.table_Estrutura.item(index, 7).setBackground(QColor(cor_vermelho))
                    self.table_Estrutura.item(index, 7).setFont(font)
                    self.table_Estrutura.item(index, 7).setForeground(QColor(cor_branco))

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def reiniciando_tela(self):
        try:
            self.line_Codigo.clear()
            self.line_Descricao.clear()
            self.line_Compl.clear()
            self.line_Ref.clear()
            self.line_UM.clear()
            self.line_Saldo.clear()
            self.line_Conjunto.clear()
            self.line_Qtde.clear()
            self.line_Obs.clear()

            self.table_Estrutura.setRowCount(0)

            self.configura_label()
            definir_proximo_registro(self.line_Num_OP, "numero", "ordemservico")
            definir_data_atual(self.date_Emissao)
            self.lanca_data_previsao()
            self.line_Codigo.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_salvamento(self):
        try:
            num_op = self.line_Num_OP.text()

            data_hoje = date.today()
            previsao = self.date_Previsao.text()
            prev = datetime.strptime(previsao, '%d/%m/%Y').date()

            cursor = conecta.cursor()
            cursor.execute(f"select id, numero from ordemservico where numero = {num_op};")
            select_op = cursor.fetchall()

            if not num_op:
                self.mensagem_alerta('O campo "Nº OP" não pode estar vazio!')
                self.line_Num_OP.setFocus()
            elif num_op == "0":
                self.mensagem_alerta('O campo "Nº OP" não pode ser "0"!')
                self.line_Num_OP.clear()
                self.line_Num_OP.setFocus()
            elif prev <= data_hoje:
                self.mensagem_alerta('O Data de Previsão deve ser maior que a data atual!')
                self.date_Previsao.setFocus()
            elif select_op:
                self.mensagem_alerta('Este número de "OP" já existe!')
                self.reiniciando_tela()
            else:
                cod_prod = self.line_Codigo.text()

                if not cod_prod:
                    self.mensagem_alerta('O campo "Código" não pode estar vazio!')
                    self.line_Codigo.setFocus()
                else:
                    cur = conecta.cursor()
                    cur.execute(f"SELECT id, descricao, COALESCE(obs, ' ') as obs, unidade "
                                f"FROM produto where codigo = {cod_prod};")
                    detalhes_produto = cur.fetchall()

                    if cod_prod == "0":
                        self.mensagem_alerta('O campo "Código" não pode ser "0"!')
                        self.line_Codigo.clear()
                        self.line_Codigo.setFocus()
                    elif not detalhes_produto:
                        self.mensagem_alerta('Este código de produto não existe!')
                        self.line_Codigo.clear()
                    else:
                        qtde = self.line_Qtde.text()

                        if len(qtde) == 0:
                            self.mensagem_alerta('O campo "Qtde:" não pode estar vazio')
                            self.line_Qtde.clear()
                            self.line_Qtde.setFocus()
                        elif qtde == "0":
                            self.mensagem_alerta('O campo "Qtde:" não pode ser "0"')
                            self.line_Qtde.clear()
                            self.line_Qtde.setFocus()
                        else:
                            self.salvar_dados()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_dados(self):
        try:
            emissao = self.date_Emissao.text()
            emissao_mov = datetime.strptime(emissao, '%d/%m/%Y').date()
            emissao_certo = str(emissao_mov)

            ano_text = emissao_mov.strftime("%Y")
            mes_text = emissao_mov.strftime("%m")
            dia_text = emissao_mov.strftime("%d")

            data_hoje = date.today()
            data_hoje_str = '{}/{}/{}'.format(data_hoje.day, data_hoje.month, data_hoje.year)
            ano_atual = data_hoje.strftime("%Y")
            mes_atual = data_hoje.strftime("%m")
            dia_atual = data_hoje.strftime("%d")

            if ano_text != ano_atual:
                self.mensagem_alerta(f'Você está lançando o consumo deste '
                                                            f'item no ano de {ano_text}!\n\n'
                                                            f'Data Atual: {data_hoje_str}')

            if mes_text != mes_atual:
                self.mensagem_alerta(f'Você está lançando o consumo deste '
                                                            f'item no mês {mes_text}!\n\n'
                                                            f'Data Atual: {data_hoje_str}')

            if dia_text != dia_atual:
                self.mensagem_alerta(f'Você está lançando o consumo deste '
                                                            f'item no dia {dia_text}!\n\n'
                                                            f'Data Atual: {data_hoje_str}')

            prev = self.date_Previsao.text()
            prev_mov = datetime.strptime(prev, '%d/%m/%Y').date()
            previsao = str(prev_mov)

            num_op = self.line_Num_OP.text()
            num_op_int = int(num_op)

            cod_barras = "SUZ000" + num_op

            cod_prod = self.line_Codigo.text()

            cur = conecta.cursor()
            cur.execute(f"SELECT id, descricao, COALESCE(obs, ' ') as obs, unidade "
                        f"FROM produto where codigo = {cod_prod};")
            detalhes_produto = cur.fetchall()
            id_prod, descricao_id, referencia_id, unidade_id = detalhes_produto[0]
            id_prod_int = int(id_prod)

            qtde = self.line_Qtde.text()
            if "," in qtde:
                qtdezinha_com_ponto = qtde.replace(',', '.')
                qtdezinha_float = float(qtdezinha_com_ponto)
            else:
                qtdezinha_float = float(qtde)

            obs = self.line_Obs.text()
            if not obs:
                obs_certo = ""
            else:
                obs_maiuscula = obs.upper()
                obs_certo = unidecode(obs_maiuscula)

            cursor = conecta.cursor()
            cursor.execute(f"Insert into ordemservico "
                           f"(id, produto, numero, quantidade, datainicial, obs, codbarras, status, codigo, "
                           f"dataprevisao) "
                           f"values (GEN_ID(GEN_ORDEMSERVICO_ID,1), {id_prod_int}, {num_op_int}, "
                           f"'{qtdezinha_float}', '{emissao_certo}', '{obs_certo}', '{cod_barras}', 'A', "
                           f"'{cod_prod}', '{previsao}');")
            conecta.commit()

            self.mensagem_alerta(f'A Ordem de Produção Nº {num_op} foi criado com sucesso!')

            self.reiniciando_tela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    opinclui = TelaOpIncluir()
    opinclui.show()
    qt.exec_()

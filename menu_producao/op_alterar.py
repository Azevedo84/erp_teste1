import sys
from banco_dados.conexao import conecta
from forms.tela_op_alterar import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import lanca_tabela, layout_cabec_tab
from comandos.telas import tamanho_aplicacao, icone
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from unidecode import unidecode
from datetime import date
import inspect
import os
import traceback


class TelaOpAlterar(QMainWindow, Ui_AlteraOP):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_producao.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_OP)

        self.btn_Consulta_num.clicked.connect(self.verifica_line_id)
        self.line_Num_OP.returnPressed.connect(lambda: self.verifica_line_id())

        self.btn_Consulta.clicked.connect(self.verifica_line_produto)
        self.line_Codigo.returnPressed.connect(lambda: self.verifica_line_produto())

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)
        self.btn_Limpar.clicked.connect(self.reiniciando_tela)

        self.define_tabela()
        self.line_Num_OP.setFocus()

        data_hoje = date.today()
        self.date_Emissao.setDate(data_hoje)
        
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

    def define_tabela(self):
        try:
            op_ab_editado = []

            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.numero, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, ordser.quantidade, ordser.obs "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where ordser.status = 'A' order by ordser.numero;")
            op_abertas = cursor.fetchall()

            for dados_op in op_abertas:
                emissao, op, cod, descr, ref, um, qtde, obs = dados_op

                data_em_texto = '{}/{}/{}'.format(emissao.day, emissao.month, emissao.year)

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, codigo FROM produto where codigo = {cod};")
                select_prod = cursor.fetchall()

                idez, cod = select_prod[0]

                cursor = conecta.cursor()
                cursor.execute(f"select count(id) from materiaprima where mestre = {idez};")
                total_itens_estrutura = cursor.fetchall()
                itens_estrut_limpa = total_itens_estrutura[0]
                itens_estrut = itens_estrut_limpa[0]

                cursor = conecta.cursor()
                cursor.execute(f"select count(id) from produtoos where numero = {op};")
                total_itens_consumo = cursor.fetchall()
                itens_consumo_limpa = total_itens_consumo[0]
                itens_consumo = itens_consumo_limpa[0]

                est_con = f"{itens_estrut}/{itens_consumo}"

                dados = (data_em_texto, op, cod, descr, ref, um, qtde, est_con, obs)
                op_ab_editado.append(dados)

            lanca_tabela(self.table_OP, op_ab_editado)
        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_id(self):
        try:
            num_op = self.line_Num_OP.text()

            if len(num_op) == 0:
                self.mensagem_alerta('O campo "Nº OP" não pode estar vazio')
                self.line_Num_OP.clear()
                self.line_Num_OP.setFocus()

            elif int(num_op) == 0:
                self.mensagem_alerta('O campo "Nº OP" não pode ser "0"')
                self.line_Num_OP.clear()
                self.line_Num_OP.setFocus()
            else:
                self.verifica_consumo_op()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_consumo_op(self):
        try:
            num_op = self.line_Num_OP.text()
            if len(num_op) == 0:
                self.mensagem_alerta('O campo "Nº OP" não pode estar vazio')
                self.line_Num_OP.clear()
                self.line_Num_OP.setFocus()

            elif int(num_op) == 0:
                self.mensagem_alerta('O campo "Nº OP" não pode ser "0"')
                self.line_Num_OP.clear()
                self.line_Num_OP.setFocus()
            else:
                self.verifica_sql_id()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_sql_id(self):
        try:
            num_op = self.line_Num_OP.text()

            cursor = conecta.cursor()
            cursor.execute(f"select ordser.datainicial, ordser.numero, prod.codigo, prod.descricao, "
                           f"COALESCE(prod.obs, '') as obs, prod.unidade, ordser.quantidade "
                           f"from ordemservico as ordser "
                           f"INNER JOIN produto prod ON ordser.produto = prod.id "
                           f"where ordser.numero = {num_op};")
            select_pela_op = cursor.fetchall()

            if not select_pela_op:
                self.mensagem_alerta('Este Nº de Ordem de Produção não existe!')
                self.line_Num_OP.clear()
                self.line_Num_OP.setFocus()
            else:
                cursor = conecta.cursor()
                cursor.execute(f"select ordser.datainicial, prod.codigo, prod.descricao, prod.descricaocomplementar, "
                               f"COALESCE(prod.obs, '') as obs, prod.unidade, ordser.quantidade, ordser.obs "
                               f"from ordemservico as ordser "
                               f"INNER JOIN produto prod ON ordser.produto = prod.id "
                               f"where ordser.status = 'A' and ordser.numero = {num_op};")
                select_op_aberta = cursor.fetchall()

                cursor = conecta.cursor()
                cursor.execute(f"select count(id) from produtoos where numero = {num_op};")
                total_itens_consumo = cursor.fetchall()
                itens_consumo_limpa = total_itens_consumo[0]
                itens_consumo = itens_consumo_limpa[0]

                if not select_op_aberta:
                    self.mensagem_alerta('Este Nº de Ordem de Produção está encerrado!')
                    self.line_Num_OP.clear()
                    self.line_Num_OP.setFocus()

                elif itens_consumo != 0:
                    self.mensagem_alerta(f'A Ordem de Produção Nº {num_op} já tem '
                                                                f'consumo de materiais!\n'
                                                                f'Exclua os itens consumidos '
                                                                f'para conseguir alterar a "OP"')
                    self.line_Num_OP.clear()
                    self.line_Num_OP.setFocus()

                else:
                    emissao, codigo, descricao, compl, ref, um, qtde, obs = select_op_aberta[0]

                    self.date_Emissao.setDate(emissao)
                    self.date_Emissao.setReadOnly(True)

                    self.line_Codigo.setText(codigo)

                    self.line_Descricao.setText(descricao)
                    self.line_Descricao.setReadOnly(True)

                    self.line_Complementar.setText(compl)
                    self.line_Complementar.setReadOnly(True)

                    self.line_Referencia.setText(ref)
                    self.line_Referencia.setReadOnly(True)

                    self.line_Um.setText(um)
                    self.line_Um.setReadOnly(True)

                    qtde_str = str(qtde)
                    self.line_Qtde.setText(qtde_str)

                    self.line_Obs.setText(obs)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_line_produto(self):
        try:
            codigo_produto = self.line_Codigo.text()
            if len(codigo_produto) == 0:
                self.mensagem_alerta('O campo "Código" não pode estar vazio')
                self.line_Codigo.clear()
                self.line_Descricao.clear()
                self.line_Referencia.clear()
                self.line_Um.clear()
            elif int(codigo_produto) == 0:
                self.mensagem_alerta('O campo "Código" não pode ser "0"')
                self.line_Codigo.clear()
                self.line_Descricao.clear()
                self.line_Referencia.clear()
                self.line_Um.clear()
            else:
                self.verifica_sql_produto()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_sql_produto(self):
        try:
            codigo_produto = self.line_Codigo.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade, localizacao, quantidade "
                           f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cursor.fetchall()
            if not detalhes_produto:
                self.mensagem_alerta('Este código de produto não existe!')
                self.line_Codigo.clear()
                self.line_Descricao.clear()
                self.line_Referencia.clear()
                self.line_Um.clear()
            else:
                self.verifica_materia_prima()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_materia_prima(self):
        try:
            codigo_produto = self.line_Codigo.text()
            cursor = conecta.cursor()
            cursor.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, unidade, localizacao, quantidade "
                           f"FROM produto where codigo = {codigo_produto} AND conjunto = 10;")
            detalhes_produto = cursor.fetchall()
            if not detalhes_produto:
                self.mensagem_alerta('Este produto não está definido como "Produto Acabado"!')
                self.line_Codigo.clear()
            else:
                self.lanca_dados_produto()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_produto(self):
        try:
            codigo_produto = self.line_Codigo.text()
            cur = conecta.cursor()
            cur.execute(f"SELECT descricao, COALESCE(obs, ' ') as obs, descricaocomplementar, unidade "
                        f"FROM produto where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            descricao, referencia, compl, unidade = detalhes_produto[0]

            self.line_Descricao.setText(descricao)
            self.line_Complementar.setText(compl)
            self.line_Referencia.setText(referencia)
            self.line_Um.setText(unidade)
            self.line_Qtde.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def reiniciando_tela(self):
        try:
            data_hoje = date.today()
            self.date_Emissao.setDate(data_hoje)
            self.date_Emissao.setReadOnly(True)

            self.line_Codigo.clear()
            self.line_Descricao.clear()
            self.line_Referencia.clear()
            self.line_Um.clear()
            self.line_Complementar.clear()
            self.line_Qtde.clear()
            self.line_Obs.clear()

            self.define_tabela()
            self.line_Num_OP.setFocus()
            self.line_Num_OP.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_salvamento(self):
        try:
            num_op = self.line_Num_OP.text()
            qtde = self.line_Qtde.text()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT * from ordemservico where numero = {num_op};")
            tabela_maq = cursor.fetchall()

            self.verifica_line_produto()

            cursor = conecta.cursor()
            cursor.execute(f"select count(id) from produtoos where numero = {num_op};")
            total_itens_consumo = cursor.fetchall()
            itens_consumo_limpa = total_itens_consumo[0]
            itens_consumo = itens_consumo_limpa[0]

            if not tabela_maq:
                self.mensagem_alerta('Este número de Ordem de Produção não existe!')
                self.reiniciando_tela()
            elif not num_op:
                self.mensagem_alerta('O campo "Nº OP" não pode estar vazio!')
                self.line_Codigo.setFocus()
            elif num_op == "0":
                self.mensagem_alerta('O campo "Nº OP" não pode ser "0"!')
                self.line_Codigo.clear()
                self.line_Codigo.setFocus()
            elif len(qtde) == 0:
                self.mensagem_alerta('O campo "Qtde:" não pode estar vazio')
                self.line_Qtde.clear()
                self.line_Qtde.setFocus()
            elif qtde == "0":
                self.mensagem_alerta('O campo "Qtde:" não pode ser "0"')
                self.line_Qtde.clear()
                self.line_Qtde.setFocus()
            elif itens_consumo != 0:
                self.mensagem_alerta(f'A Ordem de Produção Nº {num_op} já tem '
                                                            f'consumo de materiais!\n'
                                                            f'Exclua os itens consumidos '
                                                            f'para conseguir alterar a "OP"')
                self.reiniciando_tela()
            else:
                self.salvar_dados()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_dados(self):
        try:
            num_op = self.line_Num_OP.text()
            num_op_int = int(num_op)

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
            if obs == "":
                obs_certo = ""
            else:
                if "'" in obs:
                    obs_aspas1 = obs.replace("'", " ")
                    if '"' in obs_aspas1:
                        obs_aspas = obs_aspas1.replace('"', ' ')
                    else:
                        obs_aspas = obs_aspas1
                elif '"' in obs:
                    obs_aspas1 = obs.replace('"', ' ')
                    if "'" in obs_aspas1:
                        obs_aspas = obs_aspas1.replace("'", " ")
                    else:
                        obs_aspas = obs_aspas1
                else:
                    obs_aspas = obs

                obs_maiuscula = obs_aspas.upper()
                obs_certo = unidecode(obs_maiuscula)

            cursor = conecta.cursor()
            cursor.execute(f"UPDATE ordemservico SET produto = {id_prod_int}, codigo = '{cod_prod}', "
                           f"quantidade = '{qtdezinha_float}', obs = '{obs_certo}' "
                           f"where numero = {num_op_int};")

            conecta.commit()

            self.mensagem_alerta(f'A Ordem de Produção {num_op} foi alterada com sucesso!')

            self.reiniciando_tela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    opaltera = TelaOpAlterar()
    opaltera.show()
    qt.exec_()

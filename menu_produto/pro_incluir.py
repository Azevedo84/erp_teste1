import sys
from banco_dados.conexao import conectar_banco_nuvem
from comandos.conversores import valores_para_float
from forms.tela_prod_incluir import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.telas import tamanho_aplicacao, icone
from comandos.lines import validador_decimal
from comandos.conversores import data_brasileiro
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import inspect
import os
from datetime import date
import traceback
import re


class TelaProdutoIncluir(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_cadastro.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Produto)

        self.table_Produto.viewport().installEventFilter(self)

        self.btn_Limpar.clicked.connect(self.limpa_tudo)

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)

        validador_decimal(self.line_NCM, numero=9999999.000)

        self.inicio_manipula_pre_cadastro()
        self.lanca_combo_classificacao()
        self.data_emissao()

        self.processando = False
        
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
            
    def limpa_tabela(self):
        try:
            nome_tabela = self.table_Produto

            nome_tabela.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def remover_espaco_branco_ini_fim(self, string):
        try:
            if string.endswith(' '):
                string_final = string.rstrip()
            else:
                string_final = string
    
            if string_final.startswith(' '):
                string_final1 = string_final.lstrip()
            else:
                string_final1 = string_final.lstrip()
    
            return string_final1
        
        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def inicio_manipula_pre_cadastro(self):
        conecta = conectar_banco_nuvem()
        try:
            self.limpa_tabela()

            tabela = []
            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, obs, descricao, complementar, referencia, um, ncm, "
                           f"fornecedor, criacao, produto_id "
                           f"FROM PRODUTO_PRELIMINAR where produto_id IS NULL "
                           f"order by obs;")
            dados_banco = cursor.fetchall()

            if dados_banco:
                for i in dados_banco:
                    id_pre, obs, descr, compl, ref, um, ncm, forn, emissao, codigo = i

                    data_formatada = data_brasileiro(emissao)

                    dados = (data_formatada, id_pre, obs, descr, compl, ref, um, ncm, forn)
                    tabela.append(dados)

            if tabela:
                lanca_tabela(self.table_Produto, tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def data_emissao(self):
        try:
            data_hoje = date.today()
            self.date_Emissao.setDate(data_hoje)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_combo_classificacao(self):
        conecta = conectar_banco_nuvem()
        try:
            self.combo_Classificacao.clear()

            nova_lista = [""]

            cursor = conecta.cursor()
            cursor.execute('SELECT id, descricao FROM PRODUTO_CLASSIFICACAO order by descricao;')
            lista_completa = cursor.fetchall()
            for ides, descr in lista_completa:
                dd = f"{ides} - {descr}"
                nova_lista.append(dd)

            self.combo_Classificacao.addItems(nova_lista)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def eventFilter(self, source, event):
        conecta = conectar_banco_nuvem()
        try:
            qtable_widget = self.table_Produto

            if (event.type() == QtCore.QEvent.MouseButtonDblClick and
                    event.buttons() == QtCore.Qt.LeftButton and
                    source is qtable_widget.viewport()):

                self.label_Fornecedor.setText("")

                self.limpa_dados_produto()

                item = qtable_widget.currentItem()

                extrai_recomendados = extrair_tabela(qtable_widget)
                item_selecionado = extrai_recomendados[item.row()]

                data, id_pre, obs, descr, compl, ref, um, ncm, forn = item_selecionado

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, obs, descricao, complementar, referencia, um, ncm, "
                               f"fornecedor, criacao, produto_id "
                               f"FROM PRODUTO_PRELIMINAR "
                               f"where id = {id_pre};")
                dados_banco = cursor.fetchall()

                if dados_banco:
                    for i in dados_banco:
                        id_pres, obs, descrs, compl, refs, ums, ncm, forns, emissao, cod_prod = i

                        descr_sem = self.remover_espaco_branco_ini_fim(descrs)
                        compl_sem = self.remover_espaco_branco_ini_fim(compl)
                        ref_sem = self.remover_espaco_branco_ini_fim(refs)

                        ja_existe = self.verifica_ref_desenho_existe(ref_sem)
                        if ja_existe:
                            msg = ""
                            for ii in ja_existe:
                                cod_des, descr_des, ref_des = ii
                                msg += f"{cod_des} - {descr_des} - {ref_des}\n"
                            self.mensagem_alerta(f"Já existe produtos com este número de desenho!\n\n{msg}")
                        else:
                            if forns:
                                self.label_Fornecedor.setText(forns)

                            self.manipula_ref_desenho(ref_sem)

                            self.line_ID_Pre.setText(str(id_pres))
                            self.line_Descricao.setText(descr_sem)
                            self.line_DescrCompl.setText(compl_sem)
                            self.line_Referencia.setText(ref_sem)
                            self.line_NCM.setText(ncm)
                            self.line_Qtde_Mini.setText("0")

                            um_count = self.combo_UM.count()
                            for i_um in range(um_count):
                                um_text = self.combo_UM.itemText(i_um)
                                if ums in um_text:
                                    self.combo_UM.setCurrentText(um_text)

                            self.line_Codigo.setFocus()

            return super(QMainWindow, self).eventFilter(source, event)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def manipula_ref_desenho(self, refs):
        conecta = conectar_banco_nuvem()
        try:
            if self.verifica_formato_referencia(refs):
                ref_sem_d = refs[2:]
                cod_maq = int(ref_sem_d[0:2])

                tp = int(ref_sem_d[-2:])
                conj_peca = (ref_sem_d[3:5])

                if tp == 1 or tp == 3 or tp == 4 or tp == 5 or tp == 6 or tp == 7:
                    conj_count = self.combo_Classificacao.count()
                    for i_conj in range(conj_count):
                        conj_text = self.combo_Classificacao.itemText(i_conj)
                        if "10 - PRODUTOS ACABADOS" in conj_text:
                            self.combo_Classificacao.setCurrentText(conj_text)
                else:
                    conj_count = self.combo_Classificacao.count()
                    for i_conj in range(conj_count):
                        conj_text = self.combo_Classificacao.itemText(i_conj)
                        if "8 - MATERIA PRIMA" in conj_text:
                            self.combo_Classificacao.setCurrentText(conj_text)

            else:
                conj_count = self.combo_Classificacao.count()
                for i_conj in range(conj_count):
                    conj_text = self.combo_Classificacao.itemText(i_conj)
                    if "8 - MATERIA PRIMA" in conj_text:
                        self.combo_Classificacao.setCurrentText(conj_text)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def verifica_ref_desenho_existe(self, refs):
        conecta = conectar_banco_nuvem()
        try:
            ja_existe = []

            if self.verifica_formato_referencia(refs):
                cursor = conecta.cursor()
                cursor.execute(f"SELECT id_siger, descricao, obs FROM PRODUTO where obs = '{refs}';")
                lista_completa = cursor.fetchall()
                if lista_completa:
                    ja_existe = lista_completa

            return ja_existe

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def verifica_formato_referencia(self, referencia):
        try:
            padrao = re.compile(r'^D \d{2}\.\d{2}\.\d{3}\.\d{2}$')
            correspondencia = padrao.match(referencia)

            return correspondencia

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_produto(self):
        try:
            self.line_ID_Pre.clear()
            self.line_Codigo.clear()
            self.line_Referencia.clear()
            self.line_Descricao.clear()
            self.line_DescrCompl.clear()
            self.line_NCM.clear()
            self.line_Qtde_Mini.clear()
            self.plain_Obs.clear()

            self.combo_UM.setCurrentText("")
            self.combo_Classificacao.setCurrentText("")

            self.label_Maquina_Des.setText("")

            self.data_emissao()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tudo(self):
        try:
            self.limpa_dados_produto()
            self.inicio_manipula_pre_cadastro()
            self.lanca_combo_classificacao()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_salvamento(self):
        conecta = conectar_banco_nuvem()
        try:
            cod_produto = self.line_Codigo.text()
            descr = self.line_Descricao.text()
            ncm = self.line_NCM.text()
            um = self.combo_UM.currentText()

            classificacao = self.combo_Classificacao.currentText()

            if not cod_produto:
                self.mensagem_alerta(f'O Código do produto não pode estar vazio!')
            elif not descr:
                self.mensagem_alerta(f'A Descrição do produto não pode estar vazia!')
            elif not ncm:
                self.mensagem_alerta(f'A NCM do produto não pode estar vazia!')
            elif not um:
                self.mensagem_alerta(f'A Unidade de Medida (UM) do produto não pode estar vazia!')
            elif not classificacao:
                self.mensagem_alerta(f'O Conjunto do produto não pode estar vazio!')
            else:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT * FROM PRODUTO where id_siger = '{cod_produto}';")
                lista_completa = cursor.fetchall()
                if lista_completa:
                    self.mensagem_alerta(f'Este código de produto já foi cadastrado!')
                else:
                    self.salvar_produto()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def verifica_pre_cadastro(self):
        conecta = conectar_banco_nuvem()
        try:
            id_pre = self.line_ID_Pre.text()
            cod_produto = self.line_Codigo.text()

            if id_pre:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT * "
                               f"FROM PRODUTO_PRELIMINAR  "
                               f"where id = {id_pre};")
                dados_banco = cursor.fetchall()

                if dados_banco:
                    cursor = conecta.cursor()
                    cursor.execute(f"UPDATE PRODUTO_PRELIMINAR  SET produto_id = {cod_produto} WHERE id = {id_pre};")

                    conecta.commit()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()

    def salvar_produto(self):
        conecta = conectar_banco_nuvem()
        try:
            cod_produto = self.line_Codigo.text()
            ref = self.line_Referencia.text()
            descr = self.line_Descricao.text()
            compl = self.line_DescrCompl.text()

            ncm = self.line_NCM.text()
            ncm_so_numeros = re.sub(r'\D', '', ncm)

            qtde_minima = self.line_Qtde_Mini.text()
            qtde_minima_float = valores_para_float(qtde_minima)

            obs = self.plain_Obs.toPlainText()

            classifica = self.combo_Classificacao.currentText()
            classificatete = classifica.find(" - ")
            id_classifica = classifica[:classificatete]

            um = self.combo_UM.currentText()

            data_hoje = date.today()

            cursor = conecta.cursor()
            cursor.execute(f"Insert into PRODUTO (ID_SIGER, DESCRICAO, COMPLEMENTAR, REFERENCIA, UM, "
                           f"CLASSIFICACAO_PRODUTO_ID, QTDE_MINIMA, NCM, OBS, ATIVO) "
                           f"values ('{cod_produto}', '{descr}', '{compl}', "
                           f"'{ref}', '{um}', {id_classifica}, '{qtde_minima_float}', '{ncm_so_numeros}', "
                           f"'{obs}', 'A');")

            conecta.commit()

            self.mensagem_alerta(f"Cadastro do produto {cod_produto} realizado com Sucesso!")

            self.verifica_pre_cadastro()

            self.limpa_tudo()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

        finally:
            if 'conexao' in locals():
                conecta.close()


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaProdutoIncluir()
    tela.show()
    qt.exec_()

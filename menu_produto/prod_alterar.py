import sys
from banco_dados.conexao import conecta
from forms.tela_prod_alterar import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.telas import icone
from comandos.lines import validador_decimal
from comandos.conversores import valores_para_float, float_para_virgula
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtCore import pyqtSignal, QDate
import inspect
import os
import re
from datetime import date
import traceback


class TelaProdutoAlterar(QMainWindow, Ui_MainWindow):
    alteracao = pyqtSignal(bool)

    def __init__(self, dados_produto, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        self.veio_de_fora = dados_produto

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_cadastro.png")

        self.btn_Salvar.clicked.connect(self.verifica_salvamento)

        self.line_Codigo.editingFinished.connect(self.verifica_line_codigo_manual)

        validador_decimal(self.line_NCM, numero=9999999.000)

        self.lanca_combo_conjunto()
        self.lanca_combo_tipo()
        self.lanca_combo_projeto()
        self.data_emissao()

        self.processando = False

        self.dados_produto = ()

        if dados_produto:
            self.lanca_dados_produto(dados_produto)
            self.line_Codigo.setReadOnly(True)
            self.line_Descricao.setFocus()

            self.dados_produto = dados_produto

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

    def verifica_line_codigo_manual(self):
        if not self.processando:
            try:
                self.processando = True

                codigo_produto = self.line_Codigo.text()

                if codigo_produto:
                    self.verifica_sql_produto_manual()
                    self.line_Referencia.setFocus()

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
                self.line_Embalagem.setText(embalagem)
                self.line_kg_mt.setText(kg_mt)
                self.line_Custo_Unit.setText(custo_unit)
                self.line_Local.setText(local)
                self.plain_Obs.setPlainText(obs)

                if um:
                    um_count = self.combo_UM.count()
                    for i_um in range(um_count):
                        um_text = self.combo_UM.itemText(i_um)
                        if um in um_text:
                            self.combo_UM.setCurrentText(um_text)

                if conjunto:
                    cursor = conecta.cursor()
                    cursor.execute("SELECT id, conjunto FROM conjuntos where conjunto = ?", (conjunto, ))
                    lista_conj = cursor.fetchall()
                    if lista_conj:
                        id_conj, descr_conj = lista_conj[0]
                        conj_certo = f"{id_conj} - {descr_conj}"
                        conjunto_count = self.combo_Conjunto.count()
                        for i_conjunto in range(conjunto_count):
                            conjunto_text = self.combo_Conjunto.itemText(i_conjunto)
                            if conj_certo in conjunto_text:
                                self.combo_Conjunto.setCurrentText(conjunto_text)

                if tipo:
                    cursor = conecta.cursor()
                    cursor.execute("SELECT id, tipomaterial FROM TIPOMATERIAL WHERE tipomaterial = ?", (tipo,))
                    lista_tipo = cursor.fetchall()

                    if lista_tipo:
                        id_tipo, descr_tipo = lista_tipo[0]
                        tipo_certo = f"{id_tipo} - {descr_tipo}"
                        tipo_count = self.combo_Tipo.count()
                        for i_tipo in range(tipo_count):
                            tipo_text = self.combo_Tipo.itemText(i_tipo)
                            if tipo_certo in tipo_text:
                                self.combo_Tipo.setCurrentText(tipo_text)

                if projeto:
                    cursor = conecta.cursor()
                    cursor.execute("SELECT id, projeto FROM PROJETO where projeto =  ?", (projeto,))
                    lista_projeto = cursor.fetchall()
                    if lista_projeto:
                        id_projeto, descr_projeto = lista_projeto[0]
                        projeto_certo = f"{id_projeto} - {descr_projeto}"
                        projeto_count = self.combo_Projeto.count()
                        for i_projeto in range(projeto_count):
                            projeto_text = self.combo_Projeto.itemText(i_projeto)
                            if projeto_certo in projeto_text:
                                self.combo_Projeto.setCurrentText(projeto_text)

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

                dados = (codigo_produto, data, barras_sem_espacos, descr, compl, ref, um, embalagem, kg_mt,
                         custo, local, conjunto, tipo, projeto, min_str, ncm, obs)

                self.dados_produto = dados

            else:
                self.mensagem_alerta("Este cadastro de produto não existe!")
                self.line_Codigo.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_produto(self, dados):
        try:
            codigo, emissao, barra, descr, compl, ref, um, embalagem, kg_mt, custo, local, conjunto, \
            tipo, projeto, qtde_mini, ncm, obs = dados

            if emissao:
                emissao_date = QDate.fromString(emissao, 'dd/MM/yyyy')
                self.date_Emissao.setDate(emissao_date)

            self.line_Codigo.setText(codigo)
            self.line_Barras.setText(barra)
            self.line_Descricao.setText(descr)
            self.line_DescrCompl.setText(compl)
            self.line_Referencia.setText(ref)
            self.line_Embalagem.setText(embalagem)
            self.line_kg_mt.setText(kg_mt)
            self.line_Custo_Unit.setText(custo)
            self.line_Local.setText(local)
            self.line_Qtde_Mini.setText(qtde_mini)
            self.plain_Obs.setPlainText(obs)

            if not ncm:
                self.line_NCM.setText("")
                self.line_NCM.setStyleSheet("QLineEdit { background-color: yellow; }")
            else:
                self.line_NCM.setText(ncm)
                self.line_NCM.setStyleSheet("QLineEdit { background-color: white; }")

            um_count = self.combo_UM.count()
            for i_um in range(um_count):
                um_text = self.combo_UM.itemText(i_um)
                if um in um_text:
                    self.combo_UM.setCurrentText(um_text)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, conjunto FROM conjuntos where conjunto = '{conjunto}';")
            lista_conj = cursor.fetchall()
            if lista_conj:
                id_conj, descr_conj = lista_conj[0]
                conj_certo = f"{id_conj} - {descr_conj}"
                conjunto_count = self.combo_Conjunto.count()
                for i_conjunto in range(conjunto_count):
                    conjunto_text = self.combo_Conjunto.itemText(i_conjunto)
                    if conj_certo in conjunto_text:
                        self.combo_Conjunto.setCurrentText(conjunto_text)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, tipomaterial FROM TIPOMATERIAL where tipomaterial = '{tipo}';")
            lista_tipo = cursor.fetchall()
            if lista_tipo:
                id_tipo, descr_tipo = lista_tipo[0]
                tipo_certo = f"{id_tipo} - {descr_tipo}"
                tipo_count = self.combo_Tipo.count()
                for i_tipo in range(tipo_count):
                    tipo_text = self.combo_Tipo.itemText(i_tipo)
                    if tipo_certo in tipo_text:
                        self.combo_Tipo.setCurrentText(tipo_text)

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, projeto FROM PROJETO where projeto = '{projeto}';")
            lista_projeto = cursor.fetchall()
            if lista_projeto:
                id_projeto, descr_projeto = lista_projeto[0]
                projeto_certo = f"{id_projeto} - {descr_projeto}"
                projeto_count = self.combo_Projeto.count()
                for i_projeto in range(projeto_count):
                    projeto_text = self.combo_Projeto.itemText(i_projeto)
                    if projeto_certo in projeto_text:
                        self.combo_Projeto.setCurrentText(projeto_text)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_formato_referencia(self, referencia):
        try:
            padrao = re.compile(r'^D \d{2}\.\d{2}\.\d{3}\.\d{2}$')
            correspondencia = padrao.match(referencia)

            return correspondencia

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

    def data_emissao(self):
        try:
            data_hoje = date.today()
            self.date_Emissao.setDate(data_hoje)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_combo_conjunto(self):
        try:
            self.combo_Conjunto.clear()

            nova_lista = [""]

            cursor = conecta.cursor()
            cursor.execute('SELECT id, conjunto FROM conjuntos order by conjunto;')
            lista_completa = cursor.fetchall()
            for ides, descr in lista_completa:
                dd = f"{ides} - {descr}"
                nova_lista.append(dd)

            self.combo_Conjunto.addItems(nova_lista)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_combo_tipo(self):
        try:
            self.combo_Tipo.clear()

            nova_lista = [""]

            cursor = conecta.cursor()
            cursor.execute('SELECT id, tipomaterial FROM TIPOMATERIAL order by tipomaterial;')
            lista_completa = cursor.fetchall()
            for ides, descr in lista_completa:
                dd = f"{ides} - {descr}"
                nova_lista.append(dd)

            self.combo_Tipo.addItems(nova_lista)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_combo_projeto(self):
        try:
            self.combo_Projeto.clear()

            nova_lista = [""]

            cursor = conecta.cursor()
            cursor.execute('SELECT id, projeto FROM PROJETO order by projeto;')
            lista_completa = cursor.fetchall()
            for ides, descr in lista_completa:
                dd = f"{ides} - {descr}"
                nova_lista.append(dd)

            self.combo_Projeto.addItems(nova_lista)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_codigo_barras(self):
        if not self.processando:
            try:
                self.processando = True

                codigo_produto = self.line_Codigo.text()

                if codigo_produto:
                    prefixo = "SZP"
                    total_digitos = 10
                    zeros_de_preenchimento = total_digitos - len(prefixo) - len(str(codigo_produto))
                    codigo_barras = prefixo + "0" * zeros_de_preenchimento + str(codigo_produto)
                    self.line_Barras.setText(codigo_barras)

                    self.line_Custo_Unit.setText("0")
                    self.line_Qtde_Mini.setText("0")

                    if not self.line_kg_mt.text():
                        self.line_kg_mt.setText("0")

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def verifica_ref_desenho_existe(self, refs):
        try:
            ja_existe = []

            if self.verifica_formato_referencia(refs):
                cursor = conecta.cursor()
                cursor.execute(f"SELECT codigo, descricao, obs FROM produto where obs = '{refs}';")
                lista_completa = cursor.fetchall()
                if lista_completa:
                    ja_existe = lista_completa

            return ja_existe

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_descricao_tipo(self, descricao):
        try:
            posicao = descricao.find(" ")
            prim_palavra = descricao[:posicao]

            if prim_palavra == "INVERSOR":
                tip_count = self.combo_Tipo.count()
                for i_tip in range(tip_count):
                    tip_text = self.combo_Tipo.itemText(i_tip)
                    if "98 - INVERSORES" in tip_text:
                        self.combo_Tipo.setCurrentText(tip_text)

            elif prim_palavra == "PARAFUSO":
                tip_count = self.combo_Tipo.count()
                for i_tip in range(tip_count):
                    tip_text = self.combo_Tipo.itemText(i_tip)
                    if "76 - FATI/FG" in tip_text:
                        self.combo_Tipo.setCurrentText(tip_text)

            elif prim_palavra == "ROLAMENTO":
                tip_count = self.combo_Tipo.count()
                for i_tip in range(tip_count):
                    tip_text = self.combo_Tipo.itemText(i_tip)
                    if "78 - JJD" in tip_text:
                        self.combo_Tipo.setCurrentText(tip_text)

            elif prim_palavra == "CONTACTORA":
                tip_count = self.combo_Tipo.count()
                for i_tip in range(tip_count):
                    tip_text = self.combo_Tipo.itemText(i_tip)
                    if "82 - REAL CENTER" in tip_text:
                        self.combo_Tipo.setCurrentText(tip_text)

            elif prim_palavra == "CONTATOR":
                tip_count = self.combo_Tipo.count()
                for i_tip in range(tip_count):
                    tip_text = self.combo_Tipo.itemText(i_tip)
                    if "82 - REAL CENTER" in tip_text:
                        self.combo_Tipo.setCurrentText(tip_text)

            elif prim_palavra == "BOTAO":
                tip_count = self.combo_Tipo.count()
                for i_tip in range(tip_count):
                    tip_text = self.combo_Tipo.itemText(i_tip)
                    if "82 - REAL CENTER" in tip_text:
                        self.combo_Tipo.setCurrentText(tip_text)

            elif prim_palavra == "RELE":
                tip_count = self.combo_Tipo.count()
                for i_tip in range(tip_count):
                    tip_text = self.combo_Tipo.itemText(i_tip)
                    if "82 - REAL CENTER" in tip_text:
                        self.combo_Tipo.setCurrentText(tip_text)

            elif prim_palavra == "RESISTENCIA":
                tip_count = self.combo_Tipo.count()
                for i_tip in range(tip_count):
                    tip_text = self.combo_Tipo.itemText(i_tip)
                    if "109 - RESIMAC" in tip_text:
                        self.combo_Tipo.setCurrentText(tip_text)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_fornecedor_tipo(self, fornecedor):
        try:
            qtde_palavras = len(fornecedor.split())

            if qtde_palavras == 1:
                if fornecedor == "PWM":
                    tip_count = self.combo_Tipo.count()
                    for i_tip in range(tip_count):
                        tip_text = self.combo_Tipo.itemText(i_tip)
                        if "126 - PWM" in tip_text:
                            self.combo_Tipo.setCurrentText(tip_text)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_produto(self):
        try:
            self.line_Codigo.clear()
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

            self.combo_UM.setCurrentText("")
            self.combo_Conjunto.setCurrentText("")
            self.combo_Tipo.setCurrentText("")
            self.combo_Projeto.setCurrentText("")

            self.data_emissao()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tudo(self):
        try:
            self.limpa_dados_produto()
            self.lanca_combo_conjunto()
            self.lanca_combo_tipo()
            self.lanca_combo_projeto()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_salvamento(self):
        try:
            cod_produto = self.line_Codigo.text()
            cod_barras = self.line_Barras.text()
            descr = self.line_Descricao.text()
            ncm = self.line_NCM.text()
            um = self.combo_UM.currentText()

            conjunto = self.combo_Conjunto.currentText()
            tipo = self.combo_Tipo.currentText()

            if not cod_produto:
                self.mensagem_alerta(f'O Código do produto não pode estar vazio!')
            elif not cod_barras:
                self.mensagem_alerta(f'O Código de barras do produto não pode estar vazio!')
            elif not descr:
                self.mensagem_alerta(f'A Descrição do produto não pode estar vazia!')
            elif not um:
                self.mensagem_alerta(f'A Unidade de Medida (UM) do produto não pode estar vazia!')
            elif not conjunto:
                self.mensagem_alerta(f'O Conjunto do produto não pode estar vazio!')
            elif not tipo:
                self.mensagem_alerta(f'O Tipo de Material do produto não pode estar vazio!')
            else:
                if not ncm:
                    msg = f'A NCM do produto não deveria estar vazia!\n\n' \
                          f'Tem certeza que deseja continuar?'
                    if self.pergunta_confirmacao(msg):
                        status = True
                    else:
                        status = False
                else:
                    status = True

                if status:
                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT codigo, descricao, obs FROM produto where codigo = '{cod_produto}';")
                    lista_completa = cursor.fetchall()
                    if lista_completa:
                        if um == "KG":
                            kg_mt = self.line_kg_mt.text()
                            if not kg_mt:
                                self.mensagem_alerta(f'O "KG/MT" do produto não pode estar vazio!')
                            else:
                                self.salvar_alteracao()
                        else:
                            self.salvar_alteracao()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_alteracao(self):
        try:
            if self.dados_produto:
                codigo, emissao, barra, descr, compl, ref, um, embalagem, kg_mt, custo, local, conjunto, \
                tipo, projeto, qtde_mini, ncm, obs = self.dados_produto

                ref_a = self.line_Referencia.text()
                descr_a = self.line_Descricao.text()
                compl_a = self.line_DescrCompl.text()
                embalagem_a = self.line_Embalagem.text()
                ncm_a = self.line_NCM.text()
                kg_mt_a = self.line_kg_mt.text()
                qtde_mini_a = self.line_Qtde_Mini.text()
                custo_a = self.line_Custo_Unit.text()

                local_a = self.line_Local.text()
                obs_plain = self.plain_Obs.toPlainText()
                obs_a = obs_plain.upper()

                conjunt = self.combo_Conjunto.currentText()
                if conjunt:
                    conjuntotete = conjunt.find(" - ") + 3
                    conjunto_a = conjunt[conjuntotete:]
                else:
                    conjunto_a = ""

                tip = self.combo_Tipo.currentText()
                if tip:
                    tipotete = tip.find(" - ") + 3
                    tipo_a = tip[tipotete:]
                else:
                    tipo_a = ""

                um_a = self.combo_UM.currentText()

                projet = self.combo_Projeto.currentText()
                if projet:
                    projetotete = projet.find(" - ") + 3
                    projeto_a = projet[projetotete:]

                else:
                    projeto_a = ""

                kg_mt_float = valores_para_float(kg_mt)
                kg_mt_float_a = valores_para_float(kg_mt_a)

                custo_float = valores_para_float(custo)
                custo_float_a = valores_para_float(custo_a)

                qtde_mini_float = valores_para_float(qtde_mini)
                qtde_mini_float_a = valores_para_float(qtde_mini_a)

                campos_atualizados = []

                if descr != descr_a:
                    campos_atualizados.append(f"DESCRICAO = '{descr_a}'")

                if compl != compl_a:
                    campos_atualizados.append(f"DESCRICAOCOMPLEMENTAR = '{compl_a}'")

                if ref != ref_a:
                    campos_atualizados.append(f"OBS = '{ref_a}'")

                if um != um_a:
                    campos_atualizados.append(f"UNIDADE = '{um_a}'")

                if embalagem != embalagem_a:
                    campos_atualizados.append(f"EMBALAGEM = '{embalagem_a}'")

                if kg_mt_float != kg_mt_float_a:
                    campos_atualizados.append(f"KILOSMETRO = '{kg_mt_float_a}'")

                if custo_float != custo_float_a:
                    campos_atualizados.append(f"CUSTOUNITARIO = '{custo_float_a}'")

                if local != local_a:
                    campos_atualizados.append(f"LOCALIZACAO = '{local_a}'")

                if conjunto != conjunto_a:
                    if conjunto_a:
                        conjuntotete = conjunt.find(" - ")
                        id_conj_a = conjunt[:conjuntotete]
                    else:
                        id_conj_a = "NULL"
                    campos_atualizados.append(f"CONJUNTO = '{id_conj_a}'")

                if tipo != tipo_a:
                    if tipo_a:
                        tipotete = tip.find(" - ")
                        id_tipo_a = tip[:tipotete]
                    else:
                        id_tipo_a = "NULL"
                    campos_atualizados.append(f"TIPOMATERIAL = '{id_tipo_a}'")

                if projeto != projeto_a:
                    if projeto_a:
                        projetotete = projet.find(" - ")
                        id_projeto_a = projet[:projetotete]
                    else:
                        id_projeto_a = "NULL"
                    campos_atualizados.append(f"PROJETO = {id_projeto_a}")

                if qtde_mini_float != qtde_mini_float_a:
                    campos_atualizados.append(f"QUANTIDADEMIN = '{qtde_mini_float_a}'")

                if ncm != ncm_a:
                    if ncm_a:
                        campos_atualizados.append(f"NCM = '{ncm_a}'")
                    else:
                        campos_atualizados.append(f"NCM = NULL")

                if obs != obs_a:
                    if obs_a:
                        campos_atualizados.append(f"OBS2 = '{obs_a}'")
                    else:
                        campos_atualizados.append(f"OBS2 = NULL")

                if campos_atualizados:
                    campos_update = ", ".join(campos_atualizados)

                    cod_produto = self.line_Codigo.text()

                    print(campos_update)

                    cursor = conecta.cursor()
                    cursor.execute(f"UPDATE produto SET {campos_update} "
                                   f"WHERE codigo = '{cod_produto}';")

                    conecta.commit()

                    self.mensagem_alerta(f"Cadastro do produto {cod_produto} atualizado com Sucesso!")

                    if self.veio_de_fora:
                        foi_alterado = True

                        self.alteracao.emit(foi_alterado)
                        self.close()
                    else:
                        self.limpa_tudo()
                        self.line_Codigo.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaProdutoAlterar("")
    tela.show()
    qt.exec_()

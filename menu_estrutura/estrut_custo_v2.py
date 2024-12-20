import sys
from banco_dados.conexao import conecta
from forms.tela_estrut_custo_v2 import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.lines import validador_inteiro
from comandos.telas import tamanho_aplicacao, icone
from comandos.conversores import valores_para_float, valores_para_virgula, float_para_moeda_reais
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import inspect
import os
from threading import Thread
import traceback


class TelaCustoV2(QMainWindow, Ui_MainWindow):
    def __init__(self, produto, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_estrutura.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Estrutura)
        layout_cabec_tab(self.table_Problema)

        self.processando = False

        self.definir_line_bloqueados()

        self.line_Codigo_Estrut.setFocus()

        self.line_Codigo_Estrut.editingFinished.connect(self.verifica_line_codigo_acabado)

        self.btn_Salvar.clicked.connect(self.salvar_custo_servico)

        self.combo_Versao.activated.connect(self.manipula_versao_escolhida)

        self.progressBar.setHidden(True)

        self.widget_MaoObra.setHidden(True)
        self.widget_Terceiros.setHidden(True)
        self.widget_Compra.setHidden(True)

        validador_inteiro(self.line_Codigo_Estrut, 123456)

        self.lanca_custo_hora_homem()
        self.calcula_total_maodeobra()

        if produto:
            self.line_Codigo_Estrut.setText(produto)
            self.line_Codigo_Estrut.setReadOnly(True)
            self.verifica_line_codigo_acabado()

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

    def atualiza_mascara_moeda(self, unit):
        try:
            unit_float = valores_para_float(unit)
            unit_2casas = ("%.2f" % unit_float)
            valor_string = valores_para_virgula(unit_2casas)
            valor_final = "R$ " + valor_string

            return valor_final

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_line_bloqueados(self):
        try:
            self.line_Descricao_Estrut.setReadOnly(True)
            self.line_Referencia_Estrut.setReadOnly(True)
            self.line_Tipo_Estrut.setReadOnly(True)
            self.line_UM_Estrut.setReadOnly(True)

            self.line_Obs.setReadOnly(True)

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
                    self.limpa_tabelas()
                elif int(codigo_produto) == 0:
                    self.mensagem_alerta('O campo "Código" não pode ser "0"!')
                    self.limpa_dados_produto_estrutura()
                    self.limpa_tabelas()
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
                self.limpa_tabelas()
                self.line_Codigo_Estrut.clear()
            else:
                conjunto = detalhes_produto[0][3]

                if conjunto == 10:
                    self.lanca_dados_acabado()
                else:
                    self.lanca_dados_compras()

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
            self.line_UM_Estrut.setText(um)

            tipo_material = self.line_Tipo_Estrut.text()

            if not tipo_material:
                self.mensagem_alerta('O campo "Tipo de Material" não pode estar vazio!\n\n'
                                     'Entre no cadastro de produtos e defina o Tipo de Material:\n'
                                     'Exemplos: CONJUNTO, USINAGEM, INDUSTRIALIZACAO')
                self.limpa_dados_produto_estrutura()
                self.limpa_tabelas()
                self.line_Codigo_Estrut.clear()
            else:
                self.line_Obs.setText(obs)

                if tipo == "INDUSTRIALIZACAO":
                    self.widget_Terceiros.setHidden(False)
                    self.widget_MaoObra.setHidden(True)
                    self.widget_Compra.setHidden(True)
                else:
                    self.widget_MaoObra.setHidden(False)
                    self.widget_Terceiros.setHidden(True)
                    self.widget_Compra.setHidden(True)

                self.lanca_descricao_tempo_mao_de_obra(codigo_produto)
                self.lanca_descricao_custo_servico(codigo_produto)

                self.combo_Versao.setFocus()

                self.lanca_versoes()

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

            self.combo_Versao.addItems(tabela)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_versao_escolhida(self):
        try:
            self.limpa_tabelas()

            codigo_produto = self.line_Codigo_Estrut.text()
            versao = self.combo_Versao.currentText()

            if codigo_produto and versao:
                self.lanca_estrutura()
                self.soma_total_materiais()
                self.soma_custo_total()

                self.chama_barra_progresso()
                Thread(target=self.produtos_problema).start()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_dados_compras(self):
        try:
            codigo_produto = self.line_Codigo_Estrut.text()
            cur = conecta.cursor()
            cur.execute(f"SELECT prod.descricao, COALESCE(tip.tipomaterial, '') as tipus, "
                        f"COALESCE(prod.obs, '') as ref, prod.unidade, "
                        f"COALESCE(prod.ncm, '') as ncm, COALESCE(prod.obs2, '') as obs, prod.conjunto "
                        f"FROM produto as prod "
                        f"LEFT JOIN tipomaterial tip ON prod.tipomaterial = tip.id "
                        f"where codigo = {codigo_produto};")
            detalhes_produto = cur.fetchall()
            descr, tipo, ref, um, ncm, obs, id_conj = detalhes_produto[0]

            self.line_Descricao_Estrut.setText(descr)
            self.line_Referencia_Estrut.setText(ref)
            self.line_Tipo_Estrut.setText(tipo)
            self.line_UM_Estrut.setText(um)

            tipo_material = self.line_Tipo_Estrut.text()

            if not tipo_material:
                self.mensagem_alerta('O campo "Tipo de Material" não pode estar vazio!\n\n'
                                     'Entre no cadastro de produtos e defina o Tipo de Material:\n'
                                     'Exemplos: CONJUNTO, USINAGEM, INDUSTRIALIZACAO')
                self.limpa_dados_produto_estrutura()
                self.limpa_tabelas()
                self.line_Codigo_Estrut.clear()
            else:
                self.line_Obs.setText(obs)

                self.widget_Terceiros.setHidden(True)
                self.widget_MaoObra.setHidden(True)
                self.widget_Compra.setHidden(False)

                self.lanca_descricao_custo_compra(codigo_produto)

                self.table_Estrutura.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tabelas(self):
        try:
            self.table_Estrutura.setRowCount(0)
            self.table_Problema.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_produto_estrutura(self):
        try:
            self.line_Descricao_Estrut.clear()
            self.line_Tipo_Estrut.clear()
            self.line_Referencia_Estrut.clear()
            self.line_UM_Estrut.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_mao_de_obra_servico(self):
        try:
            self.label_Descricao_Mao.clear()
            self.label_Tempo_Mao.clear()

            self.label_Descricao_Servico.clear()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_tudo(self):
        try:
            self.limpa_tabelas()
            self.limpa_dados_produto_estrutura()
            self.limpa_dados_mao_de_obra_servico()
            self.line_Obs.clear()

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

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, codigo FROM produto where codigo = {codigo_produto};")
            select_prod = cursor.fetchall()
            id_pai, cod = select_prod[0]

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, num_versao, data_versao, obs, data_criacao "
                           f"from estrutura "
                           f"where id_produto = {id_pai} and num_versao = {num_versao};")
            tabela_versoes = cursor.fetchall()
            id_estrutura = tabela_versoes[0][0]

            cursor = conecta.cursor()
            cursor.execute(f"UPDATE produto SET custoestrutura = '{0}' where id = {id_pai};")
            conecta.commit()

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, "
                           f"prod.conjunto, prod.unidade, "
                           f"(estprod.quantidade * 1) as qtde, prod.terceirizado, prod.custounitario, "
                           f"prod.custoestrutura "
                           f"from estrutura_produto as estprod "
                           f"INNER JOIN produto prod ON estprod.id_prod_filho = prod.id "
                           f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                           f"where estprod.id_estrutura = {id_estrutura} "
                           f"order by conj.conjunto DESC, prod.descricao ASC;")
            tabela_estrutura = cursor.fetchall()

            if tabela_estrutura:
                for i in tabela_estrutura:
                    cod, descr, ref, conjunto, um, qtde, terc, unit, estrut = i

                    qtde_float = valores_para_float(qtde)
                    unit_float = valores_para_float(unit)
                    estrut_float = valores_para_float(estrut)

                    if conjunto == 10:
                        total = qtde_float * estrut_float

                        total_rs = self.atualiza_mascara_moeda(total)
                        estru_rs = self.atualiza_mascara_moeda(estrut)

                        dados = (cod, descr, ref, um, qtde, estru_rs, total_rs, conjunto)
                        nova_tabela.append(dados)
                    else:
                        total = qtde_float * unit_float

                        total_rs = self.atualiza_mascara_moeda(total)
                        unit_rs = self.atualiza_mascara_moeda(unit)

                        dados = (cod, descr, ref, um, qtde, unit_rs, total_rs, conjunto)
                        nova_tabela.append(dados)

            if nova_tabela:
                lanca_tabela(self.table_Estrutura, nova_tabela)

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
                        self.label_Descricao_Mao.setText(etapas)
                    if tempo:
                        self.label_Tempo_Mao.setText(str(tempo))

                self.mascara_tempo_mao_de_obra()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_tempo_mao_de_obra(self):
        try:
            tempo_mao = self.label_Tempo_Mao.text()

            if tempo_mao:
                tempo_mao_sem_espacos = tempo_mao.strip()
                if "." in tempo_mao_sem_espacos:
                    string_com_virgula = tempo_mao_sem_espacos.replace('.', ',')
                elif tempo_mao_sem_espacos.startswith(','):
                    string_com_virgula = '0' + tempo_mao_sem_espacos
                else:
                    string_com_virgula = tempo_mao_sem_espacos

                self.label_Tempo_Mao.setText(string_com_virgula)

                self.calcula_total_maodeobra()

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
                        self.label_Descricao_Servico.setText(descr_servico)
                    if custo:
                        self.line_Valor_Servico.setText(str(custo))

                self.mascara_custo_servico()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_descricao_custo_compra(self, codigo):
        try:
            custo_compra_final = "R$ 0,00"
            custo_compra_float = 0

            cursor = conecta.cursor()
            cursor.execute(f"SELECT id, custounitario FROM produto WHERE codigo = {codigo};")
            dados_produto = cursor.fetchall()
            if dados_produto:
                for i in dados_produto:
                    id_prod, custo_compra = i

                    custo_compra_float = valores_para_float(custo_compra)
                    custo_compra_final = float_para_moeda_reais(custo_compra_float)
                    self.line_Custo_Compra.setText(str(custo_compra_final))

            self.label_Custo_Materiais.setText(custo_compra_final)
            self.label_Custo_Total.setText(custo_compra_final)

            preco = (custo_compra_float + (custo_compra_float * 0.05)) / 0.7663
            valor_final = float_para_moeda_reais(preco)

            self.label_Venda_Total.setText(valor_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpar(self):
        self.limpa_tudo()
        self.line_Codigo_Estrut.clear()
        self.line_Codigo_Estrut.setFocus()

    def soma_total_materiais(self):
        try:
            dados_tabela = extrair_tabela(self.table_Estrutura)
            if dados_tabela:
                valor_final = 0.00

                for dados in dados_tabela:
                    total = dados[6]
                    total_sem_cifra = total[2:]
                    if "," in total_sem_cifra:
                        total_1_com_ponto = total_sem_cifra.replace(',', '.')
                        total_1_float = float(total_1_com_ponto)
                    else:
                        total_1_float = float(total_sem_cifra)

                    valor_final = valor_final + total_1_float

                valor_totau_dois = ("%.2f" % valor_final)
                valor_string = str(valor_totau_dois)

                valor_final = "R$ " + valor_string
                self.label_Custo_Materiais.setText(valor_final)

            else:
                valor_final = "R$ 0,00"
                self.label_Custo_Materiais.setText(valor_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def soma_custo_total(self):
        try:
            custo_materias = self.label_Custo_Materiais.text()
            custo_mao_obra = self.label_Total_Mao.text()
            custo_servico = self.line_Valor_Servico.text()

            tipo_material = self.line_Tipo_Estrut.text()

            if tipo_material == "INDUSTRIALIZACAO":
                if custo_servico:
                    custo_ser = valores_para_float(custo_servico)
                else:
                    custo_ser = 0
            else:
                if custo_mao_obra:
                    custo_ser = valores_para_float(custo_mao_obra)
                else:
                    custo_ser = 0

            if custo_materias:
                custo_materias_float = valores_para_float(custo_materias)
            else:
                custo_materias_float = 0
            totalzao = custo_materias_float + custo_ser
            if totalzao:
                totalz = ("%.2f" % totalzao)
                valor_final = "R$ " + str(totalz)
                self.label_Custo_Total.setText(valor_final)

            else:
                valor_final = "R$ 0,00"
                self.label_Custo_Total.setText(valor_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def chama_barra_progresso(self):
        try:
            self.progressBar.setHidden(False)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def produtos_problema(self):
        try:
            self.label_Venda_Total.setText("")

            tabela_nova = []

            extrai_estrutura = extrair_tabela(self.table_Estrutura)
            if extrai_estrutura:
                tab_estrut = []
                for itens in extrai_estrutura:
                    cod = itens[0]
                    qtde = itens[4]

                    estrutura = self.verifica_estrutura_problema(1, cod, qtde)
                    if estrutura:
                        for i in estrutura:
                            tab_estrut.append(i)

                tab_ordenada = sorted(tab_estrut, key=lambda x: -x[0])

                for i in tab_ordenada:
                    niv, codi, descr, ref, um, qtdi, conj, temp, terc, unit, estrut = i

                    if conj == 'PRODUTOS ACABADOS':
                        if temp or terc:
                            pass
                        else:
                            if estrut:
                                estrut_float = float(estrut)
                            else:
                                estrut_float = 0
                            total = float(qtdi) * estrut_float

                            total_rs = self.atualiza_mascara_moeda(total)
                            estru_rs = self.atualiza_mascara_moeda(estrut_float)

                            dados = (codi, descr, ref, um, qtdi, estru_rs, total_rs, conj)
                            tabela_nova.append(dados)
                    else:
                        if unit:
                            pass
                        else:
                            if unit:
                                unit_float = float(unit)
                            else:
                                unit_float = 0

                            total = float(qtdi) * float(unit_float)

                            total_rs = self.atualiza_mascara_moeda(total)
                            unit_rs = self.atualiza_mascara_moeda(unit_float)

                            dados = (codi, descr, ref, um, qtdi, unit_rs, total_rs, conj)
                            tabela_nova.append(dados)

                self.progressBar.setHidden(True)

                if tabela_nova:
                    tabela_nova_ordenada = sorted(tabela_nova, key=lambda x: (x[1], x[0]))

                    lanca_tabela(self.table_Problema, tabela_nova_ordenada)
                else:
                    tipo_material = self.line_Tipo_Estrut.text()

                    if tipo_material == "INDUSTRIALIZACAO":
                        custo_servico = self.line_Valor_Servico.text()
                        if custo_servico:
                            self.calcula_valor_venda()
                    else:
                        tempo_mao = self.label_Tempo_Mao.text()
                        if tempo_mao:
                            self.calcula_valor_venda()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_estrutura_problema(self, nivel, codigo, qtde):
        try:
            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.id, prod.codigo, prod.descricao, prod.obs, prod.unidade, conj.conjunto, "
                           f"prod.tempo, prod.terceirizado, prod.custounitario, prod.custoestrutura "
                           f"FROM produto as prod "
                           f"LEFT JOIN tipomaterial as tip ON prod.tipomaterial = tip.id "
                           f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                           f"where prod.codigo = {codigo};")
            detalhes_pai = cursor.fetchall()
            id_pai, c_pai, des_pai, ref_pai, um_pai, conj_pai, temp_pai, terc_pai, unit_pai, est_pai = detalhes_pai[0]

            filhos = [(nivel, codigo, des_pai, ref_pai, um_pai, qtde, conj_pai, temp_pai, terc_pai, unit_pai, est_pai)]

            nivel_plus = nivel + 1

            cursor = conecta.cursor()
            cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.obs, '') as obs, prod.unidade, "
                           f"(mat.quantidade * {qtde}) as qtde "
                           f"FROM materiaprima as mat "
                           f"INNER JOIN produto prod ON mat.produto = prod.id "
                           f"where mestre = {id_pai};")
            dados_estrutura = cursor.fetchall()

            if dados_estrutura:
                for prod in dados_estrutura:
                    cod_f, descr_f, ref_f, um_f, qtde_f = prod

                    filhos.extend(self.verifica_estrutura_problema(nivel_plus, cod_f, qtde_f))

            return filhos

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calcula_valor_venda(self):
        try:
            extrai_dados = extrair_tabela(self.table_Problema)
            if not extrai_dados:
                custo_total = self.label_Custo_Total.text()
                custo_tot_float = valores_para_float(custo_total)

                preco = (custo_tot_float + (custo_tot_float * 0.05)) / 0.7663

                valor_totau_dois = ("%.2f" % preco)
                valor_string = str(valor_totau_dois)

                valor_final = "R$ " + valor_string
                self.label_Venda_Total.setText(valor_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_custo_hora_homem(self):
        try:
            cursor = conecta.cursor()
            cursor.execute("SELECT valorhora FROM valoresmensais WHERE data = (SELECT MAX(data) FROM valoresmensais);")
            valores_mensais = cursor.fetchall()
            valor_mensal = f"R$ {(valores_mensais[0][0])}"

            self.label_custo_Mao.setText(valor_mensal)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def calcula_total_maodeobra(self):
        try:
            custo_mao = self.label_custo_Mao.text()
            custo_mao_float = valores_para_float(custo_mao)

            tempo_mao = self.label_Tempo_Mao.text()
            tempo_mao_str = valores_para_virgula(tempo_mao)
            tempo_mao_float = valores_para_float(tempo_mao_str)

            total1 = tempo_mao_float * custo_mao_float
            total_2casas = str("%.2f" % total1)
            total_str = valores_para_virgula(total_2casas)
            total = f"R$ {total_str}"
            self.label_Total_Mao.setText(total)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mascara_custo_servico(self):
        try:
            custo_servico = self.line_Valor_Servico.text()

            if custo_servico:
                custo_servico_sem_espacos = custo_servico.strip()
                if "." in custo_servico_sem_espacos:
                    string_com_virgula = custo_servico_sem_espacos.replace('.', ',')
                elif custo_servico_sem_espacos.startswith(','):
                    string_com_virgula = '0' + custo_servico_sem_espacos
                else:
                    string_com_virgula = custo_servico_sem_espacos

                if "R$ " in string_com_virgula:
                    final = string_com_virgula
                else:
                    final = "R$ " + string_com_virgula

                self.line_Valor_Servico.setText(final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salvar_custo_servico(self):
        try:
            self.label_Venda_Total.setText("")

            codigo_produto = self.line_Codigo_Estrut.text()

            if codigo_produto:
                terceirizado = self.line_Valor_Servico.text()
                if terceirizado:
                    terc_float = valores_para_float(terceirizado)
                else:
                    terc_float = 0

                cursor = conecta.cursor()
                cursor.execute(f"SELECT id, terceirizado FROM produto WHERE codigo = {codigo_produto};")
                valores = cursor.fetchall()

                mestre, terc_banco = valores[0]
                if terc_banco:
                    terc_banco_float = float(terc_banco)
                else:
                    terc_banco_float = 0

                if terc_banco_float != terc_float:
                    cursor = conecta.cursor()
                    cursor.execute(f"UPDATE produto SET terceirizado = '{terc_float}' "
                                   f"where id = {mestre};")
                    conecta.commit()

                    self.mensagem_alerta(f"Custo do Serviço de Terceiros atualizado com sucesso!")

                    self.soma_custo_total()

                    tipo_material = self.line_Tipo_Estrut.text()

                    if tipo_material == "INDUSTRIALIZACAO":
                        custo_servico = self.line_Valor_Servico.text()
                        if custo_servico:
                            self.calcula_valor_venda()
                    else:
                        tempo_mao = self.label_Tempo_Mao.text()
                        if tempo_mao:
                            self.calcula_valor_venda()

                    self.mascara_custo_servico()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaCustoV2("")
    tela.show()
    qt.exec_()

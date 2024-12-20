import sys
from banco_dados.conexao import conecta
from forms.tela_prod_ficha import *
from banco_dados.controle_erros import grava_erro_banco
from arquivos.chamar_arquivos import definir_caminho_arquivo
from comandos.tabelas import extrair_tabela, lanca_tabela, layout_cabec_tab
from comandos.telas import tamanho_aplicacao, icone
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
import inspect
import os
import traceback
from threading import Thread

from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
from datetime import datetime


class TelaFichasProdutos(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        icone(self, "menu_estrutura.png")
        tamanho_aplicacao(self)
        layout_cabec_tab(self.table_Estrutura)

        self.arquivos_pra_excluir = []

        self.widget_Progress.setHidden(True)

        self.processando = False

        self.line_Codigo_Manu.editingFinished.connect(self.verifica_line_codigo_manual)

        self.btn_PDF.clicked.connect(self.final)

        self.btn_ExcluirTudo.clicked.connect(self.excluir_tudo_tab_produtos)
        self.btn_ExcluirItem.clicked.connect(self.excluir_produto_tab)

        validator = QtGui.QRegExpValidator(QtCore.QRegExp(r'\d+'), self.line_Codigo_Manu)
        self.line_Codigo_Manu.setValidator(validator)

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

    def verifica_line_codigo_manual(self):
        if not self.processando:
            try:
                self.processando = True

                codigo_produto = self.line_Codigo_Manu.text()

                self.label_Texto.setText("")

                if not codigo_produto:
                    self.mensagem_alerta('O campo "Código" não pode estar vazio!')
                    self.line_Codigo_Manu.clear()
                elif int(codigo_produto) == 0:
                    self.mensagem_alerta('O campo "Código" não pode ser "0"!')
                    self.line_Codigo_Manu.clear()
                else:
                    self.verifica_sql_produto_manual()

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

            extrai_estrutura = extrair_tabela(self.table_Estrutura)

            ja_existe = False
            for itens in extrai_estrutura:
                cod_con = itens[0]
                if cod_con == codigo_produto:
                    ja_existe = True
                    break

            if not ja_existe:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.DESCRICAOCOMPLEMENTAR, ''), "
                               f"COALESCE(prod.obs, ''), "
                               f"COALESCE(prod.ncm, '') as ncm, conj.conjunto, prod.localizacao, "
                               f"prod.unidade, prod.quantidade "
                               f"FROM produto as prod "
                               f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                               f"where codigo = {codigo_produto};")
                detalhes_produto = cursor.fetchall()
                cod, descr, compl, ref, ncm, conjunto, local, um, saldo = detalhes_produto[0]

                dados1 = [cod, descr, compl, ref, um, ncm, conjunto, saldo, local]
                extrai_estrutura.append(dados1)

                if extrai_estrutura:
                    lanca_tabela(self.table_Estrutura, extrai_estrutura)
                    self.line_Codigo_Manu.clear()
                    self.line_Codigo_Manu.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_produto_tab(self):
        try:
            nome_tabela = self.table_Estrutura

            dados_tab = extrair_tabela(nome_tabela)
            if not dados_tab:
                self.mensagem_alerta(f'A tabela "Produtos Ordem de Compra" está vazia!')
            else:
                linha = nome_tabela.currentRow()
                if linha >= 0:
                    nome_tabela.removeRow(linha)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def excluir_tudo_tab_produtos(self):
        try:
            nome_tabela = self.table_Estrutura

            dados_tab = extrair_tabela(nome_tabela)
            if not dados_tab:
                self.mensagem_alerta(f'A tabela "Estrutura" está vazia!')
            else:
                nome_tabela.setRowCount(0)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpar(self):
        self.line_Codigo_Manu.clear()
        self.limpa_tabela()

    def excluir_arquivo(self, caminho_arquivo):
        try:
            if os.path.exists(caminho_arquivo):
                os.remove(caminho_arquivo)
            else:
                print("O arquivo não existe no caminho especificado.")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def criar_texto(self, imgs, pos_horizontal, pos_vertical, texto, cor, fonte, largura_tra):
        try:
            draw = ImageDraw.Draw(imgs)
            draw.text((pos_horizontal, pos_vertical), texto, fill=cor, font=fonte, stroke_width=largura_tra)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def cria_imagem_da_ficha1(self, cod, descr, ref, um, local):
        try:
            camino = os.path.join('..', 'arquivos', 'modelo excel', 'ficha_prod_modelo.png')
            caminho_arquivo = definir_caminho_arquivo(camino)

            imgs = Image.open(caminho_arquivo)

            font_cod = ImageFont.truetype("tahoma.ttf", 100)
            font_descricao = ImageFont.truetype("tahoma.ttf", 60)

            self.criar_texto(imgs, 1510, 130, cod, (0, 0, 0), font_cod, 3)
            self.criar_texto(imgs, 680, 335, descr, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 615, 455, ref, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 1700, 455, um, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 645, 575, local, (0, 0, 0), font_descricao, 0)

            camino_f = os.path.join('..', 'arquivos', 'modelo excel', f'ficha_produto_{cod}.png')
            arquivo_final = definir_caminho_arquivo(camino_f)
            imgs.save(arquivo_final)

            self.arquivos_pra_excluir.append(arquivo_final)

            return arquivo_final

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def cria_imagem_da_ficha2(self, imagem_ficha1, cod, descr, ref, um, local):
        try:
            imgs = Image.open(imagem_ficha1)

            font_cod = ImageFont.truetype("tahoma.ttf", 100)
            font_descricao = ImageFont.truetype("tahoma.ttf", 60)

            self.criar_texto(imgs, 3050, 130, cod, (0, 0, 0), font_cod, 3)
            self.criar_texto(imgs, 2220, 335, descr, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 2155, 455, ref, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 3240, 455, um, (0, 0, 0), font_descricao, 0)
            self.criar_texto(imgs, 2185, 575, local, (0, 0, 0), font_descricao, 0)

            camino_f = os.path.join('..', 'arquivos', 'modelo excel', f'ficha_produto_{cod}.png')
            arquivo_final = definir_caminho_arquivo(camino_f)
            imgs.save(arquivo_final)

            self.arquivos_pra_excluir.append(arquivo_final)

            return arquivo_final

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def converte_varias_imagens_para_pdf(self, imagens_fichas):
        try:
            pil_images = [Image.open(imagem).convert("RGB") for imagem in imagens_fichas]

            desktop = Path.home() / "Desktop"
            timestamp = datetime.now().strftime("%d%H%M")
            nome_req = f'/Ficha_{timestamp}.pdf'
            caminho = str(desktop) + nome_req

            pil_images[0].save(caminho, save_all=True, append_images=pil_images[1:], format="PDF", resolution=100.0)

            print(f"PDF com várias páginas salvo em {caminho}")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def final(self):
        try:
            self.label_Texto.setText("")

            self.widget_Progress.setHidden(False)

            Thread(target=self.gerar_pdf).start()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def gerar_pdf(self):
        try:
            imagens_fichas = []

            extrai_estrutura = extrair_tabela(self.table_Estrutura)

            if extrai_estrutura:
                for i in range(0, len(extrai_estrutura), 2):
                    if i == 0:
                        msg = "Ajustando Layout.."
                        self.label_Texto.setText(msg)
                    if i == 2:
                        msg = "Organizando produtos.."
                        self.label_Texto.setText(msg)
                    if i == 4:
                        msg = "Criando arquivo PDF.."
                        self.label_Texto.setText(msg)

                    codigo_produto1 = extrai_estrutura[i][0]

                    cursor = conecta.cursor()
                    cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.DESCRICAOCOMPLEMENTAR, ''), "
                                   f"COALESCE(prod.obs, ''), "
                                   f"COALESCE(prod.ncm, '') as ncm, conj.conjunto, prod.localizacao, "
                                   f"prod.unidade, prod.quantidade "
                                   f"FROM produto as prod "
                                   f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                                   f"where codigo = {codigo_produto1};")
                    detalhes_produto1 = cursor.fetchone()
                    cod1, descr1, compl1, ref1, ncm1, conjunto1, local1, um1, saldo1 = detalhes_produto1

                    imagem_ficha1 = self.cria_imagem_da_ficha1(cod1, descr1, ref1, um1, local1)

                    if i + 1 < len(extrai_estrutura):
                        codigo_produto2 = extrai_estrutura[i + 1][0]

                        cursor.execute(f"SELECT prod.codigo, prod.descricao, COALESCE(prod.DESCRICAOCOMPLEMENTAR, ''), "
                                       f"COALESCE(prod.obs, ''), "
                                       f"COALESCE(prod.ncm, '') as ncm, conj.conjunto, prod.localizacao, "
                                       f"prod.unidade, prod.quantidade "
                                       f"FROM produto as prod "
                                       f"INNER JOIN conjuntos conj ON prod.conjunto = conj.id "
                                       f"where codigo = {codigo_produto2};")
                        detalhes_produto2 = cursor.fetchone()
                        cod2, descr2, compl2, ref2, ncm2, conjunto2, local2, um2, saldo2 = detalhes_produto2

                        imagem_ficha2 = self.cria_imagem_da_ficha2(imagem_ficha1, cod2, descr2, ref2, um2, local2)
                        imagens_fichas.append(imagem_ficha2)
                    else:
                        imagens_fichas.append(imagem_ficha1)

                self.converte_varias_imagens_para_pdf(imagens_fichas)

                if self.arquivos_pra_excluir:
                    for arqs in self.arquivos_pra_excluir:
                        msg = "Excluindo arquivos de modelo.."
                        self.label_Texto.setText(msg)

                        self.excluir_arquivo(arqs)
                        self.label_Texto.setText("Arquivo PDF criado com sucesso!")

                self.widget_Progress.setHidden(True)
                self.limpar()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaFichasProdutos()
    tela.show()
    qt.exec_()

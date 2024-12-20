import sys
from forms.tela_oc_alterar_prod import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.cores import cor_branco
from comandos.conversores import valores_para_float, valores_para_virgula
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtCore import pyqtSignal, QDate, Qt
import inspect
import os
from datetime import date, timedelta, datetime
import traceback


class TelaOcAlterarProduto(QMainWindow, Ui_MainWindow):
    produto_escolhido = pyqtSignal(list)

    def __init__(self, lista_produto, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        self.setWindowFlags(Qt.Window | Qt.CustomizeWindowHint | Qt.WindowTitleHint | Qt.WindowCloseButtonHint)

        self.data_emissao_oc = ""

        self.definir_bloqueios()
        self.lanca_produto(lista_produto)

        self.btn_Alterar.clicked.connect(self.verifica_dados_completos)

        self.line_Qtde.editingFinished.connect(self.verifica_line_qtde)
        self.line_Unit.editingFinished.connect(self.atualiza_mascara_unit)
        self.line_Ipi.editingFinished.connect(self.atualiza_mascara_ipi)

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

    def verifica_line_qtde(self):
        if not self.processando:
            try:
                self.processando = True

                qtde = self.line_Qtde.text()

                if qtde:
                    qtde_com_virgula = valores_para_virgula(qtde)

                    self.line_Qtde.setText(qtde_com_virgula)

                    self.calcular_valor_total_prod()

                    self.line_Unit.setFocus()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def atualiza_mascara_unit(self):
        if not self.processando:
            try:
                self.processando = True

                unit = self.line_Unit.text()

                if unit:
                    unit_float = valores_para_float(unit)
                    unit_2casas = ("%.4f" % unit_float)
                    valor_string = valores_para_virgula(unit_2casas)
                    valor_final = "R$ " + valor_string
                    self.line_Unit.setText(valor_final)

                    self.calcular_valor_total_prod()

                    self.line_Ipi.setFocus()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def atualiza_mascara_ipi(self):
        if not self.processando:
            try:
                self.processando = True

                ipi = self.line_Ipi.text()

                ipi_float = valores_para_float(ipi)
                ipi_2casas = ("%.2f" % ipi_float)
                valor_string = valores_para_virgula(ipi_2casas)
                valor_final = valor_string + "%"
                self.line_Ipi.setText(valor_final)

                self.calcular_valor_total_prod()

                self.line_ValorTotal.setFocus()

            except Exception as e:
                nome_funcao = inspect.currentframe().f_code.co_name
                exc_traceback = sys.exc_info()[2]
                self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

            finally:
                self.processando = False

    def calcular_valor_total_prod(self):
        try:
            qtde = self.line_Qtde.text()
            unit = self.line_Unit.text()

            if qtde and unit:
                qtde_float = valores_para_float(qtde)

                unit_float = valores_para_float(unit)

                valor_total = qtde_float * unit_float

                total_2casas = ("%.2f" % valor_total)
                valor_string = valores_para_virgula(total_2casas)

                valor_final = "R$ " + valor_string
                self.line_ValorTotal.setText(valor_final)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_bloqueios(self):
        try:
            self.line_Codigo.setReadOnly(True)

            self.line_Descricao.setReadOnly(True)
            self.line_Referencia.setReadOnly(True)
            self.line_UM.setReadOnly(True)

            self.line_Num_Req.setReadOnly(True)
            self.line_Item_Req.setReadOnly(True)

            self.line_ValorTotal.setReadOnly(True)

            self.line_Qtde_Nf.setReadOnly(True)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def lanca_produto(self, lista_prod):
        try:
            if lista_prod:
                num_req, item_req, cod, descr, ref, um, qtde, unit, ipi, total, entrega, qtde_nf, emissao = lista_prod

                self.line_Num_Req.setText(str(num_req))
                self.line_Item_Req.setText(str(item_req))

                self.line_Codigo.setText(str(cod))
                self.line_Descricao.setText(descr)
                self.line_Referencia.setText(ref)
                self.line_UM.setText(um)

                self.line_Qtde.setText(str(qtde))

                self.line_Unit.setText(str(unit))
                self.line_Ipi.setText(str(ipi))
                self.line_ValorTotal.setText(str(total))

                entrega_date = QDate.fromString(entrega, 'dd/MM/yyyy')
                self.date_Entrega.setDate(entrega_date)

                self.line_Qtde_Nf.setText(str(qtde_nf))

                self.data_emissao_oc = emissao

                self.date_Entrega.setFocus()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def verifica_dados_completos(self):
        try:
            qtde = self.line_Qtde.text()
            unit = self.line_Unit.text()

            if self.data_emissao_oc:
                data_emissao_oc = datetime.strptime(self.data_emissao_oc, "%d/%m/%Y").date()

                data_entrega_str = self.date_Entrega.text()
                data_entrega = datetime.strptime(data_entrega_str, '%d/%m/%Y').date()

                if data_entrega < data_emissao_oc:
                    self.mensagem_alerta(f'A data de Entrega não pode ser menor que a data de '
                                         f'Emissão da Ordem: {self.data_emissao_oc}')
                elif not qtde:
                    self.mensagem_alerta('O campo "Qtde" não pode estar vazio')
                    self.line_Qtde.setFocus()
                elif not unit:
                    self.mensagem_alerta('O campo "R$/Unid" não pode estar vazio')
                    self.line_Unit.setFocus()
                else:
                    qtde_nf = self.line_Qtde_Nf.text()
                    qtde_nf_float = valores_para_float(qtde_nf)

                    qtde_float = valores_para_float(qtde)

                    if qtde_float < qtde_nf_float:
                        self.mensagem_alerta("A quantidade da Ordem não pode ser menor que a quantidade "
                                             "das Notas Fiscais de compra!")
                    else:
                        unit_float = valores_para_float(unit)

                        if qtde_float <= 0:
                            self.mensagem_alerta("A quantidade deve ser maior que 0!")
                        elif unit_float <= 0:
                            self.mensagem_alerta("O valor unitário deve ser maior que 0!")
                        else:
                            self.manipula_dados_tabela()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def manipula_dados_tabela(self):
        try:
            cod_produto = self.line_Codigo.text()
            descr = self.line_Descricao.text()
            ref = self.line_Referencia.text()
            um = self.line_UM.text()

            num_req = self.line_Num_Req.text()
            item_req = self.line_Item_Req.text()

            qtde = self.line_Qtde.text()
            unit = self.line_Unit.text()
            ipi = self.line_Ipi.text()
            total = self.line_ValorTotal.text()
            entrega = self.date_Entrega.text()
            qtde_nf = self.line_Qtde_Nf.text()

            dados = [num_req, item_req, cod_produto, descr, ref, um, qtde, unit, ipi, total, entrega, qtde_nf]

            self.produto_escolhido.emit(dados)
            self.close()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_entrega(self):
        try:
            data_hoje = date.today()
            data_entrega = data_hoje + timedelta(days=14)
            self.date_Entrega.setDate(data_entrega)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def limpa_dados_produtos(self):
        try:
            self.line_Num_Req.clear()
            self.line_Item_Req.clear()

            self.line_Codigo.clear()
            self.line_Descricao.clear()
            self.line_Referencia.clear()
            self.line_UM.clear()

            self.line_Qtde.clear()
            self.line_Unit.clear()
            self.line_Ipi.clear()
            self.line_ValorTotal.clear()

            self.line_Referencia.setStyleSheet(f"background-color: {cor_branco};")

            self.definir_entrega()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    tela = TelaOcAlterarProduto("")
    tela.show()
    qt.exec_()

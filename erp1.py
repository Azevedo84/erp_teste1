import sys
from banco_dados.conexao import conecta
from forms.tela_menu1 import *
from banco_dados.controle_erros import grava_erro_banco
from comandos.telas import tamanho_aplicacao, icone, editar_botao
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QPixmap
import os
import inspect
import socket
from threading import Thread
import getpass
import subprocess
from datetime import datetime
import traceback

from menu_cadastros.prod_principal import TelaProdutoPrincipal
from menu_compras.sol_incluir_v2 import TelaSolIncluirV2
from menu_producao.op_incluir_v2 import TelaOpIncluirV2
from menu_consumiveis.ci_incluir import TelaCiIncluir
from menu_vendas.pi_incluir import TelaPiIncluir


class TelaMenu(QMainWindow, Ui_Menu_Principal):
    def __init__(self, parent=None):
        super().__init__(parent)
        super().setupUi(self)

        nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
        self.nome_arquivo = os.path.basename(nome_arquivo_com_caminho)

        tamanho_aplicacao(self)

        tamanho_icone = 50
        editar_botao(self.btn_1, "produto.png", 'Produto', tamanho_icone)
        self.btn_1.clicked.connect(self.chama_tela_prod_consultar)

        editar_botao(self.btn_2, "compras.png", 'Solicitação', tamanho_icone)
        self.btn_2.clicked.connect(self.chama_tela_sol_incluir)

        editar_botao(self.btn_3, "producao.png", 'Ordem de Produção', tamanho_icone)
        self.btn_3.clicked.connect(self.chama_tela_op_incluir)

        editar_botao(self.btn_4, "consumiveis.png", 'Consumo Interno', tamanho_icone)
        self.btn_4.clicked.connect(self.chama_tela_consumo_interno)

        editar_botao(self.btn_5, "vendas.png", 'Pedido Interno', tamanho_icone)
        self.btn_5.clicked.connect(self.chama_tela_pi_incluir)

        self.versao = f"Versão 2.06.000"
        self.data_versao = f"08/11/2024"

        self.label_versao.setText(self.versao)
        self.label_DataVersao.setText(self.data_versao)

        pixmap = QPixmap('arquivos/Logo_sem_fundo.png')
        self.label.setPixmap(pixmap)

        self.pre_incluir = []
        self.pre_status = []

        self.prod_incluir = []
        self.prod_alterar = []
        self.prod_pesquisar = []
        self.prod_consultar = []
        self.prod_ficha = []

        self.cad_cliente = []
        self.cad_fornecedor = []
        self.cad_conjunto = []
        self.cad_funcionario = []
        self.cad_cfop = []
        self.cad_tipo_material = []
        self.cad_projeto = []
        self.local_estoque = []

        self.sol_incluir = []
        self.req_incluir = []
        self.oc_incluir = []
        self.oc_alterar = []
        self.compras_status = []

        self.ci_incluir = []

        self.est_final = []
        self.est_mov = []

        self.estr_incluir = []
        self.estr_custo = []
        self.estr_versao = []

        self.pcp_previsao = []
        self.pcp_produto = []

        self.op_incluir = []
        self.op_lote = []
        self.op_alterar = []
        self.op_consumo = []
        self.op_encerar = []
        self.op_excluir = []
        self.producao_status = []

        self.pi_incluir = []
        self.pi_alterar = []
        self.ov_incluir = []
        self.ov_alterar = []
        self.vendas_status = []

        icone(self, "menu_menu.png")

        self.definir_comando_telas()
        self.nome_computador = socket.gethostname()
        self.username = getpass.getuser()

        cur = conecta.cursor()
        cur.execute(f"SELECT id, descricao, nome_usuario, botoes FROM ENVIA_PC "
                    f"where descricao = '{self.nome_computador}' "
                    f"and nome_usuario = '{self.username}';")
        dados_usuario = cur.fetchall()
        if dados_usuario:
            botoes = dados_usuario[0][3]
            print(botoes)

        Thread(target=self.funcao_macro_tred).start()

        self.ultima_versao()
        self.salva_versao()

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

    def definir_comando_telas(self):
        try:
            self.action_Pre_Incluir_2.triggered.connect(self.definir_tela_action)
            self.action_Pre_Status_2.triggered.connect(self.definir_tela_action)

            self.action_Prod_Incluir.triggered.connect(self.definir_tela_action)
            self.actionAlterar_Produto.triggered.connect(self.definir_tela_action)
            self.actionPesquisar_Produto.triggered.connect(self.definir_tela_action)
            self.action_Consultar_Produto.triggered.connect(self.chama_tela_prod_consultar)
            self.actionImprimir_Fichas.triggered.connect(self.definir_tela_action)

            self.actionCliente.triggered.connect(self.definir_tela_action)
            self.actionFornecedor.triggered.connect(self.definir_tela_action)
            self.actionFuncionario.triggered.connect(self.definir_tela_action)
            self.actionConjunto.triggered.connect(self.definir_tela_action)
            self.actionNatureza_Opera_o_CFOP.triggered.connect(self.definir_tela_action)
            self.actionTipo_de_Material.triggered.connect(self.definir_tela_action)
            self.actionProjetos.triggered.connect(self.definir_tela_action)
            self.actionLocais_de_Estoque.triggered.connect(self.definir_tela_action)

            self.action_Sol_Incluir.triggered.connect(self.chama_tela_sol_incluir)
            self.action_Req_Incluir.triggered.connect(self.definir_tela_action)
            self.action_OC_Incluir.triggered.connect(self.definir_tela_action)
            self.action_OC_Alterar.triggered.connect(self.definir_tela_action)
            self.action_Compras_Status.triggered.connect(self.definir_tela_action)

            self.action_CI_Incluir.triggered.connect(self.chama_tela_consumo_interno)

            self.action_Est_Final.triggered.connect(self.definir_tela_action)
            self.action_Est_Mov.triggered.connect(self.definir_tela_action)

            self.action_Estr_Incluir.triggered.connect(self.definir_tela_action)
            self.action_Estr_Custo.triggered.connect(self.definir_tela_action)
            self.actionCadastrar_Vers_o.triggered.connect(self.definir_tela_action)

            self.action_Pcp_Previsao.triggered.connect(self.definir_tela_action)
            self.action_Pcp_Produto.triggered.connect(self.definir_tela_action)

            self.action_OP_Incluir.triggered.connect(self.chama_tela_op_incluir)
            self.action_OP_Lote.triggered.connect(self.definir_tela_action)
            self.action_OP_Alterar.triggered.connect(self.definir_tela_action)
            self.action_OP_Consumo.triggered.connect(self.definir_tela_action)
            self.action_OP_Encerar.triggered.connect(self.definir_tela_action)
            self.action_OP_Excluir.triggered.connect(self.definir_tela_action)
            self.action_Producao_Status.triggered.connect(self.definir_tela_action)

            self.action_PI_Incluir.triggered.connect(self.chama_tela_pi_incluir)
            self.action_PI_Alterar.triggered.connect(self.definir_tela_action)
            self.action_OV_Incluir.triggered.connect(self.definir_tela_action)
            self.action_OV_Alterar.triggered.connect(self.definir_tela_action)
            self.action_Vendas_Status.triggered.connect(self.definir_tela_action)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def definir_tela_action(self):
        try:
            sender = self.sender()

            if sender == self.action_Pre_Incluir_2:
                from menu_cadastros.pre_incluir import TelaPreIncluir
                self.pre_incluir = TelaPreIncluir()
                self.pre_incluir.show()

            elif sender == self.action_Pre_Status_2:
                from menu_cadastros.pre_status import TelaPreStatus
                self.pre_status = TelaPreStatus()
                self.pre_status.show()

            elif sender == self.action_Prod_Incluir:
                from menu_cadastros.pro_incluir import TelaProdutoIncluir
                self.prod_incluir = TelaProdutoIncluir()
                self.prod_incluir.show()

            elif sender == self.actionAlterar_Produto:
                from menu_cadastros.prod_alterar import TelaProdutoAlterar
                self.prod_alterar = TelaProdutoAlterar("")
                self.prod_alterar.show()

            elif sender == self.actionPesquisar_Produto:
                from menu_cadastros.prod_pesquisar import TelaProdutoPesquisar
                self.prod_pesquisar = TelaProdutoPesquisar("", False)
                self.prod_pesquisar.show()

            elif sender == self.actionImprimir_Fichas:
                from menu_cadastros.prod_fichas import TelaFichasProdutos
                self.prod_ficha = TelaFichasProdutos()
                self.prod_ficha.show()

            elif sender == self.actionCliente:
                from menu_cadastros.cad_cliente import TelaCadastroCliente
                self.cad_cliente = TelaCadastroCliente()
                self.cad_cliente.show()

            elif sender == self.actionFornecedor:
                from menu_cadastros.cad_fornecedor import TelaCadastroFornecedor
                self.cad_fornecedor = TelaCadastroFornecedor()
                self.cad_fornecedor.show()

            elif sender == self.actionFuncionario:
                from menu_cadastros.cad_funcionario import TelaCadastroFuncionario
                self.cad_funcionario = TelaCadastroFuncionario()
                self.cad_funcionario.show()

            elif sender == self.actionConjunto:
                from menu_cadastros.cad_conjunto import TelaCadastroConjunto
                self.cad_conjunto = TelaCadastroConjunto()
                self.cad_conjunto.show()

            elif sender == self.actionNatureza_Opera_o_CFOP:
                from menu_cadastros.cad_natureza import TelaCadastroNatureza
                self.cad_cfop = TelaCadastroNatureza()
                self.cad_cfop.show()

            elif sender == self.actionTipo_de_Material:
                from menu_cadastros.cad_tipomaterial import TelaCadastroTipoMaterial
                self.cad_tipo_material = TelaCadastroTipoMaterial()
                self.cad_tipo_material.show()

            elif sender == self.actionProjetos:
                from menu_cadastros.cad_projeto import TelaCadastroProjeto
                self.cad_projeto = TelaCadastroProjeto()
                self.cad_projeto .show()

            elif sender == self.actionLocais_de_Estoque:
                from menu_cadastros.cad_local_estoque import TelaCadastroLocalEstoque
                self.local_estoque = TelaCadastroLocalEstoque()
                self.local_estoque.show()

            elif sender == self.action_Req_Incluir:
                from menu_compras.req_incluir import TelaReqIncluir
                self.req_incluir = TelaReqIncluir()
                self.req_incluir.show()

            elif sender == self.action_OC_Incluir:
                from menu_compras.oc_incluir import TelaOcIncluir
                self.oc_incluir = TelaOcIncluir()
                self.oc_incluir.show()

            elif sender == self.action_OC_Alterar:
                from menu_compras.oc_alterar import TelaOcAlterar
                self.oc_alterar = TelaOcAlterar()
                self.oc_alterar.show()

            elif sender == self.action_Compras_Status:
                from menu_compras.compras_status import TelaComprasStatus
                self.compras_status = TelaComprasStatus()
                self.compras_status.show()

            elif sender == self.action_Est_Final:
                from menu_estoque.est_estoque import TelaEstEstoque
                self.est_final = TelaEstEstoque()
                self.est_final.show()

            elif sender == self.action_Est_Mov:
                from menu_estoque.est_mov_v2 import TelaEstMovimentacaoV2
                self.est_mov = TelaEstMovimentacaoV2()
                self.est_mov.show()

            elif sender == self.action_Estr_Incluir:
                from menu_estrutura.estrut_incluir_v2 import TelaEstruturaIncluirV2
                self.estr_incluir = TelaEstruturaIncluirV2("")
                self.estr_incluir.show()

            elif sender == self.action_Estr_Custo:
                from menu_estrutura.estrut_custo_v2 import TelaCustoV2
                self.estr_custo = TelaCustoV2("")
                self.estr_custo.show()

            elif sender == self.actionCadastrar_Vers_o:
                from menu_estrutura.estrut_versao import TelaCadastroVersoes
                self.estr_versao = TelaCadastroVersoes()
                self.estr_versao.show()

            elif sender == self.action_Pcp_Previsao:
                from menu_pcp.pcp_previsao_v2 import TelaPcpPrevisaoV2
                self.pcp_previsao = TelaPcpPrevisaoV2()
                self.pcp_previsao.show()

            elif sender == self.action_Pcp_Produto:
                from menu_pcp.pcp_produto_v2 import TelaPcpProdutoV2
                self.pcp_produto = TelaPcpProdutoV2("")
                self.pcp_produto.show()

            elif sender == self.action_OP_Alterar:
                from menu_producao.op_alterar_v2 import TelaOpAlterarV2
                self.op_alterar = TelaOpAlterarV2()
                self.op_alterar.show()

            elif sender == self.action_OP_Consumo:
                from menu_producao.op_consumir_v2 import TelaOpConsumirV2
                self.op_consumo = TelaOpConsumirV2()
                self.op_consumo.show()

            elif sender == self.action_OP_Encerar:
                from menu_producao.op_encerrar_v2 import TelaOpEncerrarV2
                self.op_encerar = TelaOpEncerrarV2()
                self.op_encerar.show()

            elif sender == self.action_OP_Excluir:
                from menu_producao.op_excluir_v2 import TelaOpExcluirV2
                self.op_excluir = TelaOpExcluirV2()
                self.op_excluir.show()

            elif sender == self.action_Producao_Status:
                from menu_producao.producao_status_v2 import TelaOpStatusV2
                self.producao_status = TelaOpStatusV2()
                self.producao_status.show()

            elif sender == self.action_PI_Alterar:
                from menu_vendas.pi_alterar import TelaPiAlterar
                self.pi_alterar = TelaPiAlterar()
                self.pi_alterar.show()

            elif sender == self.action_OV_Incluir:
                from menu_vendas.ov_incluir import TelaOvIncluir
                self.ov_incluir = TelaOvIncluir()
                self.ov_incluir.show()

            elif sender == self.action_OV_Alterar:
                from menu_vendas.ov_alterar import TelaOvAlterar
                self.ov_alterar = TelaOvAlterar()
                self.ov_alterar.show()

            elif sender == self.action_Vendas_Status:
                from menu_vendas.vendas_status import TelaVendasStatus
                self.vendas_status = TelaVendasStatus()
                self.vendas_status.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def chama_tela_prod_consultar(self):
        try:
            self.prod_consultar = TelaProdutoPrincipal()
            self.prod_consultar.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def chama_tela_sol_incluir(self):
        try:
            self.sol_incluir = TelaSolIncluirV2()
            self.sol_incluir.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def chama_tela_op_incluir(self):
        try:
            self.op_incluir = TelaOpIncluirV2()
            self.op_incluir.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def chama_tela_consumo_interno(self):
        try:
            self.ci_incluir = TelaCiIncluir()
            self.ci_incluir.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def chama_tela_pi_incluir(self):
        try:
            self.pi_incluir = TelaPiIncluir()
            self.pi_incluir.show()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def funcao_macro_tred(self):
        try:
            self.envia_email()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def mensagem_email(self):
        try:
            from datetime import datetime

            current_time = (datetime.now())
            horario = current_time.strftime('%H')
            hora_int = int(horario)
            saudacao = ""
            if 4 < hora_int < 13:
                saudacao = "Bom Dia!"
            elif 12 < hora_int < 19:
                saudacao = "Boa Tarde!"
            elif hora_int > 18:
                saudacao = "Boa Noite!"
            elif hora_int < 5:
                saudacao = "Boa Noite!"

            msg_final = f"Att,\n" \
                        f"Suzuki Máquinas Ltda\n" \
                        f"Fone (51) 3561.2583/(51) 3170.0965\n\n" \
                        f"Mensagem enviada automaticamente, por favor não responda.\n\n" \
                        f"Caso haja divergências com a movimentação dos itens, favor entrar em contato pelo email " \
                        f"fat_maq@unisold.com.br. " \
                        f"Se houver algum problema com o recebimento de emails ou conflitos com o arquivo excel, " \
                        f"favor entrar em contato pelo email maquinas@unisold.com.br.\n\n"

            email_user = 'ti.ahcmaq@gmail.com'
            to = ['<maquinas@unisold.com.br>']
            password = 'poswxhqkeaacblku'

            return saudacao, msg_final, email_user, to, password

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def envia_email(self):
        try:
            from email.mime.multipart import MIMEMultipart
            from email.mime.text import MIMEText
            import smtplib

            cur = conecta.cursor()
            cur.execute(f"SELECT id, descricao, nome_usuario FROM ENVIA_PC "
                        f"where descricao = '{self.nome_computador}' "
                        f"and nome_usuario = '{self.username}';")
            select = cur.fetchall()

            if not select:
                saudacao, msg_final, email_user, to, password = self.mensagem_email()

                subject = f'Computadores'

                msg = MIMEMultipart()
                msg['From'] = email_user
                msg['Subject'] = subject

                body = f"{saudacao}\n\n- Nome PC: {self.nome_computador}\n- Nome Usuário: {self.username}" \
                       f"\n\n\n\n{msg_final}"

                msg.attach(MIMEText(body, 'plain'))

                text = msg.as_string()
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(email_user, password)

                server.sendmail(email_user, to, text)
                server.quit()

                cursor = conecta.cursor()
                cursor.execute(f"Insert into ENVIA_PC (id, descricao, nome_usuario) "
                               f"values (GEN_ID(GEN_ENVIA_PC_ID,1), '{self.nome_computador}', '{self.username}');")
                conecta.commit()

                print("email enviado")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def ultima_versao(self):
        try:
            import fdb

            conecta_robo = fdb.connect(database=r'C:\HallSys\db\Horus\Suzuki\ROBOZINHO.GDB',
                                       host='PUBLICO',
                                       port=3050,
                                       user='sysdba',
                                       password='masterkey',
                                       charset='ANSI')

            versao_app = self.versao[7:]

            cursor = conecta_robo.cursor()
            cursor.execute("SELECT versao FROM CONTROLE_VERSOES "
                           "where programa = 'ERP SUZUKI' "
                           "ORDER BY data_criacao DESC ROWS 1;")
            dados = cursor.fetchall()
            version = dados[0]
            versao_banco = version[0]
            print(versao_banco)

            resultado_comparacao = self.comparar_versoes(versao_app, versao_banco)
            if resultado_comparacao < 0:
                QTimer.singleShot(2000, self.chama_tela_atualizar)

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def comparar_versoes(self, versao_a, versao_b):
        try:
            partes_versao_a = list(map(int, versao_a.split('.')))
            partes_versao_b = list(map(int, versao_b.split('.')))

            while len(partes_versao_a) < len(partes_versao_b):
                print("while len(partes_versao_a) < len(partes_versao_b):")
                partes_versao_a.append(0)
            while len(partes_versao_b) < len(partes_versao_a):
                print("while len(partes_versao_b) < len(partes_versao_a):")
                partes_versao_b.append(0)

            for a, b in zip(partes_versao_a, partes_versao_b):
                if a < b:
                    return -1
                elif a > b:
                    return 1
            return 0

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def salva_versao(self):
        try:
            versao_app = self.versao[7:]

            agora = datetime.now()
            timestamp = agora.strftime('%Y-%m-%d %H:%M:%S')

            cur = conecta.cursor()
            cur.execute(f"SELECT id, descricao, versao FROM ENVIA_PC;")
            select = cur.fetchall()

            if select:
                for id_pc, desc_pc, versao in select:
                    if self.nome_computador == desc_pc and versao_app != versao:
                        cursor = conecta.cursor()
                        cursor.execute(f"UPDATE ENVIA_PC "
                                       f"SET VERSAO = '{versao_app}', ATUALIZACAO = '{timestamp}' "
                                       f"WHERE DESCRICAO = '{self.nome_computador}';")

                        conecta.commit()

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)

    def chama_tela_atualizar(self):
        try:
            caminho_exe = r'\\PUBLICO\Python\atualizador\atualizador.exe'
            try:
                subprocess.run(caminho_exe, check=True)
            except subprocess.CalledProcessError as e:
                print(f"Erro ao executar o arquivo .exe: {e}")

        except Exception as e:
            nome_funcao = inspect.currentframe().f_code.co_name
            exc_traceback = sys.exc_info()[2]
            self.trata_excecao(nome_funcao, str(e), self.nome_arquivo, exc_traceback)


if __name__ == '__main__':
    qt = QApplication(sys.argv)
    telamenu = TelaMenu()
    telamenu.show()
    sys.exit(qt.exec_())

import sys
from banco_dados.controle_erros import grava_erro_banco
import os
import inspect
import pyzipper
import traceback

nome_arquivo_com_caminho = inspect.getframeinfo(inspect.currentframe()).filename
nome_arquivo = os.path.basename(nome_arquivo_com_caminho)


def parametros_usuario(login_user, senha_user):
    try:
        conteudo = f"login:{login_user}\nsenha:{senha_user}"

        return conteudo

    except Exception as e:
        nome_funcao_trat = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        tb = traceback.extract_tb(exc_traceback)
        num_linha_erro = tb[-1][1]
        print(f'Houve um problema no arquivo: {nome_arquivo} na função: "{nome_funcao_trat}"\n'
              f'{e} {num_linha_erro}')
        grava_erro_banco(nome_funcao_trat, e, nome_arquivo, num_linha_erro)


def criar_ou_atualizar_zip(caminho_arquivo, login_user, senha_user, senha):
    try:
        conteudo = parametros_usuario(login_user, senha_user)

        with pyzipper.AESZipFile(caminho_arquivo, 'w', compression=pyzipper.ZIP_LZMA) as zipf:
            zipf.setpassword(senha.encode())
            zipf.setencryption(pyzipper.WZ_AES)
            zipf.writestr("dados.txt", conteudo)

    except Exception as e:
        nome_funcao_trat = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        tb = traceback.extract_tb(exc_traceback)
        num_linha_erro = tb[-1][1]
        print(f'Houve um problema no arquivo: {nome_arquivo} na função: "{nome_funcao_trat}"\n'
              f'{e} {num_linha_erro}')
        grava_erro_banco(nome_funcao_trat, e, nome_arquivo, num_linha_erro)


def ler_arquivo_zip(caminho_arquivo, senha):
    try:
        if not os.path.exists(caminho_arquivo):
            criar_ou_atualizar_zip(caminho_arquivo, "", "", senha)

        with pyzipper.AESZipFile(caminho_arquivo, 'r') as zipf:
            zipf.setpassword(senha.encode())
            conteudo = {}

            login_usuario = ""
            senha_usuario = ""

            for nome_arquivos in zipf.namelist():
                conteudo[nome_arquivos] = zipf.read(nome_arquivos).decode()

            for arquivo, informacoes_arquivo in conteudo.items():
                linhas = informacoes_arquivo.split('\n')

                if 'login:' in linhas[0]:
                    login_usuario = linhas[0].split('login:')[1].strip()

                if len(linhas) > 1 and 'senha:' in linhas[1]:
                    senha_usuario = linhas[1].split('senha:')[1].strip()

            return login_usuario, senha_usuario

    except Exception as e:
        nome_funcao_trat = inspect.currentframe().f_code.co_name
        exc_traceback = sys.exc_info()[2]
        tb = traceback.extract_tb(exc_traceback)
        num_linha_erro = tb[-1][1]
        print(f'Houve um problema no arquivo: {nome_arquivo} na função: "{nome_funcao_trat}"\n'
              f'{e} {num_linha_erro}')
        grava_erro_banco(nome_funcao_trat, e, nome_arquivo, num_linha_erro)

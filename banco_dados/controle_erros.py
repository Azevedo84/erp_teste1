from banco_dados.conexao import conecta_robo
import socket


def grava_erro_banco(nome_funcao, e, nome_arquivo, num_linha):
    try:
        nome_computador = socket.gethostname()

        msg_editada = str(e).replace("'", "*")
        msg_editada1 = msg_editada.replace('"', '*')

        msg_final = f"{msg_editada1} {num_linha}"

        cursor = conecta_robo.cursor()
        cursor.execute(f"Insert into ERROS (id, arquivo, funcao, mensagem, nome_pc) "
                       f"values (GEN_ID(GEN_ERROS_ID,1), '{nome_arquivo}', '{nome_funcao}', '{msg_final}', "
                       f"'{nome_computador}');")
        conecta_robo.commit()

    except Exception as e:
        print(e)

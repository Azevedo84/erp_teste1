from banco_dados.conexao import conectar_banco_nuvem


def proximo_registro_autoincremento(nome_tabela):
    conecta = conectar_banco_nuvem()
    try:
        cursor = conecta.cursor()
        cursor.execute(f"SHOW TABLE STATUS LIKE '{nome_tabela}';")
        result = cursor.fetchone()
        proximo_id = result[10]  # O valor do AUTO_INCREMENT está na 11ª posição (índice 10)

        return proximo_id

    finally:
        if 'conexao' in locals():
            conecta.close()

def proximo_registro_sem_autoincremento(nome_tabela):
    conecta = conectar_banco_nuvem()
    try:
        cursor = conecta.cursor()
        cursor.execute(f"SHOW TABLE STATUS LIKE '{nome_tabela}';")
        result = cursor.fetchone()
        proximo_id = result[10]  # O valor do AUTO_INCREMENT está na 11ª posição (índice 10)

        return proximo_id

    finally:
        if 'conexao' in locals():
            conecta.close()

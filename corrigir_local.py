from banco_dados.conexao import conectar_banco_nuvem


conecta_nuvem = conectar_banco_nuvem()
try:
    cursor = conecta_nuvem.cursor()
    cursor.execute(f"SELECT PRODUTO_ID, LOCAL_ESTOQUE_ID, armazem_id, localizacao FROM ESTOQUE "
                   f"where qtde > 0 and local_estoque_id = 9;")
    lista_completa = cursor.fetchall()

    if lista_completa:
        for i in lista_completa:
            print(i)
            prod_id, local_est_id, id_armazem, local = i

            cursor = conecta_nuvem.cursor()
            cursor.execute("UPDATE PRODUTO SET LOCALIZACAO = %s, ARMAZEM_ID = %s "
                           "WHERE ID_SIGER = %s;",
                (local, id_armazem, prod_id)
            )

            conecta_nuvem.commit()

finally:
    if 'conexao' in locals():
        conecta_nuvem.close()
from banco_dados.conexao import conectar_banco_nuvem


conecta_nuvem = conectar_banco_nuvem()
try:
    cursor = conecta_nuvem.cursor()
    cursor.execute(f"SELECT PRODUTO_ID, LOCAL_ESTOQUE_ID, localizacao FROM ESTOQUE "
                   f"where qtde < 0;")
    lista_completa = cursor.fetchall()

    if lista_completa:
        for i in lista_completa:
            print(i)
            prod_id, local_est_id, local = i

            locali = None

            cursor = conecta_nuvem.cursor()
            cursor.execute(
                "UPDATE ESTOQUE SET LOCALIZACAO = %s "
                "WHERE PRODUTO_ID = %s AND LOCAL_ESTOQUE_ID = %s;",
                (locali, prod_id, local_est_id)
            )

            conecta_nuvem.commit()

finally:
    if 'conexao' in locals():
        conecta_nuvem.close()
from banco_dados.conexao import conecta

cursor = conecta.cursor()
cursor.execute(f"select ordser.numero, "
               f"prod.codigo "
               f"from ordemservico as ordser "
               f"INNER JOIN produto prod ON ordser.produto = prod.id "
               f"where ordser.status = 'A';")
op_abertas = cursor.fetchall()

if op_abertas:
    for i in op_abertas:
        num_op, cod_prod = i
        print(i)

        cursor = conecta.cursor()
        cursor.execute(f"SELECT id, id_versao FROM produto WHERE codigo = '{cod_prod}';")
        dados_prod = cursor.fetchall()
        id_estrutura = dados_prod[0][1]
        print(id_estrutura)

        cursor.execute(f"UPDATE ordemservico SET id_estrutura = {id_estrutura} "
                       f"WHERE numero = {num_op};")
        conecta.commit()

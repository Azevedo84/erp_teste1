import fdb
from banco_dados.conexao import conectar_banco_nuvem

conecta = fdb.connect(database=r"C:\HallSys\db\Horus\Suzuki\ESTOQUE.GDB",
                      host='PUBLICO',
                      port=3050,
                      user='sysdba',
                      password='masterkey',
                      charset='ANSI')

cod_primeiro = "70110"
cod_ultimo = "100000"

cursor = conecta.cursor()
cursor.execute(f"SELECT produto_id, local_estoque, saldo "
               f"FROM SALDO_ESTOQUE;")
select_estoque = cursor.fetchall()

if select_estoque:
    for i in select_estoque:
        id_prod, local_est, saldo = i

        conecta_nuvem = conectar_banco_nuvem()
        try:
            cursor = conecta_nuvem.cursor()
            cursor.execute(f"SELECT * FROM ESTOQUE where produto_id = '{id_prod}' and local_estoque_id = {local_est};")
            lista_completa = cursor.fetchall()

            if not lista_completa:
                cursor = conecta.cursor()
                cursor.execute(f"SELECT codigo, localizacao FROM PRODUTO "
                               f"where id = '{id_prod}';")
                lista_produtos = cursor.fetchall()
                if lista_produtos:
                    for ii in lista_produtos:
                        codigo, local = ii

                        if local:
                            if "Z" in local:
                                print("Depósito Aço", local)

                            if "-" in local:
                                local_tete = local.find("-")
                                letra_inicial = local[:local_tete]



        finally:
            if 'conexao' in locals():
                conecta_nuvem.close()

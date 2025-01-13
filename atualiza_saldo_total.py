import fdb
from banco_dados.conexao import conectar_banco_nuvem

conecta = fdb.connect(database=r"C:\HallSys\db\Horus\Suzuki\ESTOQUE.GDB",
                      host='PUBLICO',
                      port=3050,
                      user='sysdba',
                      password='masterkey',
                      charset='ANSI')

cursor = conecta.cursor()
cursor.execute(f"SELECT DISTINCT prod.codigo, prod.quantidade "
               f"FROM produto as prod "
               f"WHERE prod.quantidade > 0;")
select_prod = cursor.fetchall()

for i in select_prod:
    print(i)
    codigo, qtde = i

    conecta_nuvem = conectar_banco_nuvem()
    try:
        cursor = conecta_nuvem.cursor()
        cursor.execute(
            "UPDATE PRODUTO SET SALDO_TOTAL = %s "
            "WHERE ID_SIGER = %s;",
            (qtde, codigo)
        )

        conecta_nuvem.commit()

    finally:
        if 'conexao' in locals():
            conecta_nuvem.close()
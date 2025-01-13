import fdb
from banco_dados.conexao import conectar_banco_nuvem

import os
from openpyxl import Workbook

conecta = fdb.connect(database=r"C:\HallSys\db\Horus\Suzuki\ESTOQUE.GDB",
                      host='PUBLICO',
                      port=3050,
                      user='sysdba',
                      password='masterkey',
                      charset='ANSI')


cursor = conecta.cursor()
cursor.execute(f"SELECT produto_id, local_estoque, saldo "
               f"FROM SALDO_ESTOQUE;")
select_estoque = cursor.fetchall()

lista_sem_local = []
lista_gravar = []

if select_estoque:
    for i in select_estoque:
        id_prod, local_est, saldo = i

        cursor = conecta.cursor()
        cursor.execute(f"SELECT codigo, descricao, localizacao FROM PRODUTO "
                       f"where id = '{id_prod}';")
        lista_produtos = cursor.fetchall()
        if lista_produtos:
            for ii in lista_produtos:
                codigo, descr, local = ii

                if saldo:
                    armazem_id = None

                    if not local:
                        dados = (codigo, descr, saldo, local)
                        lista_sem_local.append(dados)
                    else:
                        if "Z" in local:
                            armazem_id = 2
                            # print("Depósito Aço", local)

                        elif "-" in local:
                            local_tete = local.find("-")
                            letra_inicial = local[:local_tete]

                            armazem_id = 1
                            # print("Depósuito almox", local)

                        elif local == "A":
                            armazem_id = 5

                        else:
                            if local == "z":
                                armazem_id = 2
                            elif local == "PALLET AÇO":
                                armazem_id = 2
                            elif local == "AÇO":
                                armazem_id = 2

                            elif local == "ARMÁRIO":
                                armazem_id = 1
                            elif local == "PARADE ALMOX":
                                armazem_id = 1
                            elif local == "ARMÁRIOO" or local == "ARMARIO":
                                armazem_id = 1
                            elif local == "CHÃO ALMOX":
                                armazem_id = 1
                            elif local == "CABOS":
                                armazem_id = 1
                            elif local == "H":
                                armazem_id = 1
                            elif local == "FUNDOS ALMOX":
                                armazem_id = 1
                            elif local == "pallet" or local == "PALET" or local == "Palete":
                                armazem_id = 1
                            elif local == "I":
                                armazem_id = 1
                            elif local == "B":
                                armazem_id = 1
                            elif local == "C":
                                armazem_id = 1
                            elif local == "c":
                                armazem_id = 1
                            elif local == "PRATELEIRA FUNDOS":
                                armazem_id = 1

                            elif local == "PRODUCAO" or local == "PRODUÇÃO":
                                armazem_id = 6
                            elif local == "PALLET TORNO":
                                armazem_id = 6
                            elif local == "PALLET FRENTE ALMOX":
                                armazem_id = 6
                            elif local == "Setor Eletrica":
                                armazem_id = 6
                            elif local == "PLASA":
                                armazem_id = 6
                            elif local == "CAIXA URUGUAI":
                                armazem_id = 6
                            elif local == "ALCINDO":
                                armazem_id = 6
                            elif local == "TESTE DOBRADORAS":
                                armazem_id = 6

                            elif local == "TESTE PICOTE":
                                armazem_id = 6

                            elif local == "Gerencia":
                                armazem_id = 8

                            elif local == "ARMARIO IMPRESSORA":
                                armazem_id = 7

                            elif local == "OBSOLETO":
                                armazem_id = 3

                            elif "SUCATA" in local or "Sucata" in local:
                                armazem_id = 4
                            else:
                                armazem_id = 5

                        dadosss = (codigo, local_est, saldo, armazem_id, local)
                        lista_gravar.append(dadosss)

if lista_sem_local:
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    file_path = os.path.join(desktop_path, f"rel sem local.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "Problemas"

    headers = [
        "Código", "Descrição", "Saldo", "Local"]
    ws.append(headers)

    for ii in lista_sem_local:
        cod_p, descr_p, saldo_p, local_p = ii
        ws.append([cod_p, descr_p, saldo_p, local_p])

    # Salvar o arquivo Excel na área de trabalho
    wb.save(file_path)
    print(f"Arquivo Excel salvo em: {file_path}")

if lista_gravar:
    for iii in lista_gravar:
        codigo, local_est, saldo, armazem_id, local = iii

        conecta_nuvem = conectar_banco_nuvem()
        try:
            cursor = conecta_nuvem.cursor()
            cursor.execute(f"SELECT * FROM ESTOQUE "
                           f"where produto_id = '{codigo}' "
                           f"and local_estoque_id = {local_est};")
            lista_completa = cursor.fetchall()

            if not lista_completa:
                print(iii)
                if saldo > 0:
                    query = """
                            INSERT INTO ESTOQUE (PRODUTO_ID, LOCAL_ESTOQUE_ID, QTDE, ARMAZEM_ID, LOCALIZACAO) 
                            VALUES (%s, %s, %s, %s, %s);
                            """

                    values = (codigo, local_est, saldo, armazem_id, local )

                    cursor = conecta_nuvem.cursor()
                    cursor.execute(query, values)

                    conecta_nuvem.commit()

                    print(f"Produto salvo com sucesso! {iii}")
                    print("\n")
                else:
                    query = """
                            INSERT INTO ESTOQUE (PRODUTO_ID, LOCAL_ESTOQUE_ID, QTDE) 
                            VALUES (%s, %s, %s);
                            """

                    values = (codigo, local_est, saldo)

                    cursor = conecta_nuvem.cursor()
                    cursor.execute(query, values)

                    conecta_nuvem.commit()

                    print(f"Produto salvo com sucesso! {iii}")
                    print("\n")

        finally:
            if 'conexao' in locals():
                conecta_nuvem.close()

import fdb
from banco_dados.conexao import conectar_banco_nuvem
import re
import os
from openpyxl import Workbook

conecta = fdb.connect(database=r"C:\HallSys\db\Horus\Suzuki\ESTOQUE.GDB",
                      host='PUBLICO',
                      port=3050,
                      user='sysdba',
                      password='masterkey',
                      charset='ANSI')

lista_problema = []

cod_primeiro = "21241"
cod_ultimo = "100000"

cursor = conecta.cursor()
cursor.execute(f"SELECT DISTINCT prod.codigo, prod.descricao, prod.descricaocomplementar, "
               f"prod.obs, prod.unidade, prod.ncm, prod.conjunto, "
               f"prod.quantidademin, prod.obs2 "
               f"FROM produto as prod "
               f"INNER JOIN movimentacao as mov ON prod.id = mov.produto "
               f"WHERE prod.codigo > {cod_primeiro} and prod.codigo < {cod_ultimo};")
select_prod = cursor.fetchall()

if select_prod:
    for i in select_prod:
        cod, descr, compl, ref, um, ncm, conjunto, qtde_min, obs = i

        print(i)

        conecta_nuvem = conectar_banco_nuvem()
        try:
            cursor = conecta_nuvem.cursor()
            cursor.execute(f"SELECT * FROM PRODUTO where id_siger = '{cod}';")
            lista_completa = cursor.fetchall()

            if not lista_completa:
                if not compl:
                    compl = None

                    linha_compl = 0
                else:
                    linha_compl = len(compl)

                if not ref:
                    ref = None
                    linha_ref = 0
                else:
                    linha_ref = len(ref)

                if not descr:
                    linha_descr = 0
                else:
                    linha_descr = len(descr)

                if not ncm:
                    ncm = None
                else:
                    ncm = re.sub(r'\D', '', ncm)

                if not obs:
                    obs = None

                if not qtde_min:
                    qtde_min = 0

                if linha_ref > 20 or linha_compl > 30 or linha_descr > 30:
                    dados = (cod, descr, compl, ref, um, conjunto, qtde_min, ncm, obs, 'A')
                    print("PROBLEMA ENCONTRADO!!")
                    lista_problema.append(dados)
                else:
                    query = """
                    INSERT INTO PRODUTO (
                        ID_SIGER, DESCRICAO, COMPLEMENTAR, REFERENCIA, UM,
                        CLASSIFICACAO_PRODUTO_ID, QTDE_MINIMA, NCM, OBS, ATIVO
                    ) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
                    """

                    values = (
                        cod, descr, compl, ref, um,
                        conjunto, qtde_min, ncm, obs, 'A'
                    )

                    cursor = conecta_nuvem.cursor()
                    cursor.execute(query, values)

                    conecta_nuvem.commit()

                    print(f"Produto salvo com sucesso! {i}")
                    print("\n")

        finally:
            if 'conexao' in locals():
                conecta_nuvem.close()

if lista_problema:
    # Localizar automaticamente a área de trabalho
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    file_path = os.path.join(desktop_path, f"rel {cod_primeiro} a {cod_ultimo}.xlsx")

    # Criar um novo workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Problemas"

    # Cabeçalhos para o Excel
    headers = [
        "Código", "Descrição", "Complemento", "Referência", "UM", "Conjunto", "Mínima", "NCM", "Observação", "Ativo"
    ]
    ws.append(headers)

    # Preencher o Excel com os dados da lista_problema
    if lista_problema:
        for ii in lista_problema:
            cod_p, descr_p, compl_p, ref_p, um_p, conjunto_p, qtde_min_p, ncm_p, obs_p, ativo_p = ii
            ws.append([cod_p, descr_p, compl_p, ref_p, um_p, conjunto_p, qtde_min_p, ncm_p, obs_p, ativo_p])
            print("PROBLEMA", ii)

    # Salvar o arquivo Excel na área de trabalho
    wb.save(file_path)
    print(f"Arquivo Excel salvo em: {file_path}")
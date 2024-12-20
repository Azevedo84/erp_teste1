import mysql.connector
from mysql.connector import errorcode
import config
import fdb

try:
    conecta_robo = fdb.connect(database=r'C:\HallSys\db\Horus\Suzuki\ROBOZINHO.GDB',
                               host='PUBLICO',
                               port=3050,
                               user='sysdba',
                               password='masterkey',
                               charset='ANSI')

except Exception as e:
    print(e)


def conectar_banco_nuvem():
    try:
        conexao = mysql.connector.connect(
            host=config.DB_HOST,
            user=config.DB_USER,
            password=config.DB_PASSWORD,
            database=config.DB_DATABASE
        )
        return conexao

    except mysql.connector.Error as err:
        if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
            msgerro = "Usuário ou senha incorretos!"
        elif err.errno == errorcode.ER_BAD_DB_ERROR:
            msgerro = "Banco de Dados não existe!"
        elif err.errno == errorcode.CR_CONN_HOST_ERROR:
            msgerro = "Endereço TCP/IP não encontrado!"
        else:
            msgerro = err
        raise Exception(f"Erro ao conectar ao banco de dados: {msgerro}")

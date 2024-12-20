import pyzipper


def criar_ou_atualizar_zip(caminho_arquivo, host, user, database, passwaord, senha):
    conteudo = f"DB_HOST = {host}\n" \
               f"DB_USER = {user}\n" \
               f"DB_DATABASE = {database}\n" \
               f"DB_PASSWORD = {passwaord}" \

    with pyzipper.AESZipFile(caminho_arquivo, 'w', compression=pyzipper.ZIP_LZMA) as zipf:
        zipf.setpassword(senha.encode())
        zipf.setencryption(pyzipper.WZ_AES)  # Criptografia forte (AES)
        # Criar ou substituir o arquivo "dados.txt" com as informações
        zipf.writestr("dados.txt", conteudo)


def ler_arquivo_zip(caminho_arquivo, senha):
    with pyzipper.AESZipFile(caminho_arquivo, 'r') as zipf:
        zipf.setpassword(senha.encode())
        conteudo = {}

        host = ""
        user = ""

        # Lê o conteúdo de todos os arquivos no zip
        for nome_arquivo in zipf.namelist():
            conteudo[nome_arquivo] = zipf.read(nome_arquivo).decode()

        # Procurar o usuário e a senha nos arquivos
        for arquivo, informacoes_arquivo in conteudo.items():
            # Dividir o conteúdo em linhas
            linhas = informacoes_arquivo.split('\n')

            if 'DB_HOST = ' in linhas[0]:
                host = linhas[0].split('DB_HOST = ')[1].strip()

            # Se a segunda linha contém a senha, pegamos a senha
            if len(linhas) > 1 and 'DB_USER = ' in linhas[1]:
                user = linhas[1].split('DB_USER = ')[1].strip()

        return host, user


# Exemplo de uso
senha = "minha_senha_forte"

host = "db-erp-suzuki.mysql.uhserver.com"
user = "boby_ahc"
database = "db_erp_suzuki"
passwaord = "elvT@R2D-t7S4"

arquivo = "config.zip"

criar_ou_atualizar_zip(arquivo , host, user, database, passwaord, senha)

# Ler conteúdo do ZIP
login_usuario, senha_usuario = ler_arquivo_zip(arquivo , senha)
print("loco", login_usuario)
print("senho", senha_usuario)


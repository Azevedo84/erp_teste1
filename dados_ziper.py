import pyzipper


def criar_ou_atualizar_zip(caminho_arquivo, login_user, senha_user, senha):
    conteudo = f"login:{login_user}\nsenha:{senha_user}"

    with pyzipper.AESZipFile(caminho_arquivo, 'w', compression=pyzipper.ZIP_LZMA) as zipf:
        zipf.setpassword(senha.encode())
        zipf.setencryption(pyzipper.WZ_AES)  # Criptografia forte (AES)
        # Criar ou substituir o arquivo "dados.txt" com as informações
        zipf.writestr("dados.txt", conteudo)


def ler_arquivo_zip(caminho_arquivo, senha):
    with pyzipper.AESZipFile(caminho_arquivo, 'r') as zipf:
        zipf.setpassword(senha.encode())
        conteudo = {}

        # Variáveis para armazenar o usuário e a senha
        login_usuario = ""
        senha_usuario = ""

        # Lê o conteúdo de todos os arquivos no zip
        for nome_arquivo in zipf.namelist():
            conteudo[nome_arquivo] = zipf.read(nome_arquivo).decode()

        # Procurar o usuário e a senha nos arquivos
        for arquivo, informacoes_arquivo in conteudo.items():
            # Dividir o conteúdo em linhas
            linhas = informacoes_arquivo.split('\n')

            # Se a primeira linha contém o login, pegamos o usuário
            if 'login:' in linhas[0]:
                login_usuario = linhas[0].split('login:')[1].strip()

            # Se a segunda linha contém a senha, pegamos a senha
            if len(linhas) > 1 and 'senha:' in linhas[1]:
                senha_usuario = linhas[1].split('senha:')[1].strip()

        return login_usuario, senha_usuario


# Exemplo de uso
senha = "minha_senha_forte"

login_user = "admin"
senha_user = "1234"

# criar_ou_atualizar_zip("config.zip", login_user, senha_user, senha)

# Ler conteúdo do ZIP
login_usuario, senha_usuario = ler_arquivo_zip("dados_seguros_aes.zip", senha)
print("loco", login_usuario)
print("senho", senha_usuario)


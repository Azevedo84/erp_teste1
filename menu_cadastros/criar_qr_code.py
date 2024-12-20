import qrcode

# Dados para o QR Code
data = 'Código: 16334\n' \
       'Descrição: CONJUNTO PONTO BOB ROLAMENT V3\n' \
       'D. Complementar: \n' \
       'Referência: D 21.01.03.00\n' \
       'UM: UN\n' \
       'Tipo de Material: CONJUNTO\n' \
       'Conjunto: Produtos Acabados\n' \
       'Local: F-3\n' \
       'NCM: 8477.90.00'

# Criar um objeto QR Code
qr = qrcode.QRCode(
    version=1,
    error_correction=qrcode.constants.ERROR_CORRECT_L,
    box_size=10,
    border=4,
)

# Adicionar dados ao QR Code
qr.add_data(data)
qr.make(fit=True)

# Criar uma imagem a partir do QR Code
img = qr.make_image(fill='black', back_color='white')

# Salvar a imagem
img.save('qrcode.png')

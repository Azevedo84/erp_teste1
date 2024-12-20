import requests
from datetime import timedelta
from dateutil.easter import easter

ano = 2025


def obter_feriados(ano):
    url = f"https://brasilapi.com.br/api/feriados/v1/{ano}"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Erro ao buscar feriados: {response.status_code}")
        return []


feriados = obter_feriados(ano)

print("Feriados Nacionais e Estaduais de 2024:")
for feriado in feriados:
    print(f"{feriado['date']}: {feriado['name']} ({feriado['type']})")


def calcular_feriados_moveis(ano):
    pascoa = easter(ano)
    carnaval = pascoa - timedelta(days=47)
    sexta_santa = pascoa - timedelta(days=2)
    corpus_christi = pascoa + timedelta(days=60)
    ascensao = pascoa + timedelta(days=39)

    return {
        "Carnaval": carnaval,
        "Sexta-feira Santa": sexta_santa,
        "Domingo de Páscoa": pascoa,
        "Ascensão do Senhor": ascensao,
        "Corpus Christi": corpus_christi,
    }

print("\n\n")
feriados_moveis = calcular_feriados_moveis(ano)

print(f"Feriados móveis em {ano}:")
for nome, data in feriados_moveis.items():
    print(f"{nome}: {data.strftime('%d/%m/%Y')}")

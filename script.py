import datetime
import holidays

# Defina a data para a qual você deseja saber o dia da semana
data = datetime.date(2023, 10, 21)  # Substitua com a data desejada (ano, mês, dia)

# Use a função weekday() para obter o dia da semana (segunda-feira = 0, domingo = 6)
dia_da_semana = data.weekday()

# Crie uma lista com os nomes dos dias da semana
dias_da_semana = ["Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado", "Domingo"]

# Escolha o país para verificar os feriados
pais = "BR"  # Substitua pelo código do país desejado (BR para Brasil, US para Estados Unidos, etc.)

# Verifique se a data é um feriado
if data in holidays.CountryHoliday(pais, observed=True):
    print(f"A data {data} é um feriado em {pais}.")
else:
    print(f"A data {data} não é um feriado em {pais}.")

# Imprima o resultado
print(f"A data {data} corresponde a {dias_da_semana[dia_da_semana]}")
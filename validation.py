import holidays
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def remove_weekends_and_holidays(df):
    
    df['Data2'] = pd.to_datetime(df['Data'], format='%d/%m/%Y')
    
    df['Remover'] = df['Data2'].apply(validation_weekend_and_holidays)

    # Remova os registros que atendem à condição (Remover = True)
    df = df[~df['Remover']]

    # Remova a coluna 'Remover' se não for mais necessária
    df = df.drop('Remover', axis=1)
    df = df.drop('Data2', axis=1)
    
def validation_weekend_and_holidays(date):
    
    dia_da_semana = date.weekday()

    pais = "BR"

    # Verifique se a data é um feriado
    if date in holidays.CountryHoliday(pais, observed=True) or (dia_da_semana == 5 or dia_da_semana == 6):
        return True
    else:
        return False

def adjust_lunch_time(df, hora_almoco):
    # Separe 'Hour' e 'Minute' da coluna 'Total' e converta em inteiros
    df[['Hour', 'Minute']] = df['Total'].str.split(":", n=1, expand=True).astype(int)

    # Faça o ajuste necessário com a hora do almoço
    df['Hour'] = df['Hour'] - hora_almoco

    # Converta 'Hour' e 'Minute' de volta em uma string no formato HH:MM e atribua a 'Total'
    df['Total'] = df['Hour'].astype(str) + ":" + df['Minute'].astype(str)

def generate_excel(df):
    # Salve o DataFrame em um novo arquivo Excel
    df.to_excel('Output.xlsx', sheet_name='Data', index=True)

    # Carregue o arquivo Excel com a nova planilha
    wb = load_workbook('Output.xlsx')

    # Obtenha a planilha 'Data'
    ws = wb['Data']

    # Exclua a terceira linha da planilha para preservar os cabeçalhos e a primeira coluna
    ws.delete_rows(3)

    # Defina uma cor de preenchimento vermelho
    red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

    # Itere pelas células na planilha e verifique se o valor é "9:00"
    for row in ws.iter_rows(min_row=2, min_col=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            if pd.notna(cell.value) and ':' in cell.value:
                    parts = cell.value.split(':')
                    if len(parts) > 1:
                        horas = int(parts[0])
                        minutos = int(parts[1])
                        if horas > 5 or (horas == 5 and minutos > 17):
                            cell.fill = red_fill
    # Salve o arquivo Excel atualizado
    wb.save('Output.xlsx')
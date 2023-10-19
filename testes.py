import pandas as pd
from validation import remove_weekends_and_holidays, adjust_lunch_time, generate_excel

def processar_arquivo_excel(arquivo_excel, colunas_desejadas, hora_almoco):
    df = pd.read_excel(
        arquivo_excel,
        usecols=colunas_desejadas
    )

    df = df.drop_duplicates(subset=['Matricula', 'Nome', 'Data'], keep='last')
    
    adjust_lunch_time(df, hora_almoco)
    
    remove_weekends_and_holidays(df)
    
    df['indice'] = range(len(df))

    # Defina a nova coluna como o índice do DataFrame
    df.set_index('indice', inplace=True)
    
    pivot_table = pd.pivot_table(df, values=['Total'], index=['Nome'], columns=['Data'], aggfunc='sum')
    
    pivot_table.index.name = None
    pivot_table.columns.name = None
    
    pivot_table = pivot_table.fillna('0:00')
    
    generate_excel(pivot_table)
    
    print(df)
    return df.head(1000)


# ------------------------------------------------------------------------------------------------
# Chame a função com os parâmetros necessários
arquivo_excel = 'EMTHOS.xlsx'
colunas_desejadas = ['Matricula', 'Nome', 'Data', 'Total']
hora_almoco = 1

df_resultado = processar_arquivo_excel(arquivo_excel, colunas_desejadas, hora_almoco)
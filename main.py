import os 
import pandas as pd
import numpy as np
import openpyxl
from datetime import date
from dateutil.relativedelta import relativedelta
from workadays import workdays as wd
import warnings
warnings.filterwarnings("ignore", category=FutureWarning) 

def main():
    valor_passagem = 4.4
    dia_contagem = 7
    hoje = date.today()
    
    print('\n== Relatório para gastos em condução ==')
    print('\nOpções:')
    print('1. Selecionar mês atual')
    print('2. Selecionar mês diferente')

    opc = int(input('\nEscolha a opção desejada: '))

    mes = 0
    if opc == 2:
        print('\n1. Janeiro - Fevereiro')
        print('2. Fevereiro - Março')
        print('3. Março - Abril')
        print('4. Abril - Maio')
        print('5. Maio - Junho')
        print('6. Junho - Julho')
        print('7. Julho - Agosto')
        print('8. Agosto - Setembro')
        print('9. Setembro - Outubro')
        print('10. Outubro - Novembro')
        print('11. Novembro - Dezembro')
        print('12. Dezembro - Janeiro')
        mes = int(input('\nSelecione o mês desejado: '))
        data_inicio = date(hoje.year, mes, dia_contagem)
    
    if mes == 0:
        data_inicio = date(hoje.year, hoje.month, dia_contagem)
        if hoje.day < dia_contagem:
            data_inicio = data_inicio - relativedelta(months=1)

    data_fim = data_inicio + relativedelta(months=1) - relativedelta(days=1)    
    mes = data_fim.month

    df = pd.DataFrame()
    df['Data'] = pd.date_range(data_inicio, data_fim)
    df['Data'] = pd.to_datetime(df['Data'])

    nome_da_semana = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']

    df['Dia'] = df['Data'].apply(lambda x: nome_da_semana[x.day_of_week])

    indices_dias_uteis = df[df.Data.apply(lambda x: wd.is_workday(x, country='BR', state='SP'))].index
    df.loc[indices_dias_uteis, 'Passagem de ida'] = valor_passagem
    df.loc[indices_dias_uteis, 'Passagem de volta'] = valor_passagem
    df['Total acumulado'] = df['Passagem de ida'].replace(np.nan, 0).cumsum() + df['Passagem de volta'].replace(np.nan, 0).cumsum()

    total_row = pd.Series([np.nan, 'Total', 'R$ %.2f' % df['Passagem de ida'].sum(), 'R$ %.2f' % df['Passagem de volta'].sum(), 'R$ %.2f' % df.iloc[-1,-1]])
    df['Total acumulado'] = df['Total acumulado'].apply(lambda x: 'R$ %.2f' %x)
    df = df.replace({np.nan: '-', valor_passagem: 'R$ %.2f' %valor_passagem})
    print('\nTotal previsto para o mês: ' + df.loc[df.index[-1],'Total acumulado'])

    if opc == 1:
        print('Total até a data de hoje ({0}): {1}'.format(hoje.strftime('%d/%m/%Y'), df.loc[df.Data.dt.date == hoje, 'Total acumulado'].values[0])) 
    
    df['Data'] = df['Data'].dt.strftime('%d/%m/%Y')
    path_out = os.getcwd()

    df.to_excel(path_out + '\\Condução_{}.xlsx'.format(date(hoje.year,mes,hoje.day).strftime('%b-%Y')))
    print('Planilha de gastos foi gerada!! Local: ' + path_out)
    input('\nTecle enter para fechar...')

if __name__ == '__main__':
    main()
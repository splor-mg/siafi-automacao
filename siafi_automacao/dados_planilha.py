import pandas as pd

df = pd.read_excel('/home/guilhermemelof/code/splor-mg/siafi-automacao/data/teste_automacao.xlsx', sheet_name='Remanejamento Cota Orçamentaria')
df = df.dropna(how='all')  # remove linhas completamente vazias
df = df.sort_values(by=['Anular', 'UO_COD'], ascending=[True, True]) # ordena por anulação e depois por UO

for _, row in df.iterrows():
    uo      = str(int(row['UO_COD']))
    grupo   = str(int(row['Grupo']))
    iag     = str(int(row['IAG']))
    fonte   = str(int(row['Fonte']))
    procedencia = str(int(row['IPU']))
    tipo_global = row['GLOBAL'] if pd.notna(row['GLOBAL']) else '0'
    tipo_amarrado = str(int(row['AMARRADO'])) if pd.notna(row['AMARRADO']) else '0'
    if pd.notna(row['AMARRADO']):
        amarrado = str(int(row['AMARRADO']))
        elemento = amarrado[:2]   # dois primeiros digitos
        item = amarrado[2:]       # dois ultimos digitos
    else:
        elemento = '0'
        item = '0'
    valor_anulacao = int(round(float(row['Anular']), 2) * 100) if pd.notna(row['Anular']) else 0
    valor_aprovacao = int(round(float(row['Aprovar']), 2) * 100) if pd.notna(row['Aprovar']) else 0
     
    ##Definição do valor a ser preenchido, dependendo se é anulação ou aprovação
    if pd.notna(row['Anular']):
        valor = int(round(float(row['Anular']), 2) * 100)
    else:
        valor = int(round(float(row['Aprovar']), 2) * 100)

    if valor_anulacao != 0:
        print(f"realizando procedimento de anulação")
    elif valor_aprovacao != 0:
        print(f"realizando procedimento de aprovação")
            
    if tipo_global == 'x':
        print(f"Processando UO: {uo}, Grupo: {grupo}, IAG: {iag}, Fonte: {fonte}, Procedencia: {procedencia}, Valor Anulação: {valor_anulacao}")
    elif tipo_amarrado != '0':
        print(f"Processando UO: {uo}, Grupo: {grupo}, IAG: {iag}, Fonte: {fonte}, Procedencia: {procedencia}, Valor Anulação: {valor_anulacao}")

    # -------------------- exemplo para orquestrar o fluxo --------------------

    if valor_anulacao != 0:
        anular()
    elif valor_aprovacao != 0:
        aprovar()
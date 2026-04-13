from py3270 import Emulator
from datetime import datetime
import time
import pandas as pd

month = datetime.today().strftime("%m")
em = Emulator(visible=True)
em.connect('bhmvsb.prodemge.gov.br')
em.wait_for_field()

##preparacao e tratamento da planilha
df = pd.read_excel('/home/guilhermemelof/code/splor-mg/siafi-automacao/data/teste_automacao.xlsx', sheet_name='Remanejamento Cota Orçamentaria')
df = df.dropna(how='all')  # remove linhas completamente vazias
df = df.sort_values(by=['Anular', 'UO_COD'], ascending=[True, True]) # ordena por anulação e depois por UO

# Preenche os dados de login
em.fill_field(19, 13, 'simg', 7)
em.fill_field(20, 13, 'm755791', 7)
em.fill_field(21, 13, 'qws123', 7)
em.send_enter()

# Loop: navega pelas telas até encontrar a mensagem de sucesso
max_tentativas = 10
tentativas = 0

while tentativas < max_tentativas:
    time.sleep(1)

    try:
        em.wait_for_field()

        # Tela COM campo editável — verifica se é a tela de sucesso
        if em.string_found(1, 13, 'Logon executado com sucesso'):
            print("Login realizado com sucesso!")
            break

        else:
            # Tela com campo editável, mas ainda não é a de sucesso
            print(f"Tentativa {tentativas + 1} - tela intermediária, avançando...")
            em.send_enter()

    except ValueError:
        # Tela SEM campo editável — é a tela de aviso, só dá Enter e segue
        print(f"Tentativa {tentativas + 1} - tela de aviso detectada, passando...")
        em.send_enter()

    tentativas += 1

if tentativas == max_tentativas:
    print("Não foi possível fazer login após várias tentativas.")
    em.terminate()

em.fill_field(1, 2, 'simg', 4)
em.send_enter()

##nova tela buscando login...
max_tentativas = 10
tentativas = 0

while tentativas < max_tentativas:
    time.sleep(1)

    try:
        em.wait_for_field()

        # Tela COM campo editável — verifica se é a tela de sucesso
        if em.string_found(22, 11, 'Unidade Executora'):
            print("Texto encontrado")
            break

        else:
            # Tela com campo editável, mas ainda não é a de sucesso
            print(f"Tentativa {tentativas + 1} - tela intermediária, avançando...")
            em.send_enter()

    except ValueError:
        # Tela SEM campo editável — é a tela de aviso, só dá Enter e segue
        print(f"Tentativa {tentativas + 1} - tela de aviso detectada, passando...")
        em.send_enter()

    tentativas += 1

if tentativas == max_tentativas:
    print("Não foi possível fazer login após várias tentativas.")
    em.terminate()

#Entrar com a Unidade Executora
em.fill_field(22, 30, '1500008', 7)
em.send_enter()
em.wait_for_field()
#Entrar em 03 - Movimentacao Orcamentaria
em.fill_field(21, 19, '03', 2)
em.send_enter()
em.wait_for_field()
#Entrar em 02 - Aprovacao de Cota Orcamentaria
em.fill_field(21, 19, '02', 2)
em.send_enter()
em.wait_for_field()

##leitura das linhas da planilha e preenchimento dos dados no SIAFI
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


## Verifica se é anulação ou aprovação e preencche 03-1 para aprovação e 04-1 para anulação
    if valor_anulacao != 0:
        ##Anulação de cota orçamentária
        em.fill_field(21, 19, '04', 2)
        em.fill_field(21, 41, '1', 1)
        em.send_enter()
        em.wait_for_field()
    elif valor_aprovacao != 0:
        ##Aprovação de cota orçamentária
        em.fill_field(21, 19, '03', 2)
        em.fill_field(21, 41, '1', 1)
        em.send_enter()
        em.wait_for_field()
            
            ## Verifica se é global ou amarrado e preenche os campos correspondentes
            if tipo_global == 'x':
                #Aprovação/Anulação GLOBAL
                em.fill_field(8, 52, month, 2) # mes
                em.fill_field(9, 52, 'x', 1) # global
                em.fill_field(11, 52, fonte, 2) # fonte e procendencia
                em.fill_field(11, 54, procedencia, 1) # fonte e procendencia
                em.fill_field(12, 52, uo, 4) # UO
                em.fill_field(13, 52, grupo, 1) # grupo de despesa
                em.fill_field(13, 54, iag, 1) # IAG
                em.send_enter()
                em.wait_for_field()
            elif tipo_amarrado != '0':
                #Aprovação/Anulação AMARRADO
                em.fill_field(8, 52, month, 2) # mes
                em.fill_field(10, 52, 'x', 1) # amarrado
                em.fill_field(11, 52, fonte, 2) # fonte e procendencia
                em.fill_field(11, 54, procedencia, 1) # fonte e procendencia
                em.fill_field(12, 52, uo, 4) # UO
                em.fill_field(13, 52, grupo, 1) # grupo de despesa
                em.fill_field(13, 54, iag, 1) # IAG
                em.fill_field(16, 52, elemento, 2) # elemento
                em.fill_field(16, 54, item, 2) # item
                em.send_enter()
                em.wait_for_field()


    #digitar as informações das ações e valores...
    em.fill_field(16, 16, '4527', 4) # ação
    em.fill_field(16, 21, '0001', 4)
    em.fill_field(16, 33, '1000', 15) # valor
    em.send_enter()
    em.wait_for_field()

    em.fill_field(11, 11, 'Remanejamento realizado conforme solicitado', 60) # ação
    em.send_enter()
    em.wait_for_field()

    em.send_pf(5)  # envia F5
    em.wait_for_field()
    em.send_pf(5)  # envia F5

    retorno = em.string_get(1, 1, 80).strip()
    print(f"SIAFI retornou: {retorno}")



input("Pressione ENTER para fechar...")

em.terminate()
import os
import shutil
from py3270 import Emulator
from datetime import datetime
import pandas as pd
import openpyxl
import time
from fluxo_anular_desc import anulacao
from fluxo_aprovar_desc import aprovacao

sistema = 'simg'
usuario = 'm1241897'
senha = 'Pc0987'
unidade_executora = '1510010'
month = datetime.today().strftime("%m")

# Definição dos CAMINHOS
# Cópia local de trabalho — onde o robô vai ler e salvar durante a execução... ele é criado a partir do original do OneDrive e só é salvo no final, para evitar conflitos de acesso com o OneDrive
CAMINHO_LOCAL     = '/home/guilhermemelof/code/splor-mg/siafi-automacao/data/copia.xlsx'

#Nome da aba na planilha Excel onde estão os dados a serem processados
SHEET_NAME = 'Descentraliza Cota Orcamentaria'

em = Emulator(visible=True) ##caso queira que a tela apareça utilize visible=True
em.connect('bhmvsb.prodemge.gov.br')
em.wait_for_field()

# Preenche os dados de login
em.fill_field(19, 13, sistema, 7)
em.fill_field(20, 13, usuario, 8)
em.fill_field(21, 13, senha, 7)
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

    except:
        print(f"Tentativa {tentativas + 1} - tela de aviso detectada, passando...")
        em.send_enter()

    tentativas += 1

if tentativas == max_tentativas:
    print("Não foi possível fazer login após várias tentativas.")
    em.terminate()

em.fill_field(1, 2, sistema, 4)
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

    except:
        # Tela SEM campo editável — é a tela de aviso, só dá Enter e segue
        print(f"Tentativa {tentativas + 1} - tela de aviso detectada, passando...")
        em.send_enter()

    tentativas += 1

if tentativas == max_tentativas:
    print("Não foi possível fazer login após várias tentativas.")
    em.terminate()

#Entrar com a Unidade Executora
em.fill_field(22, 30, unidade_executora, 7)
em.send_enter()
em.wait_for_field()
# Fim do login

#Entrar em 03 - Movimentacao Orcamentaria
em.fill_field(21, 19, '03', 2)
em.send_enter()
em.wait_for_field()

#Entrar em 03 - Descentralizacao de Cota Orcamentaria
em.fill_field(21, 19, '03', 2)
em.send_enter()
em.wait_for_field()

# Leitura da planilha e processamento dos dados

# -----------------------------------------------------------------------
# ETAPA 3: leitura da cópia LOCAL (não do OneDrive)
# -----------------------------------------------------------------------
df = pd.read_excel(CAMINHO_LOCAL, sheet_name=SHEET_NAME)
df = df.dropna(how='all')  # remove linhas completamente vazias
df = df.sort_values(by=['Orientacao'], ascending=[True]) # ordena por anulação
df = df.reset_index(drop=False)

# O loop agora usa "for idx, row" em vez de "for _, row".
# O idx é o índice real da linha no DataFrame e é necessário para que o
# df.at[idx, 'Progresso'] grave o retorno na linha correta em memória.
for idx, row in df.iterrows():
    data_row = {}
    data_row['month']   = month
    data_row['orientacao']      = str((row['Orientacao']))
    data_row['ue']   = str(int(row['UE_Beneficiada']))
    data_row['tipo']     = str((row['Tipo de Descentralizacao']))
    data_row['fonte']   = str(int(row['Fonte']))
    data_row['procedencia'] = str(int(row['Procedencia']))
    data_row['dea'] = str((row['Elemento 92']))
    data_row['acao'] = str(int(row['Acao']))
    data_row['natureza_despesa'] = str(int(row['Natureza_Despesa_Elemento']))

    if pd.notna(row['Natureza_Despesa_Elemento']):
        data_row['categoria'] = data_row['natureza_despesa'][:1]   # dois primeiros digitos
        data_row['grupo'] = data_row['natureza_despesa'][1:2]       # segundo digito
        data_row['modalidade'] = data_row['natureza_despesa'][2:4]  # terceiro
        data_row['elemento'] = data_row['natureza_despesa'][4:6]    # dois ultimos digitos
    else:
        data_row['categoria'] = '0'
        data_row['grupo'] = '0'
        data_row['modalidade'] = '0'
        data_row['elemento'] = '0'

    data_row['item'] = str(int(row['Item'])) if pd.notna(row['Item']) else '0'
    data_row['uo_financiadora'] = str(int(row['UO_Financiadora'])) if pd.notna(row['UO_Financiadora']) else '0'
    data_row['iag'] = str(int(row['IAG']))
    data_row['valor'] = int(round(float(row['Valor']), 2) * 100)

    # Criação da Variável de Retorno para armazenar o resultado do processamento de cada linha
    retorno = ''

    if data_row['orientacao'] == 'Anular':
        print(f"Realizando procedimento de anulação")
    elif data_row['orientacao'] == 'Aprovar':
        print(f"Realizando procedimento de aprovação")

    if data_row['tipo'] == 'Global':        
        print(f" UE = {data_row['ue']}, Tipo = {data_row['tipo']}, Fonte e IPU = {data_row['fonte']}.{data_row['procedencia']}, UO Financiadora = {data_row['uo_financiadora']}, Ação = {data_row['acao']}, Natureza de Despesa = {data_row['natureza_despesa']}, Valor: {data_row['valor']}")
    elif data_row['tipo'] == 'Amarrado':
        print(f" UE = {data_row['ue']}, Tipo = {data_row['tipo']}, Fonte e IPU = {data_row['fonte']}.{data_row['procedencia']}, Uo Financiadora = {data_row['uo_financiadora']}, Ação = {data_row['acao']}, Natureza de Despesa = {data_row['natureza_despesa']}{data_row['item']}, Valor: {data_row['valor']}")


    # -------------------- exemplo para orquestrar o fluxo --------------------
    # aqui você pode inspecionar o data_row e decidir se é anulação ou aprovação, global ou amarrado, e então chamar as funções correspondentes

    if data_row['orientacao'] == 'Anular':
        retorno = anulacao(em, data_row)
    elif data_row['orientacao'] == 'Aprovar':
        retorno = aprovacao(em, data_row)

em.terminate()
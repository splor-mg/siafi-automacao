import os
from dotenv import load_dotenv
from py3270 import Emulator
from datetime import datetime
import time

load_dotenv()
sistema = os.getenv('SISTEMA')
usuario = os.getenv('USUARIO')
senha = os.getenv('SENHA')
unidade_executora = os.getenv('UNIDADE_EXECUTORA')

month = datetime.today().strftime("%m")
em = Emulator(visible=True)
em.connect('bhmvsb.prodemge.gov.br')
em.wait_for_field()


# Preenche os dados de login
em.fill_field(19, 13, sistema, 7)
em.fill_field(20, 13, usuario, 7)
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

    except ValueError:
        # Tela SEM campo editável — é a tela de aviso, só dá Enter e segue
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

    except ValueError:
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
#Entrar em 03 - Movimentacao Orcamentaria
em.fill_field(21, 19, '03', 2)
em.send_enter()
em.wait_for_field()
#Entrar em 02 - Aprovacao de Cota Orcamentaria
em.fill_field(21, 19, '02', 2)
em.send_enter()
em.wait_for_field()


## Verifica se é anulação ou aprovação e preencche 03-1 para aprovação e 04-1 para anulação
if valor_anulacao != 0:
    ##Anulação de cota orçamentária
    em.fill_field(21, 19, '04', 2)
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
em.terminate()
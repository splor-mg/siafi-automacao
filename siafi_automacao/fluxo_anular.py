

def anular(em, data_row):
    ## Verifica se é anulação ou aprovação e preencche 03-1 para aprovação e 04-1 para anulação
    
    # Movimentação de tela
    em.fill_field(21, 19, '04', 2)
    em.fill_field(21, 41, '1', 1)
    em.send_enter()
    em.wait_for_field()
            
    ## Verifica se é global ou amarrado e preenche os campos correspondentes
    if data_row['tipo_global'] == 'x':
        #Aprovação/Anulação GLOBAL
        em.fill_field(8, 52, data_row['month'], 2) # mes
        em.fill_field(9, 52, 'x', 1) # global
        em.fill_field(11, 52, f'{data_row['fonte']}{data_row['procedencia']}', 3) # fonte e procendencia
        em.fill_field(12, 52, data_row['uo'], 4) # UO
        em.fill_field(13, 52, data_row['grupo'], 1) # grupo de despesa
        em.fill_field(13, 54, data_row['iag'], 1) # IAG
        em.send_enter()
        em.wait_for_field()
    elif data_row['tipo_amarrado'] != '0':
        #Aprovação/Anulação AMARRADO
        em.fill_field(8, 52, data_row['month'], 2) # mes
        em.fill_field(10, 52, 'x', 1) # amarrado
        em.fill_field(11, 52, f'{data_row['fonte']}{data_row['procedencia']}', 3) # fonte e procendencia
        em.fill_field(12, 52, data_row['uo'], 4) # UO
        em.fill_field(13, 52, data_row['grupo'], 1) # grupo de despesa
        em.fill_field(13, 54, data_row['iag'], 1) # IAG
        em.fill_field(16, 52, data_row['elemento'], 2) # elemento
        em.fill_field(16, 55, data_row['item'], 2) # item
        em.send_enter()
        em.wait_for_field()


    #digitar as informações das ações e valores...
    em.fill_field(16, 17, data_row['acao'], 4) # ação
    em.fill_field(16, 22, '0001', 4)
    em.fill_field(16, 36, str(data_row['valor']), 15) # valor
    em.send_enter()
    em.wait_for_field()

    em.fill_field(11, 11, 'Remanejamento realizado conforme solicitado', 60) # ação
    em.send_enter()
    em.wait_for_field()

    em.send_pf(5)  # envia F5
    em.wait_for_field()
    em.send_pf(5)  # envia F5
    em.wait_for_field()
    retorno = em.string_get(1, 1, 80).strip()
    print(f"SIAFI retornou: {retorno}")
    em.send_pf(3)  # envia F5
    em.wait_for_field()

        
        
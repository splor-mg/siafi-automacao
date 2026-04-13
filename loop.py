
# Baixar a planilha do sharepoint
# ler a planilha e pegar os dados (pandas)

from frictionless import Package
package = Package('datapcakge.json')
resource = package.get_resource('planilha')
data = resource.to_pandas()



tabela = [
    {'uo': '12345', 'fonte': '10', 'ipu': '0'},
    {'uo': '12345', 'fonte': '11', 'ipu': '0'},
]

contador = 0
for linha in tabela:
    print('-------------------')
    print(f'Linha {contador}')
    #print(linha['uo'])
    #print(linha['fonte'])
    #print(linha['ipu'])
    em.fill_field(19, 13, linha['uo'], 7)
    contador += 1

    print('-------------------')
import openpyxl
def consulta():
    planilha = openpyxl.load_workbook('Clientes 2.xlsx')
    planilha_clientes = planilha['Clientes']    
    allid = list()
    for linha in planilha_clientes.iter_rows(min_row= 11, min_col=3, max_col=3):    
        for cedula in linha:
            allid.append(cedula.value)
    allid.remove(None)
    cliente = int(input("Qual ID cliente: "))
    if cliente not in allid:
      print("Cliente nao encontrado")
      allid.clear()
      consulta()
    else:
        verDados(cliente, planilha_clientes)

def verDados(cliente, planilha_clientes):
    c = 11
    for linha in planilha_clientes.iter_rows(min_row= 11, min_col=3, max_col=3):    
        for cedula in linha:        
            if cedula.value == cliente: 
                print('=-' * 20)         
                print(f"ID: {cedula.value}")
                print(f"NOME: {planilha_clientes[f'D{c}'].value}")
                tel = planilha_clientes[f'E{c}'].value
                if len(tel) == 11:                    
                    print(f'TEL: ({tel[0]}{tel[1]}) {tel[2:6]}-{tel[6:]}')
                else:
                    print(f"TEL: {planilha_clientes[f'E{c}'].value}")
                print(f"END: {planilha_clientes[f'F{c}'].value}")
                print('=-' * 20) 
            else:
                c += 1

def pegarDados():
    while True:
        nome = str(input('NOME: '))
        if nome not in '':
            break
    while True:    
        tel = str(input('TEL: '))
        if tel not in '':
            break
    while True:    
        end = str(input('END: ')) 
        if end not in '':
            break
    cadastar(nome, tel, end)


def cadastar(nome, tel, end):
    planilha = openpyxl.load_workbook('Clientes 2.xlsx')
    planilha_clientes = planilha['Clientes']   
    allid = list()
    for linha in planilha_clientes.iter_rows(min_row= 11, min_col=3, max_col=3):    
        for cedula in linha:
            if cedula.value == None:
                pass
            else:
                allid.append(cedula.value)
    
    allid.sort()
    idusuario = allid[-1] +1
    linha_dado = len(allid) +11  
    print('=-' * 20)
    print(f'Adicionar id no contato = {idusuario}')
    print('=-' * 20)
    planilha_clientes[f'C{linha_dado}'] = idusuario
    planilha_clientes[f'D{linha_dado}'] = nome
    planilha_clientes[f'E{linha_dado}'] = tel
    planilha_clientes[f'F{linha_dado}'] = end
    planilha.save('Clientes 2.xlsx')


def menu():
    import os
    print('   CLIENTES HOMEBURGUER')
    print('=-' * 20)
    
    while True:
        escolha = str(input('[1] BUSCAR CLIENTE\n[2] CADASTRAR CLIENTE\n[3] LIMPAR TELA\n[4] SAIR\nESCOLHA: '))
        if escolha in '1234':
            break
    if escolha == '1':
        consulta()        
        menu()
    if escolha == '2':
        pegarDados()        
        menu()
    if escolha == '3':
        os.system('cls')
        menu()
    if escolha == '4':
        while True:
            decisão = str(input('Deseja sair? [S/N] ')).strip().upper()[0]
            if decisão in 'SN':
                break
        if decisão == 'S':  
            os.system('exit')
        elif decisão == 'N':
            print()
            menu()    

if __name__ == "__main__":
    print()
    menu()
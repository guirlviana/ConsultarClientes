import openpyxl
def consulta(idcliente=0):
    planilha = openpyxl.load_workbook('Clientes 2.xlsm')
    planilha_clientes = planilha['Clientes']
    cliente = idcliente
    allid = list()
    for linha in planilha_clientes.iter_rows(min_row= 11, min_col=3, max_col=3):    
        for cedula in linha:
            allid.append(cedula.value)
    allid.remove(None)
    if cliente not in allid:
      print("Cliente nao encontrado")
    else:
        verDados(cliente, planilha_clientes)

def verDados(cliente, planilha_clientes):
    c = 11
    for linha in planilha_clientes.iter_rows(min_row= 11, min_col=3, max_col=3):    
        for cedula in linha:        
            if cedula.value == cliente:            
                print(f"ID: {cedula.value}")
                print(f"NOME: {planilha_clientes[f'D{c}'].value}")
                print(f"TEL: {planilha_clientes[f'E{c}'].value}")
                print(f"ENDEREÇO: {planilha_clientes[f'F{c}'].value}")
            else:
                c += 1
                
if __name__ == "__main__":
    consulta(int(input("Qual ID cliente: ")))
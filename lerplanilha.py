def consulta(idcliente=0):
    import openpyxl

    planilha = openpyxl.load_workbook('Clientes 2.xlsm')
    planilha_clientes = planilha['Clientes']
    cliente = idcliente
    c = 11
    allid = list()
    for linha in planilha_clientes.iter_rows(min_row= 11, min_col=3, max_col=3):    
        for cedula in linha:
            allid.append(cedula.value)
    allid.remove(None)
    print(allid)
        # if cliente in allid:
        #     for cedula in linha:        
        #         if cedula.value == cliente:            
        #             print(f"ID: {cedula.value}")
        #             print(f"NOME: {planilha_clientes[f'D{c}'].value}")
        #             print(f"TEL: {planilha_clientes[f'E{c}'].value}")
        #             print(f"ENDEREÇO: {planilha_clientes[f'F{c}'].value}")
        #         else:
        #             c += 1
        # else:
        #     print("Cliente nao encontrado")
            
if __name__ == "__main__":
    consulta(103)    
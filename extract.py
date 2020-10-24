import csv
import xlrd

def buscaValores(variavel):
    valores = []
    valores.append([])
    valores.append([])
    valores.append("")
    workbook = xlrd.open_workbook('dicionario_PNS_microdados_2013.xls')
    worksheet = workbook.sheet_by_index(0)

    i = 0
    busca = ''
    while busca != variavel:
        i+=1
        busca = str(worksheet.cell(i,2)).split(":")[1].replace('\'','')

    valores[0].append(str(worksheet.cell(i,5)).split(":")[1].replace('\'',''))
    valores[1].append(str(worksheet.cell(i,6)).split(":")[1].replace('\'',''))
    valores[2] = str(worksheet.cell(i,4)).split(":")[1].replace('\'','')
    i+=1
    #print(str(worksheet.cell(i,2)).split(":")[0])
    while str(worksheet.cell(i,2)) == "empty:''":
        
        valores[0].append(str(worksheet.cell(i,5)).split(":")[1].replace('\'',''))
        valores[1].append(str(worksheet.cell(i,6)).split(":")[1].replace('\'',''))
        i+=1

    for v in range(len(valores[0])):
        
        if valores[0][v] != "":
            valores[0][v] = int(valores[0][v].replace(".0",""))
            valores[0][v] = str(valores[0][v]) + ".0"
        
    
    return valores



with open('depressao.csv') as csv_file:
    
    csv_reader = csv.reader(csv_file, delimiter=',')
    dados = []
    linha = []

    '''with open('resp.csv') as csv_file2:
    
        csv_reader2 = csv.reader(csv_file2, delimiter=',')


        linha = csv_reader2.__next__()
        print(linha)
    
    print(type(linha))'''

    #COLOCANDO OS DADOS EM UMA LISTA (DE LISTAS)
    for i in range(8471): 
        linha = csv_reader.__next__()
        dados.append(list(linha[0].split(";")))

    aux = dados[0][0]
    i = 0
    achou = False

    #PROCURANDO COLUNA Q092 (DIAGNOSTICO DE DEPRESSAO)
    while aux != "Q092":
        i+=1
        aux = dados[0][i]

    depressaoPos = i       

    continua = True
    coluna = ""

    #MONTANDO O ARQUIVO DE SAIDA
    with open('registros.csv', 'w', newline='', encoding="utf-8") as file:
    
        writer = csv.writer(file)

        #CABECALHO
        writer.writerow(["codigo","pergunta", "resp_numero","resposta", "registros","depressao","sem depressao"])

        #VARIAVEIS
        colunas = ["N010","Q132","R041","R042","P040","P039","P038","P034","P036","P045","P008",
                   "P010","P012","P014","P017","P019","P021","P024","P02601","Q002","Q030","Q05501",
                   "Q05502","Q05503","Q05504","Q05505","Q05506","Q05507","Q05508","Q05509","Q063",
                   "Q068","Q079","Q084","Q116","Q120","Q124","Q11001","Q11002","Q11003","Q11004",
                   "P027","P032","P033","VDD004","P050","P051","P052","P055","P05401","P05404",
                   "P05407","P058","P068","M010","N011","N012","N013","N014","N015","N016","N017",
                   "N018","M005","M007","M008","M01101","M01102","M01103","M01104","M01105","M01106",
                   "M01107","M01108","M016","M018","M019","O009","O014","O020","O021",
                   "O022","O023","O024","O025","O027","O028","O029","O030","O021","O032","O033",
                   "O036","O037","O038","O040","O041","O042","O043","O044","O048"]
        
        #PARA CADA VARIAVEL EM COLUNAS
        for coluna in colunas:
                   
            aux = dados[0][0]
            i = 0
            achou = False
            
            #PROCURANDO A COLUNA DA VARIAVEL CORRENTE
            while aux != coluna:
                i+=1
                aux = dados[0][i]
                if aux == coluna:
                    achou = True                    

            if not achou:
                print ("Variavel nao encontrada")
            else:
                pos = i

                #VALORES POSSIVEIS DE RESPOSTA
                # valores[0]: CONTEM AS POSSIVEIS RESPOSTAS PARA A VARIAVEL CORRENTE(EX.: 1.0, 2.0...)
                # valores[1]: CONTEM O SIGNIFICADO DE CADA RESPOSTA(EX.: SIM, NAO, NAO APLICAVEL)
                # valores[2]: CONTEM A PERGUNTA (EX.: FEZ ALGUM USO DE MEDICAMENTO...)
                valores = buscaValores(coluna)            
                #print (valores[0])
                
                conta = [] #NUMERO DE REGISTROS PARA CADA RESPOSTA
                contaDep = [] #NUMERO DE REGISTROS COM DEPRESSAO PARA CADA RESPOSTA  
                
                #PARA CADA RESPOSTA POSSIVEL PARA A VARIAVEL CORRENTE
                for valor in range(len(valores[0])):

                    aux = 0
                    conta.append(0)
                    contaDep.append(0)
                    
                    #PARA CADA REGISTRO
                    for celula in range(len(dados)):
                        
                        if dados[celula][pos] == valores[0][valor]:
                            conta[valor] = conta[valor] + 1

                            #SE TIVER DEPRESSAO
                            if dados[celula][depressaoPos] == "1.0":
                                contaDep[valor] = contaDep[valor] + 1

                    #ESCREVE NO ARQUIVO
                    writer.writerow([coluna,valores[2], valores[0][valor],valores[1][valor], conta[valor],contaDep[valor],conta[valor] - contaDep[valor]])

                #print(valores)
                #print(conta)
                #print(contaDep)


    
    


        

        
        
        

    

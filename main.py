import openpyxl
import os
import numpy

os.chdir('I:\\WeatherDataProcessing')

# 1 - Indique a planilha de dados
wb = openpyxl.load_workbook('dados_01.xlsx')
plan = wb['ES_A617_ALEGRE_fornecimento_1']


listaDados = []

conta = 11
celula = "082018"
linha = 2

# 2 - Altere o colIni para o valor da coluna correspondente
colIni = 74
colEnd = colIni + 24

# temperatura máxima - Refazer
# Temperartura mínima - Fazendo
# precicitação - Lembrar de alterar a planilha fonte de dados

while celula != None:
    conta = conta + 1
    celula = plan.cell(row=conta, column=1).value

print('O total de linhas é: ', conta)

for i in range(12, 409, 1):
    print('o valor de i é:', i)

    # 3 - Indique a planilha de dados
    wb = openpyxl.load_workbook('dados_01.xlsx')
    plan = wb['ES_A617_ALEGRE_fornecimento_1']

    if i == 12:
        for j in range(colIni, colEnd, 1):
            temperaturaDia = plan.cell(row=i, column=j).value
            if temperaturaDia != 'NULL':
                listaDados.append(temperaturaDia)

    elif i != 12:
        # Se i for diferente de 12 ele vai guardar duas datas
        # A data da linha de coleta de dados atual
        try:
            data_01 = plan.cell(row=i - 1, column=1).value
            data_01_n = str(data_01.strftime('%m/%Y'))
            print('Estamos na linha', i - 1)
            print('O valor de data_01_n é: ', data_01_n)

            data = plan.cell(row=i, column=1).value
            data_n = str(data.strftime('%m/%Y'))
            print('Estamos na linha', i)
            print('O valor de data_n é: ', data_n)

        except:
            mes = data_01_n

            # 4 - Altere o tipo de calculo que o numpy irá realizar
            media = numpy.amax(listaDados)
            pastaDados = openpyxl.load_workbook('danilo.xlsx')
            sheet = pastaDados['Sheet']

            # 5 - Altere a coluna e o nome que será dado a coluna
            sheet.cell(row=1, column=3).value = 'Temp. Máximo ºC'
            sheet.cell(row=linha, column=1).value = mes

            # 6 - Altere a coluna e o nome que será dado a coluna aqui também
            sheet.cell(row=linha, column=3).value = media

            # 7 - Copie tudo do 2 ao 4 e cole na ultima condição else

            linha = linha + 1
            pastaDados.save('danilo.xlsx')

            print('A temperatura Média é: ', numpy.mean(listaDados))
            print('A data da ultima linha de dados é: ', data_01.strftime('%d/%m/%Y'))
            listaDados = []
            print('Zerou a lista na linha igual a :', i)
            for j in range(colIni, colEnd, 1):
                temperaturaDia = plan.cell(row=i, column=j).value
                if temperaturaDia != 'NULL':
                    listaDados.append(temperaturaDia)

        else:
            if data_n == data_01_n:
                for j in range(colIni, colEnd, 1):
                    temperaturaDia = plan.cell(row=i, column=j).value
                    if temperaturaDia != 'NULL':
                        listaDados.append(temperaturaDia)

            else:
                mes = data_01_n
                ################ 8 Cole daqui

                # 4 - Altere o tipo de calculo que o numpy irá realizar
                media = numpy.amax(listaDados)
                pastaDados = openpyxl.load_workbook('danilo.xlsx')
                sheet = pastaDados['Sheet']

                # 5 - Altere a coluna e o nome que será dado a coluna
                sheet.cell(row=1, column=3).value = 'Temp. Máximo ºC'
                sheet.cell(row=linha, column=1).value = mes

                # 6 - Altere a coluna e o nome que será dado a coluna aqui também
                sheet.cell(row=linha, column=3).value = media

                ################ Até aqui

                linha = linha + 1

                pastaDados.save('danilo.xlsx')

                print('A temperatura Média é: ', numpy.mean(listaDados))
                print('A data da ultima linha de dados é: ', data_01.strftime('%d/%m/%Y'))
                listaDados = []
                print('Zerou a lista na linha igual a :', i)
                for j in range(colIni, colEnd, 1):
                    temperaturaDia = plan.cell(row=i, column=j).value
                    if temperaturaDia != 'NULL':
                        listaDados.append(temperaturaDia)
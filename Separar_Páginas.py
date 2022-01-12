from genericpath import exists
import os
import shutil
from openpyxl import Workbook, load_workbook
import time
import PyPDF2
import re

def main():

    #Coletar informações dos colaboradores na base de dados
    Planilha_base_de_dados = load_workbook("C:\\Users\\gabriel.fonseca\\OneDrive - Energisa\\Documentos\\Colaboradores\\Colaboradores.xlsx")
    Aba = Planilha_base_de_dados.active

    funcionários = []
    matriculas = []
    funções = []
    equipes= []

    for celula_nome in Aba['B']:  
        linha_nome = celula_nome.row
        nome = str(Aba["B{}".format(linha_nome)].value)
        if nome == "Funcionário":
            time.sleep(0.00001)
        else:
            funcionários.append(nome)

    for celula_matricula in Aba['A']:  
        linha_matricula = celula_matricula.row
        matricula = str(Aba["A{}".format(linha_matricula)].value)
        if matricula == "Matricula":
            time.sleep(0.00001)
        else:
            matriculas.append(matricula)

    for celula_função in Aba['C']:  
        linha_função= celula_função.row
        função = str(Aba["C{}".format(linha_função)].value)
        if função == "Função":
            time.sleep(0.00001)
        else:
            funções.append(função)

    for celula_equipe in Aba['D']:  
        linha_equipe = celula_equipe.row
        equipe = str(Aba["D{}".format(linha_equipe)].value)
        if equipe == "Equipe":
            time.sleep(0.00001)
        else:
            equipes.append(equipe)

    #Coletar Periodo do Espelho e FPM
    mes = str(input("Qual o mês?\n")).upper()
    ano = str(input("Qual o ano?\n")).upper()
    periodo = str(input("Qual o periodo?\n'01a15' ou '16a31'?\n")).lower()
    documento = str(input("Qual o tipo do documento a ser separado?\n")).lower()

    inicio = periodo.split('a')

    caminho = "C:\\Users\\gabriel.fonseca\\OneDrive - Energisa\\Documentos\\Espelho de Ponto\\{}\\{}{} - {}".format(ano,mes,ano,periodo)
    if not os.path.exists(caminho):
        os.makedirs(caminho) #Criar diretório

    paginas_imprimir = open("C:\\Users\\gabriel.fonseca\\OneDrive - Energisa\\Documentos\\Espelho de Ponto\\{}\\{}{} - {}\\Páginas_imprimir.txt".format(ano,mes,ano,periodo), "w")

    if documento == 'espelho':
        
        #Abrir Arquivo
        espelho = PyPDF2.PdfFileReader("C:\\Users\\gabriel.fonseca\\Downloads\\Espelhos Parciais.pdf")

        # Coletar número de páginas do arquivo
        NumPages = espelho.getNumPages()

        # Extrair texto do PDF
        for i in range(0, NumPages):
            PageObj = espelho.getPage(i)
            Text = PageObj.extractText() 

            # Buscar no arquivo por
            for colaborador in funcionários:
                buscarpor = funcionários[funcionários.index(colaborador)]
                
                if buscarpor in Text:
                    paginas_imprimir.write("{}, ".format(i))

    if documento == 'fpm':
        
        #Abrir Arquivo
        fpm = PyPDF2.PdfFileReader("C:\\Users\\gabriel.fonseca\\Downloads\\FOLHA MANUAL.pdf")

        # Coletar número de páginas do arquivo
        NumPages = fpm.getNumPages()

        # Extrair texto do PDF
        firstpage = 0
        lastpage = 0

        for i in range(0, NumPages):
            PageObj = fpm.getPage(i)
            Text = PageObj.extractText() 

            # Buscar no arquivo por
            for colaborador in funcionários:
                buscarpor = funcionários[funcionários.index(colaborador)]
                
                if buscarpor in Text:
                    paginas_imprimir.write("{}, ".format(i+1))

    paginas_imprimir.close()

main()
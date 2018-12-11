import re
import os
import datetime
import calendar
import PyPDF2

def PDFsplit(pdf):
    
    #abre aqruivo txt e set as variaveis usadas
    pdfObjt = open('BOL.txt', 'rb')
    empresas = []
    splits = []
    matriz = []
    x = 0
    y = 0

    #limpeza de dados e criação de matriz dos dados
    for line in pdfObjt:
        linha = line.decode('UTF-8')
        linha = linha.strip(' ')   
        linha = re.sub('\/', '', linha)
        linha = re.split('\s{4}|\n', linha)
        linha = list(filter(lambda x: not re.match(r'^\s*$', x), linha))
        if linha == [] : continue
        matriz.append(linha)

    #extração de informação das empresas e dos contratos #splits #empresa
    for i in range(0,len(matriz)):
        linha = matriz[i]
     
        for j in range(0,len(linha)):
            if linha[j] ==  'Pagador':
                empresa = matriz[i+1]
                empresa = empresa[0]
                empresas.append(empresa)
                y = y+1
                splits.append(y)             
          
          #  if re.search('MSX-', linha[j]):
           #     titulo = linha[j]
            #    print(empresa)
             #   print(titulo)
              #  titulo = int(titulo[-3:])
               # print(titulo)
                #if x == 0:
                 #   empresas.append(empresa)
                    
                #if x < titulo:
                 #   x = titulo
                                    
                #else:
                   # y = x + y   
                   # x = titulo
                   
                    
    
    print(len(splits), splits)
    # creating input pdf file object 
    pdfFileObj = open(pdf, 'rb') 
      
    # creating pdf reader object 
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
      
    # starting index of first slice 
    start = 0
      
    # starting index of last slice 
    end = splits[0] 
    #nome da pasta para add boletos
    now = datetime.datetime.now()
    mes = now.month
    ano = now.year
    firstday = '01'
    lastday = calendar.monthrange(ano, mes)[1] 
    data = (str(firstday) +' '+ '10' +' '+ '31' +' '+ '10' +' '+ '18') #bol dentro do dir com esse nome
    cwd = os.getcwd()
    print(len(empresas),empresas)
    for i in range(len(splits) 	): 
        # creating pdf writer object for (i+1)th split 
        pdfWriter = PyPDF2.PdfFileWriter() 
          
        # output pdf file name ----------------> renomear o arquivo de saida 
        outputpdf = pdf.split('.pdf')[0] + '.pdf' 
          
        # adding pages to pdf writer object 
        for page in range(start,end): 
            pdfWriter.addPage(pdfReader.getPage(page)) 
          
        #writing split pdf pages to pdf file to new directories based on empresas
        ncwd = (cwd+'\\'+empresas[i]+'\\'+data)
        print(ncwd)
        if not os.path.exists(ncwd):
            os.makedirs(ncwd)
        with open(os.path.join(ncwd, outputpdf), "wb") as f: 
                  pdfWriter.write(f)
            

  
        # interchanging page split start position for next split 
        start = end 
        try: 
            # setting split end positon for next split 
            end = splits[i+1] 
        except IndexError: 
            # setting split end position for last split 
            end = pdfReader.numPages 
          
    # closing the input pdf file object 
    pdfFileObj.close() 


def main(): 
    # pdf file to split 
    pdf = 'BOL.pdf'
     
    # calling PDFsplit function to split pdf 
    PDFsplit(pdf) 
  
if __name__ == "__main__": 
    # calling the main function 
    main() 

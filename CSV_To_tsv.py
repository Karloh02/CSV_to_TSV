from tkinter import filedialog
import customtkinter
import aspose.cells
import os
from tkinter import *


app = customtkinter.CTk()
customtkinter.set_default_color_theme("green")
customtkinter.set_appearance_mode("dark")
app.title("CSV para TSV")
app.geometry("200x240")

diretorio = ["CSV", "LOCAL"]

#pega o lugar do arquivo 
def diretorioCSV():
    diretorio[0] = filedialog.askopenfilename(title = 'selecione o arquivo CSV', filetype = (("*.csv","*.CSV"),("all files", "*.xlsm")))
   
    return(diretorio)

#Faz a leitura do arquivo CSV como string e cria listas a partir desses strings separando os dados.
def criaLista_Apartir_ListaSTR(lista):
    lista_certa = []

    var_str = ""
    
    for i in range(len(lista[0])):

        if lista[0][i] != ";":
           var_str += lista[0][i]
        
        else:
            lista_certa.append(var_str)
            var_str = ""

    return(lista_certa)

#deixa a primeira linha na formatação que o MDA lê
def junta_dados(lista1, lista2):

    listaConcat = []

    if len(lista1) == len(lista2):

        for i in range(len(lista1)):
            if i ==0:
                if lista1[0] == "TEMPS" and str(lista2[0]) != "s": 
                    FatorMultiplicacao = float(lista2[0][:-1])
                else:
                    FatorMultiplicacao = 1


            listaConcat.append(lista1[i] + "_" + lista2[i])

    else:
        return()

    return(listaConcat, FatorMultiplicacao)

#faz a substituição de ponto e virgula dependendo do idioma do usuário, se esta em INGLES mantém o ponto, se esta em portugues troca por virgula.
def substirui_ponto_virgula(lista):

    IdiomaUser = idioma.get()
    
    if IdiomaUser == "EN":

        for i in range(len(lista)):
            lista[i] = str(lista[i]).replace(".", ".")
    
    elif IdiomaUser == "PT":
                
        for i in range(len(lista)):
            lista[i] = str(lista[i]).replace(".", ",")

        
    return(lista)

#função que fará a leitura do arquivo CSV, colocando o fator de multiplicação que pode ser 1 ou diferente de 1, dependendo do timestep selecionado nas aquisições.
def readCSV():
    try:
        #var é o direotrio do arquivo CSV
        var = diretorio[0]

        with open(var) as file:
            content = file.readlines()
        
        #Pega o fator de multiplicação para os casos em que há a necessidade de se corrigir os valores
        FatorMultiplicação = junta_dados(criaLista_Apartir_ListaSTR(content[4:5]), criaLista_Apartir_ListaSTR(content[5:6]))[1]  
        
        #data são todos os dados necessários para serem upados dentro do arquivo final
        data = []

        #pega as colunas de unidade e nome, para concatenar as duas e criar uma lista nova
        data.append(junta_dados(criaLista_Apartir_ListaSTR(content[4:5]), criaLista_Apartir_ListaSTR(content[5:6]))[0])
        
        #pega o número de linhas que existem dentro do arquivo CSV
        var_num = int(criaLista_Apartir_ListaSTR(content[6:7])[0])

        #le todas as linhas importantes, troca ponto por virgula.
        i = 8
        while i <= var_num + 7:
            data.append(substirui_ponto_virgula(criaLista_Apartir_ListaSTR((content[i - 1:i]))))
            i += 1

        #parte que fará a escrita do arquivo de texto. 
        file_name = (diretorio[0])[:-4]
        i = 1
        new_name = ""
        while file_name[-i] != "/":
            new_name += file_name[-i]
            i += 1
        file_name = new_name[::-1] 
        text_file = open(file_name + ".txt", "w")

        for k in range(len(data)):
            for p in range(len(data[k])):
                
                if p == 0 and k > 0:
                    text_file.write(str(float(data[k][p])*FatorMultiplicação) + "\t")

                else:
                    text_file.write(data[k][p] + "\t")
            text_file.write("\n")

        text_file.close()    

        from aspose.cells import Workbook 
        workbook = Workbook(file_name + ".txt")

        new_name = diretorio[1] + "/" + file_name + ".tsv"
        workbook.save(new_name)

        #Essa parte pode ser melhorada, e não sei como melhorar, crio um arquivo txt manipulo com ele e depois removo o aplicativo, possibilidade de melhoria
        os.remove(file_name + ".txt")

        popup = Toplevel(app)
        popup.title("Arquivo salvo")
        popup.geometry("200x40")
        popup.config(bg = "black")
        customtkinter.CTkLabel(popup, text= "Arquivo salvo com sucesso!", text_color = "light green").pack()

    except:
        popup = Toplevel(app)
        popup.title("Erro ao salvar")
        popup.geometry("225x40")
        popup.config(bg = "black")
        customtkinter.CTkLabel(popup, text= "Ocorreu um erro ao salvar o aqruivo!", text_color = "red").pack()


    return()

#função que seleciona o local onde o arquivo será salvado
def local():

    diretorio[1] = filedialog.askdirectory(title = "Onde salvar")
    return(diretorio)

#Organização do aplicativo
#Criação do frame que irá receber todas as informações do aplicativo
app.frame = customtkinter.CTkFrame(app, width=200, corner_radius=0)
app.frame.grid(row = 0, column = 0, rowspan = 3, sticky = "nsew")
app.frame.grid_rowconfigure(4, weight = 1)

#label para colocar o nome
app.nome = customtkinter.CTkLabel(app.frame, text = "Desenvolvido por Gabriel Karloh", text_color="white", font = ("Arial", 10))
app.nome.grid(row = 0, column = 0, pady = 3, padx = 10)

#Selecionar o idioma do PC
idioma = customtkinter.CTkOptionMenu(app.frame, values = ["EN", "PT"], width = 180)
idioma.grid(row = 1, column = 0, pady = 12, padx = 10)

#Botão para selecionar o local do arquivo CSV
app.buttonCSV = customtkinter.CTkButton(app.frame, text = "Escolha o arquivo CSV", command = diretorioCSV, width= 180)
app.buttonCSV.grid(row = 2, column = 0, pady = 12, padx = 10)

#Botão para selecionar o diretório onde se quer salvar o arquivo TSV
app.local_salva = customtkinter.CTkButton(app.frame, text = "Local onde salvar", command = local, width = 180)
app.local_salva.grid(row = 3, column = 0, pady = 12, padx = 10)

#Roda a aplicação e salva o arquivo final.
app.rodar = customtkinter.CTkButton(app.frame, text = "Rodar", command = readCSV, width = 75)
app.rodar.grid(row = 4, column = 0, pady = 12, padx = 10)

app.mainloop()

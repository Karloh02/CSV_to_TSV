from tkinter import filedialog
import customtkinter
import aspose.cells
import os
from tkinter import *


app = customtkinter.CTk()
customtkinter.set_default_color_theme("green")
customtkinter.set_appearance_mode("dark")
app.title("CSV para TSV")
app.geometry("200x205")

diretorio = ["CSV", "LOCAL"]

#pega o lugar do arquivo 
def diretorioCSV():
    diretorio[0] = filedialog.askopenfilename(title = 'selecione o arquivo CSV', filetype = (("*.csv","*.CSV"),("all files", "*.xlsm")))
   
    return(diretorio)

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

def junta_dados(lista1, lista2):

    listaConcat = []

    if len(lista1) == len(lista2):

        for i in range(len(lista1)):

            listaConcat.append(lista1[i] + "_" + lista2[i])

    else:
        return()

    return(listaConcat)

def substirui_ponto_virgula(lista):

    for i in range(len(lista)):
        lista[i] = str(lista[i]).replace(".", ",")
        
    return(lista)

def readCSV():

    #var é o direotrio do arquivo CSV
    var = diretorio[0]

    with open(var) as file:
        content = file.readlines()
    
    #data são todos os dados necessários para serem upados dentro do arquivo final
    data = []

    #pega as colunas de unidade e nome, para concatenar as duas e criar uma lista nova
    data.append(junta_dados(criaLista_Apartir_ListaSTR(content[4:5]), criaLista_Apartir_ListaSTR(content[5:6])))
    
    #pega o número de linhas que existem dentro do arquivo CSV
    var_num = int(criaLista_Apartir_ListaSTR(content[6:7])[0])

    #le todas as linhas importantes, troca ponto por virgula.
    i = 8
    while i <= var_num + 7:
        data.append(substirui_ponto_virgula(criaLista_Apartir_ListaSTR((content[i - 1:i]))))
        i += 1


    #parte que fará a escrita do arquivo de texto. 

    file_name = app.nome_arquivo.get()
    text_file = open(file_name + ".txt", "w")

    for k in range(len(data)):
        for p in range(len(data[k])):
            text_file.write(data[k][p] + "\t")
        text_file.write("\n")

    text_file.close()    

    from aspose.cells import Workbook 
    workbook = Workbook(file_name + ".txt")

    new_name = diretorio[1] + "/" + file_name + ".tsv"
    workbook.save(new_name)

    os.remove(file_name + ".txt")

    popup = Toplevel(app)
    popup.title("Arquivo salvo")
    popup.geometry("200x40")
    popup.config(bg = "black")
    customtkinter.CTkLabel(popup, text= "Arquivo salvo com sucesso!", text_color = "light green").pack()

    return()

#função que seleciona o local onde o arquivo será salvado
def local():

    diretorio[1] = filedialog.askdirectory(title = "Onde salvar")
    return(diretorio)


app.frame = customtkinter.CTkFrame(app, width=200, corner_radius=0)
app.frame.grid(row = 0, column = 0, rowspan = 3, sticky = "nsew")
app.frame.grid_rowconfigure(4, weight = 1)

app.buttonCSV = customtkinter.CTkButton(app.frame, text = "Escolha o arquivo CSV", command = diretorioCSV, width= 180)
app.buttonCSV.grid(row = 0, column = 0, pady = 12, padx = 10)

app.nome_arquivo = customtkinter.CTkEntry(app.frame, placeholder_text = "Nome do arquivo")
app.nome_arquivo.grid(row = 1, column = 0, pady = 12, padx = 10)

app.local_salva = customtkinter.CTkButton(app.frame, text = "Local onde salvar", command = local)
app.local_salva.grid(row = 2, column = 0, pady = 12, padx = 10)

app.rodar = customtkinter.CTkButton(app.frame, text = "Rodar", command = readCSV)
app.rodar.grid(row = 3, column = 0, pady = 12, padx = 10)

app.mainloop()
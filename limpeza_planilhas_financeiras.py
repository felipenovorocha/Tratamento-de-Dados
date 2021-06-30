#%%
from ntpath import join
import openpyxl
import pandas as pd
import glob as gb
import time
from pandas.io import excel
import os

class Excel_file_cleaner ():


    def __init__(self, excel_files = [""], file_name=""):

        # self.excel_files = []
        self.excel_files = []
        self.file_name = ""
        excel_files = self.excel_files

    def iniciar (self):
        excel_cleaner.generate_list()
        excel_cleaner.save_files()
        excel_cleaner.manual_data_filtering()
        excel_cleaner.auto_data_filtering()

    def generate_list(self):

    
        run = True
        #inserir manualmente o nome do arquivo
        while run:
            try:
                
                # anexar path ao nome do arquivo
                file_name = input("")
                if file_name != "done":
                    home = os.path.expanduser("~")
                    download_folder_path = os.path.join(home,"Downloads")
                    file = join(f"{download_folder_path}\{file_name}")
                    self.excel_files.append(file)
                        
            #quebrar loop
                else:
                    print(self.excel_files)
                    run = False
            except Exception as e:
                
                print(f"Excessão {e}: O nome do arquivo digitado foi {file}, esse nome está correto?")

        # salvar cada arquivo em formato xlsx na pasta do projeto
    
    def save_files (self):
        

        #para cada arquivo especificado na lista definida anteriormente
        for file in self.excel_files:
            print (file) #a partir da linha de código abaixo, o "path" torna-se um arquivo
            file = pd.read_excel(file) #ler o arquivo, passando o "path encontrado anteriormente"
            pd.set_option("display.max_rows",None)
            pd.set_option("display.max_columns",None)
            display (file)
            self.new_file_name = input ("Qual será o novo nome para esse arquvio?") #renomear arquivo
            new_file_name = self.new_file_name
            print(new_file_name)
            file = file.to_excel(new_file_name)
            #salvar o arquivo com um nome novo
    
    def manual_data_filtering(self):
        
        for file in self.excel_files:
                    
        #para que esse passo seja realizado, ainda é necessária a pré formatação antes de iniciar o programa
        #para o próximo pacote de atualizações, o objetivo será automatizar esse passo
               
            #filtrar dados indesejáveis
    
            #transformar em número
            
            self.filtered_files =[]
            file = pd.read_excel(file)
            display(file)

            filtering_status = "n"
            
            while filtering_status != "S" or filtering_status != "s":
                
                unwanted_rows = input("Qual (is) linhas quer deletar agora?")
                if unwanted_rows != "done":
                    file = file.drop(labels = range(0, 23) , axis=0) #filtrar linhas manualmente
                    display(file)
                    
                
                unwanted_column = input("Qual (is) coluna (s) quer deletar agora?")
                if unwanted_column != "done":
                    file = file.drop([unwanted_column],axis=1) #filtrar colunas manualmente
                    display(file)
                

                filtered_file = file.to_excel(self.new_file_name) #salvar alterações
                self.filtered_files.append

                
                filtering_status = input("Terminou? S/N") #definir se a filtragem desse arquivo foi concluido

                if filtering_status == "S" or filtering_status == "s":
                    break
        
    def auto_data_filtering (self):
        
        for filtered_file in self.filtered_files:
        
            filtered_file = pd.read_excel(filtered_file) #filtrar automaticamente linhas contendo erros 

            display(filtered_file) #mostrar arquivo antes da filtragem automática
            pd.set_option("display.max_rows", None)
            filtered_file = filtered_file.dropna(how = "all", axis=0) #filtrar automaticamente linhas contendo erros 

            display(filtered_file) #mostrar arquivo depois da  antes da corr

            numeric_columns = input("Qual(is) é (são) a (s) coluna (s) que contém valores numéricos?")
            filtered_file[numeric_columns] = pd.to_numeric(filtered_file[numeric_columns],errors="coerce") 
            display(filtered_file)
        
    def combine_sheets (self):
        # df = pd.read_excel(file)
            # merge = merge.append(df,ignore_index=False)
            # merged_file = merge.to_excel("arquivos_combinados.xlsx")
            # print(merged_file)
            pass  


        

excel_cleaner = Excel_file_cleaner()

excel_cleaner.iniciar()


# %%

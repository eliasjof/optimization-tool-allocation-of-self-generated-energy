import tkinter as tk
# para instalar: conda install -c anaconda tk
from tkinter.filedialog import askopenfilename
import pandas as pd
from pandas import ExcelWriter
import os
from xlrd import open_workbook
from lpsolve55 import *
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class myGUI(tk.Frame):
    
    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.root = parent
        self.build_gui()
        self.createFigure()
        ### RODANDO
        self.root.mainloop()
        self.name_file = ""
    
    def build_gui(self):                    
        # Build GUI
        ####################### MENU #######################################
        
        self.root.title("Ferramenta de Alocação Ótima de Geração Própria")
        windowWidth = self.root.winfo_screenwidth()
        windowHeight = self.root.winfo_screenheight() -150 #barra do windows
        size = str(windowWidth)+"x"+str(windowHeight) 
        self.root.geometry(size) #Tamanho da janela igual a da tela
        # self.root.geometry("500x500") #Tamanho da janela Width x Height
        self.root.state("zoomed") # maximizado

        ####################################################################

        ####################### MENU #######################################
        menu = tk.Menu(self.root)
        self.root.config(menu=menu)
        filemenu = tk.Menu(menu)
        menu.add_cascade(label="Arquivo", menu=filemenu)
        # filemenu.add_command(label="New", command=NewFile)
        filemenu.add_command(label="Abrir...", command=self.OpenFile)
        filemenu.add_separator()
        filemenu.add_command(label="Sair", command=self.root.quit)

        helpmenu = tk.Menu(menu)
        menu.add_cascade(label="Ajuda", menu=helpmenu)
        helpmenu.add_command(label="Sobre...", command=self.About)
        ##############################################################

        ######################### BOTÕES #####################################
        button_exec = tk.Button(self.root, text='Executar', command=self.Run, width=20, height=3, bg='#0052cc', fg='#ffffff', activebackground='#0052cc', activeforeground='#aaffaa')
        button_exec.pack()
        button_exec.place(x=10, y=windowHeight-3) # localização do botão


        ######################## CAIXA DE TEXTO INFORMATIVO #################
        w = tk.Label(self.root, text="Info:")
        w.pack()
        w.place(x=10, y=20) # localização do label
        S = tk.Scrollbar(self.root)
        self.info = tk.Text(self.root, height=30, width=50)
        S.pack(side=tk.RIGHT, fill=tk.Y)
        self.info.pack(side=tk.LEFT, fill=tk.Y)
        S.config(command=self.info.yview)        
        self.info.config(yscrollcommand=S.set)
        self.info.place(x=10, y=40) # localização do botão
        text = "Selecione o arquivo com os dados...\n"
        self.info.insert(tk.END, text)

    def createFigure(self):
        figure1 = plt.Figure(figsize=(9,5), dpi=100)
        self.axfig = figure1.add_subplot(111)
        self.fig = FigureCanvasTkAgg(figure1, root)
        self.fig.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH)
        self.fig.get_tk_widget().place(x=450, y = 50)    


    def Run(self):
        self.info.insert(tk.END, "Executando...\n")
        if(self.name_file !=""):
            self.myexternalfunction()
        else:
            self.info.insert(tk.END, "Selecione primeiro um arquivo...\n")

    def NewFile(self):
        print("New File!")

    # Abre arquivo: salva o nome do arquivo a ser trabalhado
    def OpenFile(self):
        self.name_file = askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        self.info.insert(tk.END, "Arquivo selecionado: \n")
        self.info.insert(tk.END,self.name_file)        
        self.info.insert(tk.END, "\n")

    def About(self):
        window = tk.Tk()
        window.title("Sobre")
        string = "Este trabalho foi desenvolvido no IFMG - Campus avançado Itabirito."    
        msg = tk.Message(window, text = string)
        msg.config(bg='lightgreen', font=('times', 24, 'normal'))
        msg.pack()

    def myexternalfunction(self):
        allocation(self.name_file, self.info, self.root)

#######################################################
# Função para alocação otimizada de energia
########################################################
def allocation(name_file, info, root):
    # Lendo o arquivo de entrada (Excel), abas: consumo_encargos e geração ###############
    
    # Agrupando o arquivo do Excel usando ExcelFile
    dados = pd.ExcelFile(name_file)

    # Armazenando cada aba do excel em uma variável
    d = dados.parse("consumo_encargos")
    p = dados.parse("geracao")

    ############################################# Separando os dados por Mês para apresentar o resultado mensal ########################

    mes = sorted(list(set(d['Mês']))) # criando a lista dos meses apresentados nos dados 

    for mesj in range(0, len(mes)): # laço para possibilitar que o algoritmo de otimização forneça o resultado para cada mês
        c = d.loc[d['Mês'] == mes[mesj]] # separando os dados consumo_encargos do mês j em uma variável
        g = p.loc[p['Mês'] == mes[mesj]] # separando os dados geracao do mês j em uma variável
        m = c.shape[0] # verificando a quantidade de linhas do planilha consumo_encargos
        # print ("a quantidade de linhas é: ", m)

        # subtraindo o encargo do encargo ape na coluna criada Tarifa (obs: se ONS: coluna "encargos" =  CDE + proinfa e coluna "encargos ape" = 0)
        # c['DescontoEncargos'] = c['Encargo'] - c['Encargo APE']
        # print(c)

        ############################################### Aba Geração ############################################################################

        ####definição da geração passível de alocação
        # g['GeracaoAlocacao'] = ((g['GarantiaFisicaComperdas'] + g['GSFouSecundaria']) * (g['PGDA'])) - (g['Contrato'])
        # print(g)
        gdata = list(((g['GarantiaFisicaComperdas'] + g['GSFouSecundaria']) * (g['PGDA'])) - (g['Contrato']))
        g.insert(0,'GeracaoAlocacao',gdata, True)

        #### soma do total da geração passível de alocação
        soma_geracao = g['GeracaoAlocacao'].sum(axis=0)
        # print("O valor total de geração passível de alocação é = ", soma_geracao, type(soma_geracao))

        ####criando coluna do montante de Alocação
        # c['Valor Alocado'] = 0
        # c['Desconto'] = 0
        # print(c)

        ###########################3###### Início do Algoritmo Simplex utilizando o LpSolve #######################################################

        #### SPE - Aplicação da regra: "usinas participantes de SPE podem alocar a geração apenas para cargas com demanda >= 3 MW"

        # restrição: soma de todos elementos <= geração passível de alocação
        c2 = c.copy()
        c2['Consumo'].loc[c['Demanda'] < 3] = 0
        g2 = g.loc[g['SPE'] == 'sim'] # separando os dados de geração referentes a usinas SPE
        soma_geracao_spe = g2['GeracaoAlocacao'].sum(axis=0) # encontrando a geração total passível de alocação de usinas SPE
        valores_restricoes2 = list(c2['Consumo']) # criando uma lista com os valores de consumo de unidades cuja demanda é >= 3MW
        restricao_soma2 = np.ones((1, m), dtype=float)[0] # criando matriz de 1 (uns) de mesma dimensão para possibilitar a aplicação da restrição
        lp1=lpsolve('make_lp',0, m) # iniciando lpsolve
        lpsolve('add_constraint',lp1,restricao_soma2, LE, float(soma_geracao_spe)) # restrição: soma de todos elementos <= geração passível de alocação

        # restrição: o valor alocado deve ser <= consumo da unidade
        for i in range(m): 
            restricaoi2 = np.zeros((1,m))[0]
            restricaoi2[i] = 1 # a matriz inicialmente de zeros, receberá o valor de 1 no elemento iterado pelo laço para que a restriçaõ seja aplicada
            lpsolve('add_constraint',lp1,restricaoi2, LE,valores_restricoes2[i]) # restrição: o valor alocado deve ser <= consumo da unidade

        custo2 = list(c2['Encargo'] - c2['Encargo APE'])
        c.insert(0,'DescontoEncargos',custo2, True)
        
        # executando o algoritmo simplex para encontrar o resultado otimizado
        lpsolve('set_obj_fn',lp1,custo2) # função obj: custo refere-se ao desconto total que é o peso de cada alocação
        lpsolve('set_maxim',lp1) # maximização: aplicando a maxim. da função objetivo
        result=np.round(lpsolve('solve',lp1), 2) # comando para resulver o algoritimo simplex e armazenando em uma variável
        objspe=np.round(lpsolve('get_objective', lp1), 2) # comando para obter o valor da maximizado da função objetivo e reservando em uma variável
        xspe=np.array(lpsolve('get_variables', lp1)[0]) # comando para obter os valores de cada alocação e reservando em uma variável

        #### Alocação restante gerado por usinas que não são SPE
        custo = list(c['Encargo'] - c['Encargo APE'])
        c.insert(0,'DescontoEncargos',custo, True)
        valores_restricoes = list(c['Consumo'])
        restricao_soma = np.ones((1, m), dtype=float)[0]
        lp2=lpsolve('make_lp',0, m)
        lpsolve('add_constraint',lp2,restricao_soma, LE, float(soma_geracao-soma_geracao_spe)) # restrição soma de todos elementos <= geração p. aloc.

        for i in range(m):
            restricaoi = np.zeros((1,m))[0]
            restricaoi[i] = 1
            lpsolve('add_constraint',lp2,restricaoi, LE,(valores_restricoes[i] - xspe[i]))
            
        # executando o algoritmo simplex para encontrar o resultado otimizado
        lpsolve('set_obj_fn',lp2,custo) # função obj: custo refere-se ao desconto total que é o peso de cada alocação
        lpsolve('set_maxim',lp2) # maximização: aplicando a maxim. da função objetivo
        result=np.round(lpsolve('solve',lp2), 2) # comando para resulver o algoritimo simplex e armazenando em uma variável
        objgeral=np.round(lpsolve('get_objective', lp2), 2) # comando para obter o valor da maximizado da função objetivo e reservando em uma variável
        xgeral=np.array(lpsolve('get_variables', lp2)[0]) # comando para obter os valores de cada alocação e reservando em uma variável

        ##Alocação proporcional    ///
        xspeprop = np.zeros((1,m))[0]
        xgeralprop = np.zeros((1,m))[0]
        xtotalprop = np.zeros((1,m))[0]
        for t in range(m):
            xspeprop[t] = (valores_restricoes2[t]/np.round(sum(valores_restricoes2), 2))*np.round(soma_geracao_spe, 2)
            if xspeprop[t] > valores_restricoes2[t]:
                xspeprop[t] = valores_restricoes2[t]
            xgeralprop[t] = (valores_restricoes[t]/np.round(sum(valores_restricoes), 2))*np.round((soma_geracao-soma_geracao_spe), 2)
            if xgeralprop[t] > valores_restricoes[t]:
                xgeralprop[t] = valores_restricoes[t]
            xtotalprop[t] = xspeprop[t] + xgeralprop[t]
            if xtotalprop[t] > valores_restricoes[t]:
                xgeralprop[t] = valores_restricoes[t]

        #### resultados
        obj=np.round(objgeral, 2) +  np.round(objspe, 2)
        x = np.round(xgeral, 2) +  np.round(xspe, 2)
        # ///
        soma_geracao_geral = soma_geracao - soma_geracao_spe
        pspe = np.array(xspe/soma_geracao_spe)
        pgeral = np.array(xgeral/(soma_geracao_geral))
        pspeprop = np.array(xspe/soma_geracao_spe)
        pgeralprop = np.array(xgeral/(soma_geracao_geral))
        desconto_proporcional = sum(xtotalprop*custo)

        # print(obj) # imprimindo o valor máximo da função obj
        # print("solução = ", np.round(x, 2)) # imprimindo os valores de cada alocação
        # print("percentual = ", np.round(custo, 2)) # imprimindo os % de cada alocação
        # print("valores máximos de alocação = ", np.round(valores_restricoes, 2)) # imprimindo os valores limites de alocação de cada unidade
        # print("ALOCAÇÃO PARA O MÊS {0:.10}".format(str([mes[j]]))) ## ARRUMAR A VISUALIZAÇÃO
        # for ii in range(0,len(x)):
        #     print("Unidade consumidora {}: alocar {:.2f} (p = {:.2f})".format(ii+1, x[ii], custo[ii]))
        
        #\\\\
        
        
        info.insert(tk.END, "\nALOCAÇÃO PARA O MÊS {0:.10}\n".format(str(mes[mesj]))) ## ARRUMAR A VISUALIZAÇÃO
        info.insert(tk.END, "\nDesconto alocação otimizada: "+ "R$ " + str(np.round(obj,2)) + "\n") # imprimindo o valor máximo da função obj
        info.insert(tk.END, "Desconto alocação não otimizada: "+ "R$ " + str(np.round(desconto_proporcional,2)) +"\n")
        info.insert(tk.END, "Ganho alocação otimizada: "+ "R$ " + str(np.round(obj-desconto_proporcional)) + ", " + str(np.round(((obj-desconto_proporcional)/desconto_proporcional)*100, 2)) + "%"+"\n")
        info.insert(tk.END, "conferência da % total alocada spe: "+ str(np.round(sum(pspe)))+"\n")
        info.insert(tk.END, "conferência da % total alocada geral: "+ str(np.round(sum(pgeral)))+"\n\n")
        xx = []
        yy = []
        zz = []
        for ii in range(0,len(x)):
            info.insert(tk.END,"Unidade consumidora {}: \nalocar spe {:.2f} MWh (% = {:.2f})\nalocar geral {:.2f} MWh (% = {:.2f})\n\n".format(ii+1, xspe[ii], pspe[ii], xgeral[ii], pgeral[ii]))
        
        for kk in range(0,len(mes)):
            xx.append(str(kk+1))
            yy.append(obj)
            
        zz.append(desconto_proporcional)        

        figure1 = plt.Figure(figsize=(9,5), dpi=100)
        axfig = figure1.add_subplot(111)
        fig = FigureCanvasTkAgg(figure1, root)
        fig.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH)
        fig.get_tk_widget().place(x=450, y = 40)
        
        # axfig.title('Alocação Otimizada')
        axfig.bar(xx,list(c['Consumo']),color = 'b')
        axfig.bar(xx,yy,color = 'b')
        axfig.bar(xx,zz,color = 'r')
        # axfig.bar(xx,zz,color = 'r',bottom=list(c['Consumo']))
        axfig.legend(labels=['Desconto otimizado', 'Desconto não otimizado'])
        # axfig.grid(c='k')
        
        print(yy)
        print(zz)

        #\\\\
        # info.insert(tk.END, "Desconto alocação proporcional: "+ str(np.round(desconto_proporcional,2)) +"\n") # imprimindo o valor máximo da função obj
        
        # info.insert(tk.END, "ALOCAÇÃO PARA O MÊS {0:.10}\n".format(str(mes[mesj]))) ## ARRUMAR A VISUALIZAÇÃO
        # xx2 = []
        # yy2 = []
        # for ii in range(0,len(xtotalprop)):
        #     # info.insert(tk.END,"Unidade consumidora {}: alocar spe {:.2f} (p spe = {:.2f}), alocar geral {:.2f} (p geral = {:.2f})\n".format(ii+1, xspe[ii], pspe[ii], xgeral[ii], pgeral[ii]))
        #     xx2.append(str(ii+1))
        #     yy2.append(xtotalprop[ii])

        # figure2 = plt.Figure(figsize=(9,5), dpi=100)
        # axfig2 = figure1.add_subplot(212)
        # fig2 = FigureCanvasTkAgg(figure1, root)
        # fig.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH)
        # fig.get_tk_widget().place(x=450, y = 40)
        
        # # axfig2.title(title = 'Alocação Não Otimizada')
        # axfig2.bar(xx2,list(c['Consumo']),color = 'b')
        # # axfig.bar(xx,yy,color = 'r',bottom=list(c['Consumo']))
        # axfig2.bar(xx2,yy2,color = 'r')
        # axfig2.legend(labels=['Consumo', 'Alocado'], title = 'Alocação Não Otimizada')
        # # axfig.grid(c='k')
        
        # print(yy2)
        # print(list(c['Consumo']))


if __name__ == "__main__":
    root = tk.Tk()
    myGUI(root)

    
    
    
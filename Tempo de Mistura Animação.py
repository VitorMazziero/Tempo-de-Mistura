#Bibliotecas:
import matplotlib;
import matplotlib.pyplot as plt
import matplotlib.animation as animation
from matplotlib.pyplot import plot, ion, show
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
import serial
import serial.tools.list_ports
import time 
from PIL import Image, ImageTk
from scipy.stats import linregress
from itertools import count 
from threading import Timer
import statistics
import pandas as pd
import xlsxwriter
import os
import datetime

#GIF do reator***
class ImageLabel(tk.Label):
    def load(self, im):
        if isinstance(im, str):
            im = Image.open(im)
        self.loc = 0
        self.frames = []

        try:
            for i in count(1):
                self.frames.append(ImageTk.PhotoImage(im.copy()))
                im.seek(i)
        except EOFError:
            pass

        try:
            self.delay = im.info['duration']
        except:
            self.delay = 100

        if len(self.frames) == 1:
            self.config(image=self.frames[0])
        else:
            self.next_frame()
    def unload(self):
        self.config(image=None)
        self.frames = None
    def next_frame(self):
        if self.frames:
            self.loc += 1
            self.loc %= len(self.frames)
            self.config(image=self.frames[self.loc])
            self.after(self.delay, self.next_frame)    
#Imagens
def Imagem1(): #Reator parado
    root.original1 = Image.open('reactor.gif') #Nome da imagem - Local da imagem > pasta do arquivo
    resized1 = root.original1.resize((151, 271),Image.ANTIALIAS) #Mudança de escala
    root.image1 = ImageTk.PhotoImage(resized1) 
    root.display1 = Label(root, image = root.image1)
    root.display1.grid(row=1, column=21, rowspan=10, columnspan=11) #Local na janela
    root.display1.configure(background='white')
def Imagem2(): #Reator agitado
    lbl = ImageLabel(root)
    lbl.grid(row=1, column=21, rowspan=10, columnspan=11)
    lbl.load('reactor.gif')
    lbl.configure(background='white') 
    
#Criação da janela
root = tk.Tk() #Nome da janela (root - main)

#root.geometry('1280x480')
root.title('Tempo de Mistura v.1.0') #Título do programa
root.configure(background='white') #plano de fundo da janela
root.grid_columnconfigure(5, minsize=100)

#Início da janela no centro da área de trabalho
w = 1250; h = 675 
ws = root.winfo_screenwidth() #Largura da janela
hs = root.winfo_screenheight() # Altura da tela
x = (ws/2) - (w/2); y = ((hs/2)-35) - (h/2)
root.geometry('%dx%d+%d+%d' % (w, h, x, y))

#Ícone do Programa
img = PhotoImage(file='Imagem3.png') 
root.tk.call('wm', 'iconphoto', root._w, img)

#Interface
spacer=Label(root); spacer.grid(row=10, column=1)
spacer=Label(root); spacer.grid(row=15, column=15)

#Imagem de fundo
root.original = Image.open('bg.png') #Nome da imagem - Local da imagem > pasta do arquivo
resized = root.original.resize((1250, 675),Image.ANTIALIAS) #Mudança de escala
root.image = ImageTk.PhotoImage(resized) 
root.display = Label(root, image = root.image)
root.display.grid(row=0, column=0, columnspan=40, rowspan=35) #Local na janela
root.display.configure(background='white')

#Gráfico 
fig=Figure(figsize=(10,4), dpi=100) #Tamanho e qualidade da figura
plot=fig.add_subplot(1,1,1) #Figura = Gráfico em subplot do Matlib
plotcanvas = FigureCanvasTkAgg(fig, root) #Gráfico
plotcanvas.get_tk_widget().grid(column=0, row=0, columnspan=21, rowspan=10)
plot.set_xlabel('Tempo (s)', color='k') #Legenda de eixo x
plot.set_ylabel('Temperatura (°C)', color='k') # Legenda de eixo y

#Legendas e entrada das variáveis
#Parâmetros iniciais
#Entrada e Legenda 1: Duração do pulso
text1=StringVar()#Nome da variável
e1=Entry(root,textvariable=text1, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)#Entrada do usuário
e1.grid(row=11, column=3)#Local da entrada na janela
e1.insert(END, '1.5')#Valor padrão da célula
l1=Label(root, text="Duração do pulso (s)", relief="flat", font= "Arial 9", width=20, anchor=E, justify=RIGHT, bg='white') #Legenda
l1.grid (row=11, column=2)#Local da legenda na janela
#Entrada e Legenda 2: Volume total (m3)
text2=StringVar()
e2=Entry(root,textvariable=text2, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e2.grid(row=12, column=3)
e2.insert(END, '0.03')
l2=Label (root, text="Volume total do reator (m³)", bd=5, relief="flat", font= "Arial 9", width=20, height=1, anchor=E, justify=RIGHT, bg='white')
l2.grid (row=12, column=2)
#Entrada e Legenda 3: Volume do pulso (m3)
text3=StringVar()
e3=Entry(root,textvariable=text3, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e3.grid(row=13, column=3)
e3.insert(END, '0.001')
l3=Label (root, text="Volume do pulso (m³)", relief="flat", font= "Arial 9",width=20, anchor=E, justify=RIGHT, bg='white')
l3.grid (row=13, column=2)
#Entrada e Legenda 4: Diferança de temperatura:
text4=StringVar()
e4=Entry(root,textvariable=text4, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e4.grid(row=14, column=3)
e4.insert(END, '20')
l4=Label (root, text="ΔT (°C)", bd=5, relief="flat", font= "Arial 9",width=20, height=1, anchor=E, justify=RIGHT, bg='white')
l4.grid (row=14, column=2)

#Flag Binário
#Entrada e Legenda 5: Erro em °C
text5=StringVar() 
e5=Entry(root,textvariable=text5, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2) 
e5.grid(row=16, column=3) 
e5.insert(END, '0.07') 
l5=Label (root, text="Erro (°C)", relief="flat", font= "Arial 9", width=20, anchor=E, justify=RIGHT, bg='white')
l5.grid (row=16, column=2) 
#Entrada e Legenda 6: i(t) mínimo:
text6=StringVar()
e6=Entry(root,textvariable=text6, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e6.grid(row=17, column=3)
e6.insert(END, '0.2')
l6=Label (root, text="i(t) mínimo", relief="flat", font= "Arial 9",width=20, anchor=E, justify=RIGHT, bg='white')
l6.grid (row=17, column=2)
#Entrada e Legenda 7: inclinação mínima:
text7=StringVar()
e7=Entry(root,textvariable=text7, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e7.grid(row=18, column=3)
e7.insert(END, '0.05'); e7.configure(state='disabled')
l7=Label (root, text="dT/dt mínimo", relief="flat", font= "Arial 9", width=20, anchor=E, justify=RIGHT, bg='white')
l7.grid (row=18, column=2); l7.config(fg="gray")

#Variáveis do teste
#Entrada e Legenda 8: Temperatura do reator
text8=StringVar()
e8=Entry(root,textvariable=text8, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e8.grid(row=13, column=6)
e8.configure(state='disabled')
l8=Label (root, text="T(°C) do reator", relief="flat", font= "Arial 9", width=20, height=1, anchor=E, justify=RIGHT, bg='white')
l8.grid (row=13, column=5); l8.config(fg="gray")
#Entrada e Legenda 9: Temperatura do traçador
text9=StringVar()
e9=Entry(root,textvariable=text9, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e9.grid(row=14, column=6)
e9.configure(state='disabled')
l9=Label (root, text="T(°C) do traçador", relief="flat", font= "Arial 9", width=20, height=2, anchor=E, justify=RIGHT, bg='white')
l9.grid (row=14, column=5); l9.config(fg="gray")
#Entrada e Legenda 10: Tempo de início do pulso
text10=StringVar()
e10=Entry(root,textvariable=text10, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e10.grid(row=11, column=6)
e10.configure(state='disabled')
l10=Label (root, text="Início do pulso (s)", relief="flat", font= "Arial 9", width=20, height=1, anchor=E, justify=RIGHT, bg='white')
l10.grid (row=11, column=5); l10.config(fg="gray")
#Entrada e Legenda 11: Tempo de fim do pulso
text11=StringVar()
e11=Entry(root,textvariable=text11, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e11.grid(row=12, column=6)
e11.configure(state='disabled')
l11=Label (root, text="Fim do pulso (s)", relief="flat", font= "Arial 9", width=20, height=2, anchor=E, justify=RIGHT, bg='white')
l11.grid (row=12, column=5); l11.config(fg="gray")
#Entrada e Legenda 12: Vazão do traçador
text12=StringVar()
e12=Entry(root,textvariable=text12, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e12.grid(row=15, column=6)
e12.configure(state='disabled')
l12=Label (root, text="Vazão do traçador (L/s)", relief="flat", font= "Arial 9", width=20, height=1, anchor=E, justify=RIGHT, bg='white')
l12.grid (row=15, column=5); l12.config(fg="gray")
#Entrada e Legenda 13: Temperatura de equilíbrio
text13=StringVar()
e13=Entry(root,textvariable=text13, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e13.grid(row=16, column=6)
e13.configure(state='disabled')
l13=Label (root, text="T(°C) de equilíbrio", relief="flat", font= "Arial 9", width=20, height=2, anchor=E, justify=RIGHT, bg='white')
l13.grid (row=16, column=5); l13.config(fg="gray")
#Entrada e Legenda 14 e 15: Densidades
text14=StringVar()
e14=Entry(root,textvariable=text14, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e14.grid(row=17, column=6)
e14.configure(state='disabled')
l14=Label (root, text="ρ água reator (kg/m³)", relief="flat", font= "Arial 9", width=20, height=1, anchor=E, justify=RIGHT, bg='white')
l14.grid (row=17, column=5); l14.config(fg="gray")
text15=StringVar()
e15=Entry(root,textvariable=text15, width=10, justify=RIGHT, font="bold 10", relief="groove", bd=2)
e15.grid(row=18, column=6)
e15.configure(state='disabled')
l15=Label (root, text="ρ água traçador (kg/m³)", relief="flat", font= "Arial 9", width=20, height=2, anchor=E, justify=RIGHT, bg='white')
l15.grid (row=18, column=5); l15.config(fg="gray")

#Desabilitar e habilitar entradas
def disabilitar1():
    if (v1.get()==0):
        e1.configure(state='normal'); l1.config(fg="black"); e2.configure(state='normal'); l2.config(fg="black"); 
        e3.configure(state='normal'); l3.config(fg="black"); e4.configure(state='normal'); l4.config(fg="black")
        e5.configure(state='normal'); l5.config(fg="black"); e6.configure(state='normal'); l6.config(fg="black")
        e7.configure(state='disabled'); l7.config(fg="gray")
    else:
        e1.configure(state='disabled'); l1.config(fg="gray"); e2.configure(state='disabled'); l2.config(fg="gray")
        e3.configure(state='disabled'); l3.config(fg="gray"); e4.configure(state='disabled'); l4.config(fg="gray")
        e5.configure(state='disabled'); l5.config(fg="gray"); e6.configure(state='disabled'); l6.config(fg="gray")
        e7.configure(state='normal'); l7.config(fg="black")
        
#VariáveisRadiobutton: seleçao de fluido e repetições
v1 = tk.IntVar(); v2 = tk.IntVar()
# Grupo 1 de radiobutton
s1r1=tk.Radiobutton(root,text="Óleo", variable=v1, value=1, command=disabilitar1) 
s1r2=tk.Radiobutton(root,text="Água", variable=v1, value=0, command=disabilitar1)
# Grupo 2 de radiobutton
s2r1=tk.Radiobutton(root,text="Teste único", variable=v2, value=0)
s2r2=tk.Radiobutton(root,text="Duplicata", variable=v2, value=1)
s2r3=tk.Radiobutton(root,text="Triplicata", variable=v2, value=2)
#Local dos radiobuttons na janela
s2r1.configure(background='white') 
s2r2.configure(background='white')
s2r3.configure(background='white')
s1r1.configure(background='white')
s1r2.configure(background='white')
s1r1.grid(column=11, row=11, rowspan=2, columnspan=3, sticky=W)
s1r2.grid(column=11, row=11, rowspan=2, columnspan=3, sticky=E)
s2r1.grid(column=12, row=14, sticky=S)
s2r2.grid(column=10, row=15, rowspan=2, columnspan=8, sticky=W)
s2r3.grid(column=9, row=15, rowspan=2, columnspan=6, sticky=E)

#Entrada 16: Caminho para salvar
text16=StringVar(); text16.set(os.getcwd()) 
e16=Entry(root,textvariable=text16, width=48, justify=LEFT, font="bold 9", relief="groove", bd=2)
e16.grid(row=15, column=15, rowspan=5, columnspan=25)
#Botão para seleção do caminho
def path():
    e16.delete(0, 'end') 
    root.directory = filedialog.askdirectory()
    text16.set(root.directory)
photo=PhotoImage(file="save.png")
btn1 = Button(root, command=path, fg='white', width=15, bd=2, image = photo, anchor=W)
btn1.grid(row=15, column=16, rowspan=5)

#Legenda 16 a 20: Conexão
l16=Label (root, text='Desconectado', relief="sunken", font= "Arial 9", width=15, height=1, justify=LEFT, bg='white')
l16.grid (row=18, column=11, columnspan=3, rowspan=10);  l16.config(fg="gray")
l17=Label (root, text='-', relief="flat", font= "Arial 8", width=6, height=1, justify=LEFT, bg='white', anchor=N)
l17.grid (row=18, column=14, columnspan=6, rowspan=15);
l18=Label (root, text='T1', relief="flat", font= "Arial 8", width=10, height=1, bg='white')
l18.grid (row=18, column=20, rowspan=15);
l19=Label (root, text='T2', relief="flat", font= "Arial 8", width=10, height=1, bg='white')
l19.grid (row=18, column=25, rowspan=15)
l20=Label (root, text='T3', relief="flat", font= "Arial 8", width=10, height=1, bg='white')
l20.grid (row=18, column=30, rowspan=15);

#ScrolledText: Resultados Tempo de Mistura
txt = ScrolledText(root,width=29,height=8)
txt.grid(row=8, column=20, rowspan=12, columnspan=16)

#Imagem do reator
Imagem1()

#Funções de bloqueio de fechamento da janela
def close_program():
    root.destroy()
def disable_event():
    pass   

#Encontrar arduino:
def encon_arduino(): #Função para deterninação da porta onde os dados serão coletados (para funcionar em diferentes PCs)
    for pinfo in serial.tools.list_ports.comports(): #procura na lista de informações das portas disponíveis
        if pinfo.description.startswith("Arduino"): 
            Arduino=(pinfo.device) #info.device=COM(n)
            l17.config(text=Arduino)
            return Arduino
        
#Coleta de dados do Arduino e Cálculo do Tempo de mistura instantaneo
def runaniA():
    #Reset dados para repetição
    e8.delete(0, 'end'); e9.delete(0, 'end') ; e10.delete(0, 'end') ;e11.delete(0, 'end') ; e12.delete(0, 'end')  ; e13.delete(0, 'end'); e14.delete(0, 'end'); e15.delete(0, 'end'); txt.delete(1.0,END)  
    l8.config(fg="gray");l9.config(fg="gray"); l10.config(fg="gray"); l11.config(fg="gray"); l12.config(fg="gray"); l13.config(fg="gray"); l14.config(fg="gray"); l15.config(fg="gray")
    k=IntVar(); k.set(0); w=IntVar(); w.set(0) #Contadores de execução única
    j=IntVar(); j.set(0) #Contador de execução final
    
    #Listas de valores para aquisição dos dados
    DadosArray=[] #Dados do sensor 
    T1=[]; T2=[]; T3=[]#Dados separados por sensor e convertidos para float 
    t=[]; _POPt=[]; t1=[]; tmlist=[] #Dados de tempo
    tm=IntVar(); tm_=StringVar(); tPS=IntVar(); tPE=StringVar(); Ts=StringVar(); Tp=StringVar(); Te=StringVar(); Q=StringVar(); pat=StringVar(); par=StringVar(); Vrs=StringVar() ; nteste=StringVar();
    _POPdados1=[]; _POPdados2=[]; _POPdados3=[]; _POPit=[]; _POPmedia1=[] #Dados para gráfico simultâneo
    flag_=[] #Resposta binária: flag (0 para valores de MEDIA abaixo do erro e 1 para valores de MEDIA acima do erro)
    
    #Valores instantâneos de temperatura, tempo, inhomogeneidade e média de inclinação 
    it_=[]; t_1=[]; t_2=[]; T1_1=[]; T1_2=[]; T2_1=[]; T2_2=[]; T3_1=[]; T3_2=[]; 
    T1incli1_=[]; T2incli1_=[]; T3incli1_=[]; media1_=[]
    
    #Parâmetros iniciais
    Erro= float(e5.get()) ; #Erro de temperatura para determinação do tPS:
    Pulsoduracao= float(e1.get()); #Duração do pulso do traçador (s)
    Vre = float(e2.get()); #Volume do reator (total) (m3)
    Vpulso = float(e3.get()); #Volume do traçador (m3)
    DeltaTemp = float(e4.get()); #Diferença de temperatura entre reator e traçador (°C)
    Nsensor = 3; #N° de sensores
    itmin=float(e6.get()) #Mínimo para função de inhomogeneidade
    tm_.set(0);    tm.set(2592000);    tPS.set(2592000);    dalayf=5
    
    #Gráfico
    fig=Figure(figsize=(10,4), dpi=100) #Tamanho e qualidade da figura
    plot=fig.add_subplot(1,1,1) #Figura = Gráfico em subplot do Matlib
    ax2 = plot.twinx() #Eixo secundário ao plot
    line1, = plot.plot(_POPt,_POPdados1, 'k-', label="T1") #Dados no gráfico
    line2, = plot.plot(_POPt,_POPdados2, 'k-', label="T2") 
    line3, = plot.plot(_POPt,_POPdados3, 'k-', label="T3")
    
    #Configurar radiobuttons
    if(v2.get()==0): s2r2.configure(state = DISABLED); s2r3.configure(state = DISABLED)
    if(v2.get()==1): s2r1.configure(state = DISABLED); s2r3.configure(state = DISABLED)
    if(v2.get()==2): s2r1.configure(state = DISABLED); s2r2.configure(state = DISABLED)
    
    #Linha horizontal e eixo secundário
    if(v1.get()==0):
        line4, = ax2.plot(_POPt,_POPit, 'c:', label="i(t)") 
        ax2.set_ylabel('i(t) (.%)', color='k') # Legenda de eixo y
        ax2.plot(_POPt,_POPit, 'c:', label="i(t)") 
        s1r1.configure(state = DISABLED)
    else: 
        line4, = ax2.plot(_POPt,_POPmedia1, 'c:', label="Inclinação") 
        ax2.set_ylabel('Inclinação', color='k') # Legenda de eixo y
        ax2.plot(_POPt,_POPmedia1, 'c:', label="i(t)") 
        s1r2.configure(state = DISABLED)
        
    #Eixos
    plot.set_xlabel('Tempo (s)', color='k') #Legenda de eixo x
    plot.set_ylabel('Temperatura (°C)', color='k') # Legenda de eixo y
    
    #Início da corrida
    l16.config(text='Conectando...')
    ser=serial.Serial(encon_arduino(), 9600) #Local e taxa de atualização
    
    def tmlabel(): #Função que retorna os valores de tempo de mistura na janela dependendo do tipo de teste
        if(j.get()==0 and v2.get()==0): txt.insert(INSERT,'     Tempo de mistura \n  (Teste único)= ' + str(round(float(tm_.get()), 2)) + ' s')
        if(j.get()==0 and v2.get()!=0): txt.insert(INSERT,'     Tempo de mistura \n   (Ensaio 1)= ' + str(round(float(tm_.get()), 2)) + ' s\n')
        if(j.get()==1 and v2.get()==1): txt.insert(INSERT,'\n     Tempo de mistura \n   (Ensaio 2)= ' + str(round(float(tm_.get()), 2)) + ' s \n\n       (Duplicata)\n      Média= ' + str(round((sum(tmlist)/len(tmlist)),2)) + ' s \n\n  Desvio Padrão= ' + str(round(statistics.stdev(tmlist),2)) + ' s')
        if(j.get()==1 and v2.get()!=1): txt.insert(INSERT,'\n     Tempo de mistura \n   (Ensaio 2)= ' + str(round(float(tm_.get()), 2)) + ' s')
        if(j.get()==2 and v2.get()==2): txt.insert(INSERT,'\n\n     Tempo de mistura \n  (Ensaio 3)= ' + str(round(float(tm_.get()), 2)) + ' s \n\n       (Triplicata)\n      Média= ' + str(round((sum(tmlist)/len(tmlist)),2)) + ' s \n\n  Desvio Padrão= ' + str(round(statistics.stdev(tmlist),2)) + ' s')       
        
    if encon_arduino() is not None: 
        ser.reset_input_buffer() #Reset buffer
        now = datetime.datetime.now()
        
        #Criar Pasta com resultados dependendo do
        if(v2.get()==0): caminho="Simples_Tempo_de_mistura_" + str(now.day) + "_" + str(now.month) + "_ " + str(now.year)
        if(v2.get()==1): caminho="Duplicata_Tempo_de_mistura_" + str(now.day) + "_" + str(now.month) + "_ " + str(now.year)
        if(v2.get()==2): caminho="Triplicata_Tempo_de_mistura_" + str(now.day) + "_" + str(now.month) + "_ " + str(now.year) 
        if os.path.exists(os.path.join(e8.get(), caminho)): pass #Pasta pré-existente (não criar pasta)
        else: os.mkdir(os.path.join(e8.get(), caminho)) #Criação da pasta no local(caminho) definido anteriormente dependendo do tipo de teste
        
        #Salvar dados no excel
        def salvar_dados_excel():
            now = datetime.datetime.now() #Aquisição dos valores de horário local
            if(v1.get()==0):
                df = pd.DataFrame(list(zip(t, T1, T2, T3, it_)), columns =['tempo (s)', 'T1 (°C)', 'T2 (°C)', 'T3 (°C)', 'i(t)']) #Definição do dataframe com os resultados 
            else:
                df = pd.DataFrame(list(zip(t, T1, T2, T3, media1_)), columns =['tempo (s)', 'T1 (°C)', 'T2 (°C)', 'T3 (°C)', 'MédiaInc']) #Definição do dataframe com os resultados 
            if(v1.get()==0): nteste.set("Corrida_Agua" + " - " + str(now.hour) + "h_"  + str(now.minute) + "min_" + str(now.second) + "s.xlsx") #Nome do arquivo a ser criado
            else: nteste.set("Corrida_Óleo" + " - " + str(now.hour) + "h_"  + str(now.minute) + "min_" + str(now.second) + "s.xlsx") #Nome do arquivo a ser criado
            path = os.path.join(e16.get(), caminho, nteste.get()) #Caminho do arquivo]
            writer=pd.ExcelWriter(path, engine='xlsxwriter') #Criação do arquivo
            df.to_excel(writer, sheet_name='Corrida_1', startrow=1) #Local na planilha
            #Geração do gráfico com os dados do sensor: Local, estilo e tamanho do gráfico
            workbook  = writer.book
            workbook.set_size(3600, 2400)
            chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth_with_markers'})
            chart.add_series({'name': '=Corrida_1!$C$2','categories': '=Corrida_1!$B$3:$B$5000','values': '=Corrida_1!$C$3:$C$5000', 'marker':{'type': 'circle','size': 5}})
            chart.add_series({'name': '=Corrida_1!$D$2','categories': '=Corrida_1!$B$3:$B$5000','values': '=Corrida_1!$D$3:$D$5000', 'marker':{'type': 'circle','size': 5}})
            chart.add_series({'name': '=Corrida_1!$E$2','categories': '=Corrida_1!$B$3:$B$5000','values': '=Corrida_1!$E$3:$E$5000', 'marker':{'type': 'circle','size': 5}})
            chart.add_series({'name': '=Corrida_1!$F$2','categories': '=Corrida_1!$B$3:$B$5000','values': '=Corrida_1!$F$3:$F$5000', 'y2_axis': 1, 'marker': {'type': 'none', 'color': 'black', 'size': 5}, 'line': {'dash_type': 'dash'}})
            chart.set_title({'name': 'Dados de temperatura'})
            chart.set_x_axis({'name': 'Tempo(s)'})
            chart.set_y_axis({'name': 'Temperatura (°C)'})
            chart.set_y2_axis({'name': 'i(t)'})
            chart.set_plotarea({'layout': {'x':0.14,'y':0.26,'width':0.55,'height': 0.6,}})
            chart.set_style(2)
            sheet = writer.sheets['Corrida_1']
            if(v1.get==0): sheet.insert_chart('I2', chart, {'x_offset': 10, 'y_offset': 10}) 
            else: sheet.insert_chart('L2', chart, {'x_offset': 10, 'y_offset': 10}) 
            writer.save() # Salvar planilha
        #Chamar gif
        Imagem2()
        
    #Iniciar contadores de tempo e flag
    start=time.time(); t.append(0); flag_.append(0);
    def Tmistura(i):
        if encon_arduino() is None: #Quando não for encontrado arduino
            #Imagem de Erro
            root.original7 = Image.open('error.PNG') #Nome da imagem - Local da imagem > pasta do arquivo
            resized7 = root.original7.resize((60, 60),Image.ANTIALIAS) #Mudança de escala
            root.image7 = ImageTk.PhotoImage(resized7) 
            root.display7 = Label(root, image = root.image7)
            root.display7.grid(row=2, column=3, rowspan=3, columnspan=13) #Local na janela
            root.display7.configure(background='white')
            #Legenda do Erro
            l30=Label (root, text="ERRO: Sensor \nnão encontrado\n Conecte o sensor", bd=5, relief="flat",
                  font= "Times 11",width=25, height=4, bg='gray9', fg='white')
            l30.grid (row=3, column=3, rowspan=5,columnspan=13)
            l16.config(text='Desconectado')
            ani.event_source.stop() #Parar animação
        else:
            #Garantir que a comunicação está ativa
            if(ser.isOpen() == False): ser.open()
            #Configurar entradas
            l16.config(fg="black"); l16.config(text='Conectado');
            root.protocol("WM_DELETE_WINDOW", disable_event) #Bloquear fechamento da janela
            if(int(t[-1])<tm.get()+dalayf):
                arduinoData=ser.readline().decode('utf8') #Decodificação dos dados do arduino
                end=time.time()
                DadosArray=arduinoData.split(' , ')
                T1.append(float(DadosArray[0])) #Dados de cada sensor separados 
                T2.append(float(DadosArray[1])) #(Objetivo: Salvar lista de temperaturas e 
                T3.append(float(DadosArray[2])) #possibilitar a atualização automática dos valores no gráfico ao vivo)
                t.append(end-start)
                #Mudar legenda 
                l18.config(text='T1='+str(T1[-1])); l19.config(text='T2='+str(T2[-1])); l20.config(text='T3='+str(T3[-1]))
                if(w.get()==0):
                    Ts.set((T1[-1]+T2[-1]+T3[-1])/3);
                    w.set(1)
                if(t[-2]<tPS.get()): #Início do tempo de mistura (Método das diferenças)
                    if(i>=1): #Valores instantâneos de temperatura e tempo 
                        media=(((T1[-2]+T2[-2]+T3[-2])/3))
                        if(v1.get()==0):it_.append(0)
                        else: media1_.append(0)
                        if(abs((T1[-1]-media)>Erro or abs(T2[-1]-media)>Erro or abs(T3[-1]-media)>Erro) 
                           and (T1[-1]>=T1[-2]+Erro or T2[-1]>=T2[-2]+Erro or T3[-1]>=T3[-2]+Erro)): #Flag
                            flag=1
                            flag_.append(flag)
                        else:
                            flag=0
                            flag_.append(flag)
                            
                        #Correção do flag para variação local de temperatura 
                        #Seleção de flag
                        if(T1[-1]>=T1[-2] or T2[-1]>=T2[-2] or T3[-1]>=T3[-2]):
                            if((flag_[-2]==1 and flag_[-1]==0) or (flag_[-2]==0 and flag_[-1]==0) or (flag_[-2]==0 and flag_[-1]==1)):
                                pass
                            else:
                                tPS.set(int(t[-2])) #Tempo de início do pulso tPS (s) 
                                e10.configure(state='normal'); e10.delete(0, 'end'); e10.insert(END, tPS.get()); l10.config(fg="black")
                                if(v1.get()==0): 
                                    tPE.set(tPS.get() + Pulsoduracao) #Tempo de fim do pulso tPE (s)
                                    e11.configure(state='normal'); e11.delete(0, 'end'); e11.insert(END, tPE.get()); l11.config(fg="black")
                                    Tp.set(float(Ts.get()) + DeltaTemp) #Temperatura do traçador (Tp) (°C)
                                    e8.configure(state='normal'); e8.delete(0, 'end'); e8.insert(END, round(float(Ts.get()),2)); l8.config(fg="black")
                                    e9.configure(state='normal'); e9.delete(0, 'end'); e9.insert(END, round(float(Tp.get()),2)); l9.config(fg="black")
                                    Vrs.set(Vre - Vpulso) #Volume do reator antes do pulso (m3)
                                    Q.set(Vpulso / Pulsoduracao) #Vazão do pulso (F) ( m3/s)
                                    e12.configure(state='normal'); e12.delete(0, 'end'); e12.insert(END, round(float(Q.get())*1000,3)); l12.config(fg="black")
                                    pat.set(((999.83952)+(16.945179*float(Tp.get()))-((7.9870401*10**-3)*(float(Tp.get())**2))-((46.170561*10**-6)*(float(Tp.get())**3))+((105.56302*10**-9)*(float(Tp.get())**4))-((280.54253*10**-12)*(float(Tp.get())**5)))/(1+((16.87985*10**-3)*float(Tp.get())))) #Densidade da água (Traçador)
                                    e14.configure(state='normal'); e14.delete(0, 'end'); e14.insert(END, round(float(pat.get()))); l14.config(fg="black")
                                    par.set(((999.83952)+(16.945179*float(Ts.get()))-((7.9870401*10**-3)*(float(Ts.get())**2))-((46.170561*10**-6)*(float(Ts.get())**3))+((105.56302*10**-9)*(float(Ts.get())**4))-((280.54253*10**-12)*(float(Ts.get())**5)))/(1+((16.87985*10**-3)*float(Ts.get())))) #Densidade da água (Reator)
                                    e15.configure(state='normal'); e15.delete(0, 'end');  e15.insert(END, round(float(par.get()))); l15.config(fg="black")
                                    Te.set(((float(pat.get())*Vpulso*float(Tp.get()))+(float(par.get())*float(Vrs.get())*float(Ts.get())))/((float(pat.get())*Vpulso)+(float(par.get())*float(Vrs.get())))) #Temperatura do meio após pulso (Te) (°C) - por balanço de energia
                                    e13.configure(state='normal'); e13.delete(0, 'end'); e13.insert(END, round(float(Te.get()),2)); l13.config(fg="black")
                                else: pass
                        #Atualizar lista de dados do gráfico
                        _POPdados1.append(T1[-2]);_POPdados2.append(T2[-2]);_POPdados3.append(T3[-2]);_POPt.append(t[-2]);
                        t1.append(t[-1]) 
                        if(v1.get()==0):_POPit.append(it_[-1])
                        else: _POPmedia1.append(media1_[-1])
                if(int(t[-2])>=tPS.get()):
                    if(v1.get()==0): #Método (MAYR et al., 1992) descrito para água
                        media=((T1[-2]+T2[-2]+T3[-2])/3)
                        #Valores da Curva de resposta ideal (M(t))
                        if (t[-2]<=float(tPE.get())): M=float(Tp.get())-((float(Tp.get())-float(Ts.get())))/(1+((float(Q.get())/float(Vrs.get()))*t[-2]))**(float(pat.get())/float(par.get()))
                        else:
                            M=float(Te.get())
                        #Cálculo do grau de heterogeneidade (i(t))  
                        st=(1/Nsensor)*abs((T1[-2]-M)+(T2[-2]-M)+(T3[-2]-M))
                        st2=(1/Nsensor)*abs((T1[-1]-M)+(T2[-1]-M)+(T3[-1]-M)) #Comparação entre dois valores (1) e (2) para garantia de flag
                        it=(st/(float(Te.get())-float(Ts.get())))
                        it2=(st2/(float(Te.get())-float(Ts.get()))) #Lista para comparação de i(t)
                        it_.append(it) # Lista dos valores para o gráfico
                        if(it<=itmin and k.get()==0 and it2<it):
                            tm.set(t[-2]) #Valor para linha vertical no gráfico
                            tm_.set((t[-2])-tPS.get()) #Cálculo do tempo de mistura
                            k.set(1) #Contador de execução única
                            print("\nTempo de Mistura =", float(tm_.get()),"s") 
                            tmlist.append(float(tm_.get()))  #Lista para exportação para excel
                            tmlabel() #Legendas dos resultados na janela
                        #Atualizar lista de dados do gráfico
                        _POPdados1.append(T1[-2]);_POPdados2.append(T2[-2]);_POPdados3.append(T3[-2]);_POPt.append(t[-2]);_POPit.append(it_[-1]); #Dados da lsta POP p/ gráfico
                        t1.append(t[-1]) 
                    else:
                        #Lista de valores de tempo utilizados na inclinação
                        t_1=[[t[-5],t[-4],t[-3],t[-2]]] 
                        t_2=[[t[-4],t[-3],t[-2],t[-1]]]
                        #Lista dos valores de temperatura utilizados na inclinação
                        T1_1=[T1[-5],T1[-4],T1[-3],T1[-2]];T2_1=[T2[-5],T2[-4],T2[-3],T2[-2]];T3_1=[T3[-5],T3[-4],T3[-3],T1[-2]] 
                        T1_2=[T1[-4],T1[-3],T1[-2],T1[-1]];T2_2=[T2[-4],T2[-3],T2[-2],T2[-1]];T3_2=[T3[-4],T3[-3],T3[-2],T1[-1]]
                        #Cálculo das inclinações (lineregress[0])
                        incli1_1=linregress(T1_1,t_1);incli2_1=linregress(T2_1,t_1);incli3_1=linregress(T3_1,t_1)
                        incli1_2=linregress(T1_2,t_2);incli2_2=linregress(T2_2,t_2);incli3_2=linregress(T3_2,t_2)
                        T1incli1=(incli1_1[0]);T2incli1=(incli2_1[0]);T3incli1=(incli3_1[0])
                        T1incli2=(incli1_2[0]);T2incli2=(incli2_2[0]);T3incli2=(incli3_2[0])
                        T1incli1_.append(T1incli1); T2incli1_.append(T2incli1); T3incli1_.append(T3incli1)
                        #Cálculo da média das inclinações
                        media1=(abs(T1incli1)+abs(T2incli1)+abs(T3incli1))/3
                        media2=(abs(T1incli2)+abs(T2incli2)+abs(T3incli2))/3
                        #Lista para construção do gráfico
                        media1_.append((abs(T1incli1)+abs(T2incli1)+abs(T3incli1))/3)
                        #Flag de média da inclinação em comparação com o erro indicado por e7.get())
                        if(abs((T1incli1-media1))>float(e7.get()) and abs((T1incli2-media2))>float(e7.get()) 
                           and abs((abs(T2incli1)-media1))>float(e7.get()) and abs((abs(T2incli2)-media2))>float(e7.get())
                           and abs((abs(T3incli1)-media1))>float(e7.get()) and abs((abs(T3incli2)-media2))>float(e7.get())): 
                            flag=0
                            flag_.append(flag)
                        else:
                            flag=1
                            flag_.append(flag)
                            if(k.get()==0):
                                tm.set(t[-2])
                                tm_.set((t[-2])-tPS.get())
                                print("\nTempo de Mistura =", float(tm_.get()),"s") 
                                k.set(1)
                                tmlist.append(float(tm_.get()))
                                tmlabel()    
                        _POPdados1.append(T1[-2]);_POPdados2.append(T2[-2]);_POPdados3.append(T3[-2]);_POPt.append(t[-1]);_POPmedia1.append(media1_[-2]) #Dados da lsta POP p/ gráfico
                        t1.append(t[-1]) 
                if(i>25): 
                    _POPdados1.pop(0);_POPdados2.pop(0);_POPdados3.pop(0);_POPt.pop(0); #Deletar dados p/ gráfico
                    if(v1.get()==0): _POPit.pop(0)
                    else: _POPmedia1.pop(0)
            else: #Fim do teste
                if(j.get()==v2.get()): salvar_dados_excel()
                k.set(0); w.set(0); j.set(j.get()+1); tm_.set(0); tm.set(2592000); tPS.set(2592000)
                if(j.get()>v2.get()):
                    ani.event_source.stop(); Imagem1() #Parar animação
                    ser.close() #Fechar porta serial
                    root.protocol("WM_DELETE_WINDOW", close_program)
                    #Configurar entradas
                    s1r1.configure(state = NORMAL); s1r2.configure(state = NORMAL); s2r1.configure(state = NORMAL); s2r2.configure(state = NORMAL); s2r3.configure(state = NORMAL)
                    l16.config(fg="gray"); l16.config(text='Desconectado')
            if(i>=1):
                media=(((T1[-2]+T2[-2]+T3[-2])/3))
                
                #Plot Gráfico
                #POP de dados e limite de de eixo
                if(int(t[-1])>25):
                    plot.set_xlim(_POPt[0], _POPt[-1]+5)
                else:
                    plot.set_xlim(_POPt[0], 30)
                plot.set_ylim(media-5,media+5)
                #Cores
                if(t[-1]>tm.get()):
                    line1.set_color('blue') 
                    line2.set_color('blue') 
                    line3.set_color('blue') 
                else:
                    if(t[-1]>=tPS.get()):   
                        line1.set_color('green') # Cor da linha
                        line2.set_color('green') # Cor da linha
                        line3.set_color('green') # Cor da linha
                    else:
                        line1.set_color('black') 
                        line2.set_color('black') 
                        line3.set_color('black') 
                line1.set_data(_POPt,_POPdados1) #Dados do gráfico
                line2.set_data(_POPt,_POPdados2) #Dados do gráfico
                line3.set_data(_POPt,_POPdados3) #Dados do gráfico
                plot.axvline(tPS.get(), color='b',linestyle='--') #Linha vertical de início da mistura
                plot.axvline(tm.get(), color='b',linestyle='--') #Linha vertical de fim da mistura
                if(v1.get()==0):
                    ax2.set_ylim(0, (max(_POPit)+(itmin*3)))
                    ax2.hlines(float(e6.get()), 0, 1000, color='r',linestyle=':')
                    line4.set_data(_POPt,_POPit) 
                else:
                    ax2.set_ylim(0, (max(media1_)+float(e7.get())*3))
                    line4.set_data(_POPt,_POPmedia1) 
    plotcanvas = FigureCanvasTkAgg(fig, root) #Gráfico
    plotcanvas.get_tk_widget().grid(column=0, row=0, columnspan=21, rowspan=10)
    ani = animation.FuncAnimation(fig, Tmistura, interval=785) #Animação da função Teste único
    
    #Definir botão de parada do teste
    def parar():
        ani.event_source.stop(); Imagem1() #Parar animação
        ser.close() #Fechar porta serial
        root.protocol("WM_DELETE_WINDOW", close_program)
        #Garantir que não haja perda de dados
        salvar_dados_excel()
        #Configurar entradas
        l16.config(fg="gray", text='Desconectado')
        s1r1.configure(state = NORMAL); s1r2.configure(state = NORMAL); s2r1.configure(state = NORMAL); s2r2.configure(state = NORMAL); s2r3.configure(state = NORMAL)
    #Botão de parada
    btn2 = Button(root, text="Desconectar", command=parar, bg='gray95', fg='black', width=15, height=2, justify=RIGHT, font="bold 8", bd=2)
    btn2.grid(column=16, row=10, rowspan=18, columnspan=4)
        
#Botão de execução de Tm
btn3 = Button(root, text="Conectar", command=runaniA, bg='gray95', fg='black', width=9, height=2, justify=RIGHT, font="bold 12", bd=2)
btn3.grid(column=16, row=7, rowspan=11, columnspan=4)
btn2 = Button(root, text="Desconectar", bg='gray95', fg='black', width=15, height=2, justify=RIGHT, font="bold 8", bd=2)
btn2.grid(column=16, row=10, rowspan=18, columnspan=4)

#Fim de execução da janela
root.mainloop()

# -*- coding: utf-8 -*-
"""
Created on Sun Nov 21 11:28:37 2021

@author: mrbra
"""

import os
import warnings
warnings.filterwarnings("ignore")

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd 
from PIL import ImageTk,Image
from pandastable import Table
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from wrangling.data_wrangling import Wrangler
from Data.Validation_excel import validation_excel, Excel_template
from tkinter_custom_button import TkinterCustomButton
from functools import reduce

plt.style.use('dark_background')

class App():
    
    def __init__(self):
        
       
        #Generate main window
        self.main_window = tk.Tk()
        #Set background image
        self.image2 = Image.open(os.path.join(os.path.dirname(__file__),'Images','foto1.png'))\
                                                            .resize((600,500))
        self.image1 = ImageTk.PhotoImage(self.image2)
        self.label = tk.Label(self.main_window, image = self.image1)
        self.label.pack()
        
        #Window title and dimensions
        self.main_window.title('Welcome to your personal dashboard')
        windowWidth = 600
        windowHeight = 500
        positionRight = int(self.main_window.winfo_screenwidth()/2 - windowWidth/2)
        positionDown = int(self.main_window.winfo_screenheight()/2 - windowHeight/2)-40

        self.main_window.resizable(width=False, height=False)

        self.main_window.geometry("{}x{}+{}+{}".format(windowWidth,
                                                       windowHeight,
                                                       positionRight,
                                                       positionDown))


        #Button to start App
        self.btn_main = TkinterCustomButton(master=self.main_window,
                                            text="START APP",
                                            corner_radius=0,
                                            fg_color="#CDB79E",
                                            hover_color="#EED5B7",
                                            command = lambda : self.openfile())
        
        self.btn_main.place(relx=0.5, rely=0.4, anchor=tk.CENTER) 
        
        self.btn_temp = TkinterCustomButton(master=self.main_window,
                                            text="DATA TEMPLATE",
                                            corner_radius=0,
                                            fg_color="#CDB79E",
                                            hover_color="#EED5B7",
                                            command = lambda : self.template_creation())
        
        self.btn_temp.place(relx=0.5, rely=0.6, anchor=tk.CENTER)        

# =============================================================================
#     Function to open and validate xlsx file 
# =============================================================================
    def openfile(self):
              
        self.filepath = fd.askopenfilename(filetypes=[("Excel files","*.xlsx")])
        num, mess = validation_excel(self.filepath)
        
        if self.filepath == '':
            tk.messagebox.showwarning("WARNING","Please, select a xlsx file")
        
        elif num == 0:
           
            tk.messagebox.showerror("ERROR",mess)
            
        else:    
            self.mainwindow()

    
    def template_creation(self):
        
        if Excel_template() == 0:
            tk.messagebox.showinfo("DONE","Template successfully exported to your Desktop")
        
        else:
            tk.messagebox.showerror("ERROR","Template has not been exported")
            


# =============================================================================
#    Function to open window with different metrics to execute
# =============================================================================
    def mainwindow(self):
        
        #Window config
        newWindow = tk.Toplevel(self.main_window)
        
        newWindow.title('Investment dashboard metrics')
        newWindow.geometry("600x500")
        newWindow.resizable(width=False, height=False)
        
        #Image background
        image3 = Image.open(os.path.join(os.path.dirname(__file__),'Images','foto2.png'))\
                                                            .resize((600,500))
        image4 = ImageTk.PhotoImage(image3)
        label2 = tk.Label(master=newWindow, image = image4)
        label2.pack()
        label2.image = image4
        
        #Buttons
        btn = TkinterCustomButton(master =newWindow,text="TABLE VIEW",
                                  corner_radius=0,
                                  fg_color="#CDB79E",
                                  hover_color="#EED5B7",
                                  command = lambda : self.openNewWindow())
        
        btn.place(relx=0.3, rely=0.5, anchor=tk.CENTER)
        
        btn2 = TkinterCustomButton(master =newWindow, text="CUM. RETURN",
                                   corner_radius=0,
                                   fg_color="#CDB79E",
                                   hover_color="#EED5B7",
                                   command = lambda : self.openNewWindow2())
        
        btn2.place(relx=0.3, rely=0.7, anchor=tk.CENTER)
        
        btn3 = TkinterCustomButton(master =newWindow, text="LAST PICTURE",
                                   corner_radius=0,
                                   fg_color="#CDB79E",
                                   hover_color="#EED5B7",
                                         command = lambda : self.openNewWindow3())
        
        btn3.place(relx=0.3, rely=0.3, anchor=tk.CENTER)
        
        btn4 = TkinterCustomButton(master =newWindow,text="SHARE PRICE",
                                   corner_radius=0,
                                   fg_color="#CDB79E",
                                   hover_color="#EED5B7",
                                   command = lambda : self.openNewWindow4())
        
        btn4.place(relx=0.7, rely=0.3, anchor=tk.CENTER)      
        
        btn5 = TkinterCustomButton(master =newWindow,text="COMPOSITION",
                                   corner_radius=0,
                                   fg_color="#CDB79E",
                                   hover_color="#EED5B7",
                                   command = lambda : self.openNewWindow5())
        
        btn5.place(relx=0.7, rely=0.5, anchor=tk.CENTER) 
        
        btn6 = TkinterCustomButton(master =newWindow,text="PORTFOLIO VOL.",
                                   corner_radius=0,
                                   fg_color="#CDB79E",
                                   hover_color="#EED5B7",
                                   command = lambda : self.openNewWindow6())
        
        btn6.place(relx=0.7, rely=0.7, anchor=tk.CENTER)         

        
# =============================================================================
#  Functions for each of the metrics 
# =============================================================================
    def openNewWindow(self):
        
        table = Wrangler(self.filepath).Agreg()
        newWindow = tk.Toplevel(self.main_window)
    
        newWindow.title("Table view")
 
        newWindow.geometry("800x600")
                 
        pt = Table(newWindow, dataframe=table,showtoolbar=True, showstatusbar=True)
        pt.show()
        
    def openNewWindow2(self):
        
        table = Wrangler(self.filepath).Agreg()
        newWindow = tk.Toplevel(self.main_window)
    
        newWindow.title("Total cumulative return of your fund")
 
        newWindow.geometry("800x600")
        
        fig, ax  = plt.subplots()
        
        ax.plot(table['Date'],
                table['Total_return_fund']*100,
                'lightsteelblue',table['Date'],table['Returns_SP500']*100,
                'antiquewhite')
                                  
        plt.title('CUMULATIVE RETURN VS SP500 CUMULATIVE RETURN')
        plt.legend(['Personal fund','SP500'])
        ax.set_xlabel('Date', fontsize=18)
        ax.set_ylabel('Percentage', fontsize=16)
        ax.yaxis.set_major_formatter(mtick.PercentFormatter())
        plt.grid(True)
        canvas = FigureCanvasTkAgg(fig, newWindow)
        canvas._tkcanvas.pack(fill=tk.BOTH, expand=1)    
        
    def openNewWindow3(self):
        
        table = Wrangler(self.filepath).Agreg()
        table = table.iloc[-1,[0,15,19,20,21]].tolist()
        table = [ round(elem, 3) if type(elem) == np.float64 else elem.strftime("%d/%m/%Y") for elem in table]
             
        newWindow = tk.Toplevel(self.main_window)    
        newWindow.title("Last updated metrics") 
        newWindow.geometry("1000x600")
        tk.Label(newWindow, text="MAIN METRICS", font=("Arial",30)).grid(row=0, columnspan=5)
        
        # create Treeview with 5 columns
        cols = ('Date', 'Total Market Value','Total investment', 'Price','Total Shares')
        listBox = ttk.Treeview(newWindow, columns=cols, show='headings')
        tempList = [table]
        tempList.sort(key=lambda e: e[0], reverse=True)

        for date, mkt_value,t_inv, price, shares in tempList:
            listBox.insert("", "end", values=(date, mkt_value,t_inv, price, shares))
            
        # set column headings
        for col in cols:
            listBox.heading(col, text=col)    
            listBox.grid(row=1, column=0, columnspan=2)
            
    def openNewWindow4(self):
        
        table = Wrangler(self.filepath).Agreg()
        newWindow = tk.Toplevel(self.main_window)
    
        newWindow.title("Share price of your fund")
 
        newWindow.geometry("800x600")
 
        #ttk.Label(newWindow,text ="This is a new window").pack()
        fig, ax = plt.subplots()
        ax.plot(table['Date'],table['Price'],'lightsteelblue')
        plt.title('SHARE PRICE')
        plt.grid(True)
        canvas = FigureCanvasTkAgg(fig, newWindow)
        canvas._tkcanvas.pack(fill=tk.BOTH, expand=1)
        
    def openNewWindow5(self):
        
        table = Wrangler(self.filepath).Agreg()
        newWindow = tk.Toplevel(self.main_window)
    
        newWindow.title("Portfolio composition")
 
        newWindow.geometry("900x600")
 
        #ttk.Label(newWindow,text ="This is a new window").pack()
        
        fig, ax = plt.subplots()
        ax.yaxis.set_major_formatter(mtick.PercentFormatter())
        
        data_pie = table.filter(regex='Weight_').iloc[-1,:]
        
        index = table.filter(regex='Weight_').columns
    
        _,_,m=ax.pie(data_pie,labels=index,autopct='%1.1f%%',textprops={'size': 'x-large','color':'w'})
        [m[i].set_color('black') for i in range(len(m))]
        
        plt.title('PORTFOLIO COMPOSITION', fontsize=16)
        plt.grid(True)
        canvas = FigureCanvasTkAgg(fig, newWindow)
        canvas._tkcanvas.pack(fill=tk.BOTH, expand=1)
        
    def openNewWindow6(self):

        newWindow = tk.Toplevel(self.main_window)
    
        newWindow.title("Portfolio volatility")
 
        newWindow.geometry("800x600")
           
        table = Wrangler(self.filepath).Agreg()
        
        #calculate volatility for SP500
        vol_sp500 = table['Returns_SP500'].rolling(10).std()
        #select needed variables for portfolio volatility calculation
        df_sele = table.filter(regex='valor_liquidativo_|Weight|Date')
        
        #get name of assets
        columns_sele = df_sele.filter(regex='valor_liquidativo').columns 
        names = [i[18:] for i in columns_sele] #names of assets
        
        #calculate returns of your assets
        df_sele.loc[:,columns_sele] = df_sele.loc[:,columns_sele].pct_change()
        
              
# =============================================================================
#     Calculate portfolio volatility for 10d window    
# =============================================================================
        
        #create all combinations of assets and place it into a list             
        list_comb = []
            
        for i in names:
            for j in names:
                list_comb.append(sorted([i,j]))                   
           
        list_comb = set([tuple(i) for i in list_comb])
            
        list_fin = []
        
        #for all combinations calculate COV and WEIGHT (check V(A+X+...+N) formulae)
        #and save the results into a new list
                  
        for z in list_comb:
            
            list_fin.append(df_sele['valor_liquidativo_'+str(z[0])].rolling(10).cov(df_sele['valor_liquidativo_'+str(z[1])])*df_sele['Weight_'+str(z[0])].rolling(10).mean()*df_sele['Weight_'+str(z[1])].rolling(10).mean())
           
        #sum all rows of the series allocated in the list to get portfolio vol.  
        d = reduce(lambda x,y : x.add(y,fill_value=0),list_fin)
        d = d.apply(lambda x : np.sqrt(x))
        

        fig, ax = plt.subplots()   
        ax.plot(df_sele['Date'], d,'-r', df_sele['Date'], vol_sp500,'-b')
        ax.tick_params('x', labelrotation=45,labelsize=7)
        plt.title('PORTFOLIO VOLATILITY 10d', fontsize=16)
        plt.grid(True)
        plt.legend(['Portfolio','SP500'])
        canvas = FigureCanvasTkAgg(fig, newWindow)
        canvas._tkcanvas.pack(fill=tk.BOTH, expand=1)
 
            
app = App()

app.main_window.mainloop()





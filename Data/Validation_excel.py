# -*- coding: utf-8 -*-
"""
Created on Wed Dec 15 22:56:35 2021

@author: mrbra
"""
import pandas as pd
import xlsxwriter
import os
import numpy as np

def validation_excel(file):
    
    df_keys = pd.read_excel(file,None)
    
    if set(df_keys.keys()) != set (['Transactions','Tickers']):
        return (0, 'Please use valid sheet names')   
    
    else:
        
        df = pd.read_excel(file,sheet_name = 'Transactions')
        df['Amount'] = df['Amount'].astype(np.float64)
        df['Price'] = df['Price'].astype(np.float64)     
        df['Number_shares'] = df['Number_shares'].astype(np.float64) 
        df = df[df.columns.drop(list(df.filter(regex='Unnamed')))]
        df2 = pd.read_excel(file,sheet_name = 'Tickers')
        df2 = df2[df2.columns.drop(list(df2.filter(regex='Unnamed')))]                   

        if set(df.columns) != set(['Code','Description','Brokerage',
                                     'Type of transaction','Date','Amount',
                                     'Number_shares','Price',
                                     'Product']):
            
            return (0,'Please use valid column names in Transactions sheet')
        

        elif df.isnull().values.any() == True or df2.isnull().values.any() == True:
            return (0, 'Please fill all the entries for any row in the sheets')
        
        elif df['Code'].dtype != 'int64' or df['Description'].dtype != 'object' or df['Brokerage'].dtype != 'object' or df['Type of transaction'].dtype != 'object' or df['Date'].dtype != 'datetime64[ns]' or df['Amount'].dtype != 'float64' or df['Number_shares'].dtype != 'float64' or df['Price'].dtype != 'float64' or df['Product'].dtype != 'object':
                    return (0,'Please use valid data formats in Transactions sheet')
        
        elif set(df2.columns) != set(['Description','Tickers_YF']):
            return (0,'Please use valid column names in Tickers_YF sheet')
        
        elif df2['Description'].dtype != 'object' or df2['Tickers_YF'].dtype != 'object':
            return (0,'Please use valid data formats in Tickers_YF sheet')
        
        elif set(np.delete(df['Description'].unique(),np.argwhere(df['Description'].unique() == 'Cash'))) != set(np.delete(df2['Description'].unique(),np.argwhere(df2['Description'].unique() == 'Cash'))):
            return (0,'Please use same Description names in both sheets')
        
        elif set(df['Type of transaction']) != set(['Buy','Sell']):
            return (0, 'Please use \'Buy\' or \'Sell\' in \'Type of transaction field\'')
        
        elif set(df.loc[df['Type of transaction']=='Buy']['Number_shares'].values > 0) == set([False,True]) or set(df.loc[df['Type of transaction']=='Buy']['Number_shares'].values > 0) == set([False]):
            return (0, 'Please use valid data for \'Number_shares\' in \'Buy\' transaction')

        elif set(df.loc[df['Type of transaction']=='Buy']['Amount'].values > 0) == set([False,True]) or set(df.loc[df['Type of transaction']=='Buy']['Amount'].values > 0) == set([False]):
            return (0, 'Please use valid data for \'Amount\' in \'Buy\' transaction')

        
        elif set(df.loc[df['Type of transaction']=='Sell']['Number_shares'].values < 0) == set([False,True]) or set(df.loc[df['Type of transaction']=='Sell']['Number_shares'].values < 0) == set([False]):
            return (0, 'Please use valid data for \'Number_shares\' in \'Sell\' transaction')

        elif set(df.loc[df['Type of transaction']=='Sell']['Amount'].values < 0) == set([False,True]) or set(df.loc[df['Type of transaction']=='Sell']['Amount'].values < 0) == set([False]):
            return (0, 'Please use valid data for \'Amount\' in \'Sell\' transaction')
                        
        else:
            return (1,'')
            

def Excel_template():
    
    try:
        
        workbook = xlsxwriter.Workbook(os.path.join(
                                        os.path.join(
                                        os.path.join(
                                        os.environ['USERPROFILE']), 'Desktop'),
                                        'Investment_dashboard_template.xlsx'))
        
        #Formatos
        bold_color = workbook.add_format({'bold': True, 'bg_color': 'yellow',
                                          'align':'center_across'})
        date_format = workbook.add_format()
        date_format.set_num_format('dd/mm/yyyy')
        num_format = workbook.add_format()
        num_format.set_num_format('#.##0,00')

        
        worksheet = workbook.add_worksheet('Transactions')
        worksheet.hide_gridlines(2)        
        worksheet.set_column('A:I', 20)
        worksheet.set_column('E:E',20,date_format)
        worksheet.set_column('F:H',20,num_format)
        
        worksheet.write(0,0,'Code',bold_color)
        worksheet.write(0,1,'Description',bold_color)
        worksheet.write(0,2,'Brokerage',bold_color)
        worksheet.write(0,3,'Type of transaction',bold_color)
        worksheet.write(0,4,'Date',bold_color)
        worksheet.write(0,5,'Amount',bold_color)
        worksheet.write(0,6,'Number_shares',bold_color)
        worksheet.write(0,7,'Price',bold_color)
        worksheet.write(0,8,'Product',bold_color)
        
        #Second sheet
        worksheet2 = workbook.add_worksheet('Tickers')
        worksheet2.hide_gridlines(2)
        worksheet2.set_column('A:B', 20)
        worksheet2.write(0,0,'Description', bold_color)
        worksheet2.write(0,1,'Tickers_YF', bold_color)
        workbook.close()
    
        return 0
    
    except:
        
        return 1
    
    

    
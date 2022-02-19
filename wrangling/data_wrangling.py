# -*- coding: utf-8 -*-
"""
Created on Thu Nov 11 22:37:13 2021

@author: mrbra
"""

import pandas as pd
from datetime import timedelta
import yfinance as yf

class Wrangler:
 
# =============================================================================
#          Initiate the Wrangler class with Nshares and amount time series
# =============================================================================

    def __init__(self, path):
        
        self.path = path
        data = pd.read_excel(self.path,index_col=0,sheet_name=
                                  'Transactions')
        agreg = data.groupby(['Description','Date'])[['Number_shares','Amount']]\
                                                        .sum()\
                                                        .groupby(level=0)\
                                                        .cumsum().reset_index()
            
        

        #Eliminate assets that have 'number_shares' equal to zero (Assets that
        #are no longer in your portfolio)
        
        list_out = agreg.loc[agreg['Number_shares'] == 0.0]['Description']\
                                                            .to_list()
        
        for i in list_out:
            
            agreg.drop(agreg[agreg['Description'] == i ].index, inplace=True)
            agreg = agreg.reset_index(drop=True)

        #Generate table with all dates     
        edate = pd.Timestamp.now()
        edate = edate.replace(minute=0,second=0)
        sdate = agreg['Date'].min()
        Dates_full = pd.date_range(sdate,edate-timedelta(days=0),freq='d')\
                                                                .to_series(name='Date')\
                                                                .reset_index(drop=True)

        list_ = []
        
        for i in agreg['Description'].unique():
            
            frame_iter = agreg.loc[agreg['Description']==i,:]\
                                    .merge(Dates_full,on='Date',how='right')
                                    
            frame_iter.drop(['Description'],axis=1,inplace=True)
            
            frame_iter.rename(columns={'Number_shares':'Number_shares_'+i,
                                       'Amount':'Amount_investment_'+i},inplace=True)
            
            frame_iter.fillna(method='ffill',inplace=True)
            
            list_.append(frame_iter)
            

        full_table = list_[0]

        for i in range(1,len(list_)):
            
            full_table = full_table.merge(list_[i],on='Date',how='left')

        self.agreg = agreg
        self.full_table = full_table
        
# =============================================================================
#          Function to retrieve your assets' market prices from YF API
# =============================================================================

    def Mkt_value(self):
        
        df_tickers = pd.read_excel(self.path,
                                   sheet_name='Tickers')
        list_vl = []
        
        tickers_ = df_tickers.Tickers_YF.unique()
        
        for i,j in zip(tickers_,range(len(tickers_))):
            
            list_vl.append(yf.Ticker(i).history(period='max'))
            
            list_vl[j].reset_index(inplace=True)
            
            list_vl[j].rename(columns={'Close':'valor_liquidativo_'+df_tickers.loc[
                df_tickers['Tickers_YF']==i,'Description'].item(),'Date':'Date'},
                                   inplace=True)
        
            list_vl[j] = list_vl[j].filter(regex='Date|valor')
            
        return list_vl 

# =============================================================================
#        Function to retrieve market price of SP500 Index (SP500 is used as benchmark index)   
# =============================================================================
    def SP500_data(self):
        
        df_sp500 = yf.Ticker('^GSPC').history('10y')
        df_sp500.reset_index(inplace=True)
        df_sp500.rename(columns={'Close':'SP500_Index','Date':'Date'}, inplace=True)
        return df_sp500
 
# =============================================================================
#        Join retrieved data in a table
# =============================================================================
    def Agreg(self):
        
        #Call portfolio market prices function
        mkt_value = self.Mkt_value()
        full_table = self.full_table
    
        for i in range(len(mkt_value)):
            
            full_table = full_table.merge(mkt_value[i],on='Date',how='left')
        
        full_table[full_table.filter(
            regex='valor_liquidativo').columns.tolist()]=full_table[full_table.filter(regex='valor_liquidativo').columns.tolist()].fillna(method='ffill')
        
        #Call SP500 function
        full_table = full_table.merge(self.SP500_data()[['Date','SP500_Index']],on='Date',how='left')
        full_table['SP500_Index'] = full_table['SP500_Index'].fillna(method='ffill')
        full_table['Returns_SP500'] = (full_table['SP500_Index']/full_table['SP500_Index'].iloc[0])-1

   
        
        #Total market value for each of your assets
        list_assets = self.agreg['Description'].unique()

        for i in list_assets:
            
            if i != 'Cash':
                full_table['Market_value_'+i] = full_table.filter(regex='valor_liquidativo_'+i).values*full_table.filter(regex='Number_shares_'+i).values
                
            else:
                full_table['Market_value_'+i] = full_table['Amount_investment_Cash']

        #Agg market value of your portfolio and weights
        full_table['Total_market_value'] = full_table.filter(regex='Market_value').sum(axis=1) 
  
        for i in list_assets:
            full_table['Weight_'+i] = full_table['Market_value_'+i] / full_table['Total_market_value']
            
        #Total value of investment disbursed
        full_table['Total_investment_disbursed'] = full_table.filter(regex='Amount_investment').sum(axis=1)
        
        #Investment fund participations and share price dynamics calculated as a loop
        inv_desem = full_table['Total_investment_disbursed'].values
        mkt_value = full_table['Total_market_value'].values
        participations = [200] #Initial participations are set to 200. Trivial number
        vliq_ = [mkt_value[0]/participations[0]] #Initial share price of your fund
        
        for i in range(1,len(mkt_value)):
            participations.append(participations[i-1] + (inv_desem[i]-inv_desem[i-1])/vliq_[i-1])
            vliq_.append(mkt_value[i]/participations[i])
        
        full_table['Price'] = vliq_
        full_table['Amount_fund_shares'] = participations
        
        #Total return fund
        full_table['Total_return_fund'] = (full_table['Price']/full_table['Price'].values[0])-1
        
        self.full_table = full_table
        
        return self.full_table
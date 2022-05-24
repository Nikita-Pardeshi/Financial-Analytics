from io import BytesIO
import plotly
import PIL
from PIL import Image
from PIL import Image, ImageFile, ImageFont, ImageDraw
ImageFile.LOAD_TRUNCATED_IMAGES = True
import plotly.graph_objects as go
import plotly.figure_factory as ff
from plotly.subplots import make_subplots
from reportlab import platypus
from reportlab.lib.styles import ParagraphStyle as PS
from reportlab.lib.pagesizes import letter, inch
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.legends import Legend
from reportlab.lib.validators import Auto
from reportlab.graphics.shapes import Drawing, String
from reportlab.platypus import *
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch, mm,cm
from reportlab.lib.pagesizes import inch, A4, landscape
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT
import io 
import random, time, datetime
from django.http import HttpResponse
from PyPDF2 import PdfFileReader,PdfFileWriter,PdfFileMerger
import textwrap
import PyPDF2
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import Table, TableStyle


#NEW imports 
#from beautifulsoup import BeautifulSoup as bs4
import requests 
from googlesearch import search
import yahoo_fin.stock_info as si

from urllib import request
import wikipedia
from hashlib import sha1

import json
import newsapi
from newsapi import NewsApiClient

newsapi = NewsApiClient(api_key='7db51e8f7dd84ae89bb119a83d2bda2b')

#from kaleido.scopes.plotly import PlotlyScope
import plotly.io as pio 

styles = getSampleStyleSheet()
style_left = PS(name='left', parent=styles['Normal'], alignment=TA_CENTER)


h1 = PS(name='Heading1', fontName='Times-Bold', fontSize=14, leading=28,alignment=TA_JUSTIFY)
h1_1=PS(name='Heading11', fontName='Times-Roman', fontSize=14, leading=16,alignment=TA_JUSTIFY)
h2 = PS(name='Heading2', fontName='Times-Roman', fontSize=12, leading=16, leftIndent=16,alignment=TA_JUSTIFY)
h3 = PS(name='Heading3', fontName='Times-Roman', fontSize=12, leading=16, leftIndent=30,alignment=TA_JUSTIFY)
h4 = PS(name='Heading4', fontName='Times-Roman',fontSize=12, leading=16, leftIndent=50,alignment=TA_JUSTIFY)
h5 = PS(name='Heading5', fontName='Times-Roman', fontSize=12, leading=16, leftIndent=80,alignment=TA_JUSTIFY)
h6=PS(name='Heading6', fontName='Times-Roman', leading=16,fontSize=12,alignment=TA_JUSTIFY)
center=PS(name='center', fontName='Times-Roman', leading=16,fontSize=16,alignment=TA_CENTER)




def merge(data):

    def name_convert(self):

        #search for company on yahoo finance
        searchval = 'yahoo finance '+self
        link = []
        #limits to the first link
        for url in search(searchval, tld='com', lang='en', stop=1):
            link.append(url)
        link = str(link[0])
        link=link.split("/")

        #extract ticker from Yahoo finance URL
        if link[-1]=='':
            ticker=link[-2]
        else:
            x=link[-1].split('=')
            ticker=x[-1]
        if ticker=='history':
            ticker=link[-3]

        return(ticker)

    #extract company name
    company_name=data['name']
    company=name_convert(company_name)
    print('Ticker: ',company)

    #Get information about the company from Wikipedia to be included in the final pdf
    about=''
    try:
        about=wikipedia.summary(company_name)
    except wikipedia.exceptions.PageError:
        pass
    except wikipedia.exceptions.HTTPTimeoutError(company_name):
        print ("Error Connecting the server. Please reload")

    comp_link=[]
    for url1 in search(company_name, tld='com', lang='en', stop=1):
        comp_link.append(url1)

    #get income statement and balace sheet of the company
    income_statement = si.get_income_statement(company)
    balance_sheet = si.get_balance_sheet(company)
    i_transposed = income_statement.transpose()
    b_transposed = balance_sheet.transpose()

    #Extract company specific news
    n_title=[]
    n_desc=[]
    n_url=[]
    n_date=[]
    news_data=newsapi.get_everything(q=company_name,page=1,page_size=100,language='en')

    articles=news_data['articles']
    if (articles)!=0:
        for i in range(0,len(articles)):
            while i<=int(len(articles)/2):
                t=articles[i]['title']
                n_title.append(t)
                des=articles[i]['description']
                n_desc.append(des)
                link=articles[i]['url']
                n_url.append(link)
                d=articles[i]['publishedAt'][0:10]
                n_date.append(d)  

    #extract dates 
    objects=income_statement.columns
    arr=objects.date
    year=arr.tolist()
    dates=[]
    for i in year:
        i1=str(i)
        date=i1.replace("datetime.date(","").replace(")","").replace(", ",":")
        dates.append(date)

    graph_dates=[]
    for i in dates: 
        numeric=int(i[0:4])
        graph_dates.append(numeric)

    #vertical analysis 
    
    gross_profit=i_transposed['grossProfit'].to_list()
    revenue=i_transposed['totalRevenue'].to_list()
    net_earnings=i_transposed['netIncome'].to_list()
    if 'sellingGeneralAdministrative' in i_transposed:
        sga=i_transposed['sellingGeneralAdministrative'].to_list()
    if 'incomeBeforeTax' in i_transposed:
        ebt=i_transposed['incomeBeforeTax'].to_list()
    if 'costOfRevenue' in i_transposed:
        cogs=i_transposed['costOfRevenue'].to_list()

    p_revenue=[]
    p_revenue_str=[]
    for y in range(0,len(dates)):
        p_revenue.append(100)
        val="<b>100%</b>"
        p_revenue_str.append(val)
        
    
    percent_cogs=[]
    percent_gross_profit=[]
    percent_sga=[]
    percent_ebt=[]
    percent_net_earnings=[]

    graph_cogs=[]
    graph_gross_profit=[]
    graph_sga=[]
    graph_ebt=[]
    graph_net_earnings=[]

    revenue_mn=[]
    cogs_mn=[]
    gross_profit_mn=[]
    sga_mn=[]
    ebt_mn=[]
    net_earnings_mn=[]
    for i in range(0,len(dates)):
        if 'costOfRevenue' in i_transposed:
            cogs_int=round((cogs[i]/revenue[i])*100,2)
        if gross_profit is not None:
            gross_profit_int=round((gross_profit[i]/revenue[i])*100,2)
        if 'sellingGeneralAdministrative' in i_transposed:
            sga_int=round((sga[i]/revenue[i])*100,2)
        if ebt is not None:
            ebt_int=round((ebt[i]/revenue[i])*100,2)
        if net_earnings is not None:
            net_earnings_int=round((net_earnings[i]/revenue[i])*100,2)
        
        if 'costOfRevenue' in i_transposed:
            graph_cogs.append(cogs_int)
            cogs_value=str(cogs_int)+"%"
            percent_cogs.append(cogs_value)
            c_mn=str(round((cogs[i]/1000000000),2))
            cogs_mn.append(c_mn)  
        
        if gross_profit is not None:

            graph_gross_profit.append(gross_profit_int)
            gross_profit_value="<b>"+str(gross_profit_int)+"%</b>"
            percent_gross_profit.append(gross_profit_value)
            gp_mn="<b>"+str(round((gross_profit[i]/1000000000),2))+"</b>"
            gross_profit_mn.append(gp_mn)

        if 'sellingGeneralAdministrative' in i_transposed:
            graph_sga.append(sga_int)
            sga_value=str(sga_int)+"%</b>"
            percent_sga.append(sga_value)
            s_mn=str(round((sga[i]/1000000000),2))
            sga_mn.append(s_mn)

        if ebt is not None:    
            graph_ebt.append(ebt_int)
            ebt_value="<b>"+str(ebt_int)+"%</b>"
            percent_ebt.append(ebt_value)
            e_mn="<b>"+str(round((ebt[i]/1000000000),2))+"</b>"
            ebt_mn.append(e_mn)
        
        if net_earnings is not None:
            graph_net_earnings.append(net_earnings_int)
            net_earnings_value="<b>"+str(net_earnings_int)+"%</b>"
            percent_net_earnings.append(net_earnings_value)
            ne_mn="<b>"+str(round((net_earnings[i]/1000000000),2))+"</b>"
            net_earnings_mn.append(ne_mn)

        r_mn="<b>"+str(round((revenue[i]/1000000000),2))+"</b>"
        revenue_mn.append(r_mn) 
        

    final_revenue=revenue_mn+p_revenue_str
    final_revenue.insert(0,'<b>Revenue<br>(Billion)</b>')
    if 'costOfRevenue' in i_transposed:
        final_cogs=cogs_mn+percent_cogs
        final_cogs.insert(0,'<b>Cost of<br>Goods<br>Sold</b>')
    final_gross_profit=gross_profit_mn+percent_gross_profit
    final_gross_profit.insert(0,'<b>Gross<br>Profit</b>')
    if 'sellingGeneralAdministrative' in i_transposed:
        final_sga=sga_mn+percent_sga
        final_sga.insert(0,'<b>Selling<br>General<br>and<br>administ<br>rative</b>')
    if ebt is not None:
        final_ebt=ebt_mn+percent_ebt
        final_ebt.insert(0,'<b>Earnings<br>before<br>Tax</b>')
    if net_earnings is not None:
        final_net_earnings=net_earnings_mn+percent_net_earnings
        final_net_earnings.insert(0,'<b>Net<br>Earnings</b>')

    #horizontal analysis 
    h_revenue=[]
    h_cogs=[]
    h_gross_profit=[]
    h_sga=[]
    h_ebt=[]
    h_net_earnings=[]

    graph_h_revenue=[]
    graph_h_cogs=[]
    graph_h_gross_profit=[]
    graph_h_sga=[]
    graph_h_ebt=[]
    graph_h_net_earnings=[]

    for i in range(0,len(dates)):
        
        if i<len(dates)-1:
            yoy_r_int=round(((revenue[i]/revenue[i+1])-1)*100,2)
            if len(cogs)!=0:
                try:
                    yoy_cogs_int=round(((cogs[i]/cogs[i+1])-1)*100,2)
                    yoy_cogs= str(yoy_cogs_int)+"%"
                    h_cogs.append(yoy_cogs)
                    graph_h_cogs.append(yoy_cogs_int)
                except ZeroDivisionError as error:
                    h_cogs.append("NA")
                    

            yoy_gross_int=round(((gross_profit[i]/gross_profit[i+1])-1)*100,2)
            if len(sga)!=0:
                yoy_sga_int=round(((sga[i]/sga[i+1])-1)*100,2)
                yoy_sga= str(yoy_sga_int)+"%"
                h_sga.append(yoy_sga)
                graph_h_sga.append(yoy_sga_int)
            if len(ebt)!=0:
                yoy_ebt_int=round(((ebt[i]/ebt[i+1])-1)*100,2)
                yoy_ebt= "<b>"+str(yoy_ebt_int)+"%</b>"
                h_ebt.append(yoy_ebt)
                graph_h_ebt.append(yoy_ebt_int)
            if len(net_earnings)!=0:
                yoy_net_int=round(((net_earnings[i]/net_earnings[i+1])-1)*100,2)
                yoy_net_earnings= "<b>"+str(yoy_net_int)+"%</b>"
                h_net_earnings.append(yoy_net_earnings)
                graph_h_net_earnings.append(yoy_net_int)

            yoy_r= "<b>"+str(yoy_r_int)+"%</b>"
            yoy_gross_profit= "<b>"+str(yoy_gross_int)+"%</b>"
            h_revenue.append(yoy_r)
            h_gross_profit.append(yoy_gross_profit)

            graph_h_revenue.append(yoy_r_int)
            graph_h_gross_profit.append(yoy_gross_int)
        else:
            break
       
    h_revenue.append("NA")
    if 'costOfRevenue' in i_transposed:
        h_cogs.append("NA")
        h_final_cogs= cogs_mn+ h_cogs
        h_final_cogs.insert(0,'<b>Cost of<br>Goods<br>Sold</b>')
    if gross_profit is not None:
        h_gross_profit.append("NA")
        h_final_gross_profit= gross_profit_mn+ h_gross_profit
        h_final_gross_profit.insert(0,'<b>Gross<br>Profit</b>')
    if 'sellingGeneralAdministrative' in i_transposed:
        h_sga.append("NA")
        h_final_sga= sga_mn+ h_sga
        h_final_sga.insert(0,'<b>Selling<br>General<br>and<br>administ<br>rative</b>')
    
    if ebt is not None:
        h_ebt.append("NA")
        h_final_ebt= ebt_mn+ h_ebt
        h_final_ebt.insert(0,'<b>Earnings<br>before<br>Tax</b>')
    if net_earnings is not None:
        h_net_earnings.append("NA")
        h_final_net_earnings= net_earnings_mn+ h_net_earnings
        h_final_net_earnings.insert(0,'<b>Net<br>Earnings</b>')
    
    h_final_revenue=revenue_mn+ h_revenue
    h_final_revenue.insert(0,'<b>Revenue</b>')
    


    #balance sheet analysis 
    #vertical analysis
    t_assets= b_transposed['totalAssets'].to_list()
    cash= b_transposed['cash'].to_list()
    if 'inventory' in b_transposed:
        inventory= b_transposed['inventory'].to_list()
    t_current_assets= b_transposed['totalCurrentAssets'].to_list()
    fixed_assets= b_transposed['propertyPlantEquipment'].to_list()
    other_assets= b_transposed['otherAssets'].to_list()

    #liabilities analysis 
    accounts_payable= b_transposed['accountsPayable'].to_list()
    t_current_liab= b_transposed['totalCurrentLiabilities'].to_list()
    retained_earnings= b_transposed['retainedEarnings'].to_list()
    common_stock= b_transposed['commonStock'].to_list()


    
    # total assets  
    p_t_assets=[]
    p_t_assets_str=[]
    for p in range(0,len(dates)):
        p_t_assets.append(100)
        val="<b>100%</b>"
        p_t_assets_str.append(val)

    #liabilities 
    t_liab= b_transposed['totalLiab'].to_list()
    t_she=b_transposed['totalStockholderEquity'].to_list()
    t_liab_she=[]
    for i in range(0,len(t_liab)):
        add=t_liab[i]+t_she[i]
        t_liab_she.append(add)

    p_t_liab=[]
    p_t_liab_str=[]
    for p in range(0,len(dates)):
        p_t_liab.append(100)
        val="<b>100%</b>"
        p_t_liab_str.append(val)
    
    

    percent_cash=[]
    percent_inventory=[]
    percent_t_current_assets=[]
    percent_fixed_assets=[]
    percent_other_assets=[]

    percent_accounts_payable=[]
    percent_current_liab=[]
    percent_retained_earnings=[]
    percent_common_stock=[]

    graph_cash=[]
    graph_inventory=[]
    graph_t_current_assets=[]
    graph_fixed_assets=[]
    graph_other_assets=[]

    graph_accounts_payable=[]
    graph_current_liab=[]
    graph_retained_earnings=[]
    graph_common_stock=[]

    t_assets_mn=[]
    cash_mn=[]
    inventory_mn=[]
    t_current_assets_mn=[]
    fixed_assets_mn=[]
    other_assets_mn=[]

    t_liab_she_mn=[]
    accounts_payable_mn=[]
    current_liab_mn=[]
    retained_earnings_mn=[]
    common_stock_mn=[]

    for i in range(0,len(dates)):
        
        cash_int=round((cash[i]/t_assets[i])*100,2)
        graph_cash.append(cash_int)
        cash_value=str(cash_int)+"%"
        percent_cash.append(cash_value)  
        cashmn=str(round((cash[i]/1000000000),2))
        cash_mn.append(cashmn)

        if 'inventory' in b_transposed:
            inventory_int=round((inventory[i]/t_assets[i])*100,2)
            graph_inventory.append(inventory_int)
            inventory_value=str(inventory_int)+"%"
            percent_inventory.append(inventory_value)
            inventorymn=str(round((inventory[i]/1000000000),2))
            inventory_mn.append(inventorymn)
        t_current_assets_int=round((t_current_assets[i]/t_assets[i])*100,2)
        fixed_assets_int=round((fixed_assets[i]/t_assets[i])*100,2)
        other_assets_int=round((other_assets[i]/t_assets[i])*100,2)

        accounts_payable_int=round((accounts_payable[i]/t_liab_she[i])*100,2)
        current_liab_int=round((t_current_liab[i]/t_liab_she[i])*100,2)
        retained_earnings_int=round((retained_earnings[i]/t_liab_she[i])*100,2)
        common_stock_int=round((common_stock[i]/t_liab_she[i])*100,2)
        
        
        
        graph_t_current_assets.append(t_current_assets_int)
        graph_fixed_assets.append(fixed_assets_int)
        graph_other_assets.append(other_assets_int)

        graph_accounts_payable.append(accounts_payable_int)
        graph_current_liab.append(current_liab_int)
        graph_retained_earnings.append(retained_earnings_int)
        graph_common_stock.append(common_stock_int)

        
        
        t_current_assets_value="<b>"+str(t_current_assets_int)+"%</b>"
        fixed_assets_value="<b>"+str(fixed_assets_int)+"%</b>"
        other_assets_value=str(other_assets_int)+"%"

        accounts_payable_value=str(accounts_payable_int)+"%"
        current_liab_value=str(current_liab_int)+"%"
        retained_earnings_value=str(retained_earnings_int)+"%"
        common_stock_value=str(common_stock_int)+"%"
        
        
        
        percent_t_current_assets.append(t_current_assets_value)
        percent_fixed_assets.append(fixed_assets_value)
        percent_other_assets.append(other_assets_value)

        percent_accounts_payable.append(accounts_payable_value)
        percent_current_liab.append(current_liab_value)
        percent_retained_earnings.append(retained_earnings_value)
        percent_common_stock.append(common_stock_value)

        t_assetsmn="<b>"+str(round((t_assets[i]/1000000000),2))+"</b>"
        t_assets_mn.append(t_assetsmn)
        

        t_current_assetsmn="<b>"+str(round((t_current_assets[i]/1000000000),2))+"</b>"
        t_current_assets_mn.append(t_current_assetsmn)

        fixed_assetsmn="<b>"+str(round((fixed_assets[i]/1000000000),2))+"</b>"
        fixed_assets_mn.append(fixed_assetsmn)

        other_assetsmn=str(round((other_assets[i]/1000000000),2))
        other_assets_mn.append(other_assetsmn)

        
        t_liab_shemn="<b>"+str(round((t_liab_she[i]/1000000000),2))+"</b>"
        t_liab_she_mn.append(t_liab_shemn)

        accounts_payablemn=str(round((accounts_payable[i]/1000000000),2))
        accounts_payable_mn.append(accounts_payablemn)

        current_liabmn=str(round((t_current_liab[i]/1000000000),2))
        current_liab_mn.append(current_liabmn)

        retained_earningsmn=str(round((retained_earnings[i]/1000000000),2))
        retained_earnings_mn.append(retained_earningsmn)

        common_stockmn=str(round((common_stock[i]/1000000000),2))
        common_stock_mn.append(common_stockmn)
    

    final_t_assets=t_assets_mn+p_t_assets_str
    final_t_assets.insert(0,'<b>Total<br>Assets</b>')
    if cash is not None:
        final_cash=cash_mn+percent_cash
        final_cash.insert(0,'Cash')
    if 'inventory' in b_transposed:
        final_inventory=inventory_mn+percent_inventory
        final_inventory.insert(0,'Inventory')
    final_t_current_assets=t_current_assets_mn+percent_t_current_assets
    final_t_current_assets.insert(0,'<b>Total<br>Current<br>Assets</b>')
    final_fixed_assets=fixed_assets_mn+percent_fixed_assets
    final_fixed_assets.insert(0,'<b>Fixed<br>Assets</b>')
    final_other_assets=other_assets_mn+percent_other_assets
    final_other_assets.insert(0,'Other<br>Assets')

    final_t_liab_she=t_liab_she_mn+ p_t_liab_str
    final_t_liab_she.insert(0,"<b>Total<br>Liabilities<br>and<br>Stock<br>holder's<br>Equity</b>")
    final_accounts_payable=accounts_payable_mn+ percent_accounts_payable
    final_accounts_payable.insert(0,"Accounts<br>Payable")
    final_current_liab=current_liab_mn+ percent_current_liab
    final_current_liab.insert(0,'Current<br>Liabilities')
    final_retained_earnings= retained_earnings_mn+ percent_retained_earnings
    final_retained_earnings.insert(0,"Retained<br>Earnings")
    final_common_stock= common_stock_mn+ percent_common_stock
    final_common_stock.insert(0,'Common<br>Stock')

    #BALANCE SHEET- HORIZONTAL ANALYSIS 
    h_t_assets=[]
    h_cash=[]
    h_inventory=[]
    h_t_current_assets=[]
    h_fixed_assets=[]
    h_other_assets=[]
    

    h_t_liab_she=[]
    h_accounts_payable=[]
    h_current_liab=[]
    h_retained_earnings=[]
    h_common_stock=[]
    

    graph_h_t_assets=[]
    graph_h_cash=[]
    graph_h_inventory=[]
    graph_h_t_current_assets=[]
    graph_h_fixed_assets=[]
    graph_h_other_assets=[]

    graph_h_t_liab_she=[]
    graph_h_accounts_payable=[]
    graph_h_current_liab=[]
    graph_h_retained_earnings=[]
    graph_h_common_stock=[]
    

    for i in range(0,len(dates)):
       
        if i<len(dates)-1:
        
            yoy_assets_int=round(((t_assets[i]/t_assets[i+1])-1)*100,2)
            yoy_cash_int=round(((cash[i]/cash[i+1])-1)*100,2)
            if 'inventory' in b_transposed:
                yoy_inventory_int=round(((inventory[i]/inventory[i+1])-1)*100,2)
                yoy_inventory= str(yoy_inventory_int)+"%"
                h_inventory.append(yoy_inventory)
                graph_h_inventory.append(yoy_inventory_int)
            yoy_current_int=round(((t_current_assets[i]/t_current_assets[i+1])-1)*100,2)
            yoy_fixed_int=round(((fixed_assets[i]/fixed_assets[i+1])-1)*100,2)
            yoy_other_int=round(((other_assets[i]/other_assets[i+1])-1)*100,2)

            yoy_liab_int=round(((t_liab_she[i]/t_liab_she[i+1])-1)*100,2)
            yoy_accounts_int=round(((accounts_payable[i]/accounts_payable[i+1])-1)*100,2)
            yoy_current_liab_int=round(((t_current_liab[i]/t_current_liab[i+1])-1)*100,2)
            yoy_retained_int=round(((retained_earnings[i]/retained_earnings[i+1])-1)*100,2)
            yoy_common_int=round(((common_stock[i]/common_stock[i+1])-1)*100,2)


            yoy_assets= "<b>"+str(yoy_assets_int)+"%</b>"
            yoy_cash= str(yoy_cash_int)+"%"
            
            yoy_current= "<b>"+str(yoy_current_int)+"%</b>"
            yoy_fixed= "<b>"+str(yoy_fixed_int)+"%</b>"
            yoy_other= str(yoy_other_int)+"%"

            yoy_liab= "<b>"+str(yoy_liab_int)+"%</b>"
            yoy_accounts= str(yoy_accounts_int)+"%"
            yoy_current_liab= str(yoy_current_liab_int)+"%"
            yoy_retained= str(yoy_retained_int)+"%"
            yoy_common= str(yoy_common_int)+"%"


            h_t_assets.append(yoy_assets)
            h_cash.append(yoy_cash)
            
            h_t_current_assets.append(yoy_current)
            h_fixed_assets.append(yoy_fixed)
            h_other_assets.append(yoy_other)

            h_t_liab_she.append(yoy_liab)
            h_accounts_payable.append(yoy_accounts)
            h_current_liab.append(yoy_current_liab)
            h_retained_earnings.append(yoy_retained)
            h_common_stock.append(yoy_common)
            
            graph_h_t_assets.append(yoy_assets_int)
            graph_h_cash.append(yoy_cash_int)
            
            graph_h_t_current_assets.append(yoy_current_int)
            graph_h_fixed_assets.append(yoy_fixed_int)
            graph_h_other_assets.append(yoy_other_int)

            graph_h_t_liab_she.append(yoy_liab_int)
            graph_h_accounts_payable.append(yoy_accounts_int)
            graph_h_current_liab.append(yoy_current_liab_int)
            graph_h_retained_earnings.append(yoy_retained_int)
            graph_h_common_stock.append(yoy_common_int)
            

        else:
            break  


    h_t_assets.append("NA")
    h_cash.append("NA")
    if 'inventory' in b_transposed:
        h_inventory.append("NA")
        h_final_inventory= inventory_mn+ h_inventory
        h_final_inventory.insert(0,'Inventory')
    h_t_current_assets.append("NA")
    h_fixed_assets.append("NA")
    h_other_assets.append("NA")

    h_t_liab_she.append("NA")
    h_accounts_payable.append("NA")
    h_current_liab.append("NA")
    h_retained_earnings.append("NA")
    h_common_stock.append("NA")


    
    h_final_t_assets=t_assets_mn+ h_t_assets
    h_final_t_assets.insert(0,'<b>Total<br>Assets</b>')
    h_final_cash= cash_mn+ h_cash
    h_final_cash.insert(0,'Cash')
    
    h_final_current_assets= t_current_assets_mn+ h_t_current_assets
    h_final_current_assets.insert(0,'<b>Total<br>Current<br>Assets</b>')
    h_final_fixed_assets= fixed_assets_mn+ h_fixed_assets
    h_final_fixed_assets.insert(0,'<b>Fixed<br>Assets</b>')
    h_final_other_assets= other_assets_mn+ h_other_assets
    h_final_other_assets.insert(0,'Other<br>Assets')

    h_final_t_liab_she=t_liab_she_mn+ h_t_liab_she
    h_final_t_liab_she.insert(0,"<b>Total<br>Liabilities<br>and<br>Stock<br>holder's<br>Equity</b>")
    h_final_accounts_payable= accounts_payable_mn+ h_accounts_payable
    h_final_accounts_payable.insert(0,'Accounts<br>Payable')
    h_final_current_liab= current_liab_mn+ h_current_liab
    h_final_current_liab.insert(0,'Current<br>Liabilities')
    h_final_retained_earnings= retained_earnings_mn+ h_retained_earnings
    h_final_retained_earnings.insert(0,'Retained<br>Earnings')
    h_final_common_stock= common_stock_mn+ h_common_stock
    h_final_common_stock.insert(0,'Common<br>Stock')
    
    #liquidity ratios
    #CURRENT RATIO=ca/cl
    ca=b_transposed['totalCurrentAssets']
    cl=b_transposed['totalCurrentLiabilities']
    current_ratio=[]

    #QUICK RATIO=cash+short term inv + accounts receivable / current liab
    if 'shortTermInvestments' in b_transposed:
        short_term_inv=b_transposed['shortTermInvestments'].to_list()
    if 'netReceivables' in b_transposed:
        accounts_rec=b_transposed['netReceivables'].to_list()
    quick_ratio=[]

    #profitability ratios 
    #PERCENT GROSS PROFIT 
    cor=i_transposed['costOfRevenue'].to_list()
    gross_profit_ratio=[]

    #% PROFIT MARGIN ON SALES
    profit_margin=[]
    ebt=i_transposed['incomeBeforeTax'].to_list()
    #RETURN ON ASSETS 
    roa=[]
    #RETURN ON EQUITY
    roe=[]

    #activity ratio 
    #ASSET TURNOVER= revenue/net assets
    asset_turnover=[]
    #inventory turnover= cogs/ inventory
    inventory_turnover=[]
    #sales to asset= revenue/ total assets
    sales_asset=[]

    #LEVERAGE RATIOS
    #debt/total assets = total liab/total assets*100- indicates the % of assets financed by debts 
    debt_tassets=[]
    #debt/equity=total liab/ total shareholders equity
    debt_equity=[]
    #equity ratio= total shareholder's equity/ total assets 
    equity_ratio=[]
    #equity multiplier= total assets/ total shareholder's equity
    equity_multiplier=[]
    for i in range(0, len(dates)):
        
            
        #liquidity ratios
        ratio=round(ca[i]/cl[i],2)
        current_ratio.append(ratio)
        if 'shortTermInvestments' in b_transposed:
            if 'netReceivables' in b_transposed:
                ratio1=round((cash[i]+short_term_inv[i]+accounts_rec[i])/cl[i],2)
                quick_ratio.append(ratio1)

        #profitability ratios
        p_gp=round((gross_profit[i]/revenue[i])*100,2)
        gross_profit_ratio.append(p_gp)
        p_margin=round((ebt[i]/revenue[i])*100,2)
        profit_margin.append(p_margin)
        return_on_assets=round((ebt[i]/t_assets[i])*100,2)
        roa.append(return_on_assets)
        return_on_equity=round((ebt[i]/t_she[i])*100,2)
        roe.append(return_on_equity)

        #activity ratios
        r=round(revenue[i]/t_assets[i],2)
        asset_turnover.append(r)
        if 'inventory' and 'costOfRevenue' in b_transposed:
            r2=round(cogs[i]/inventory[i],2)
            inventory_turnover.append(r2)
        r3=round(revenue[i]/t_assets[i],2)
        sales_asset.append(r3)

        #leverage ratios
        da=round(t_liab[i]/t_assets[i],2)
        debt_tassets.append(da)
        de=round(t_liab[i]/t_she[i],2)
        debt_equity.append(de)
        eq_r=round(t_she[i]/t_assets[i],2) 
        equity_ratio.append(eq_r)
        eq_mul=round(t_assets[i]/t_she[i],2)
        equity_multiplier.append(eq_mul)

    current_ratio.insert(0,'Current Ratio')
    

    gross_profit_ratio.insert(0,'Percent Gross Profit')
    profit_margin.insert(0,'Percent Profit Margin')
    roa.insert(0,'Return on Assets')
    roe.insert(0,'Return on Equity')
       
    asset_turnover.insert(0,'Asset Turnover')
    
    sales_asset.insert(0,'Sales to Assets')

    debt_tassets.insert(0,'Debt to Total Assets')
    equity_ratio.insert(0,'Equity Ratio')
    debt_equity.insert(0,'Debt to Equity')
    equity_multiplier.insert(0,'Equity Multiplier')

    current_revenue=str(round(revenue[0]/1000000000,2))+" USD Billion"

    scope = PlotlyScope()   
    # Build story.
    story = []

    ptext='<font size=24 color=darkblue>'+data['name']+' Financial Statement Analysis </font>'
    story.append(Paragraph(ptext,h1))
    story.append(Spacer(1,20))
    if len(about)!=0:
        ptext=about
        story.append(Paragraph(ptext,h1_1))
        story.append(Spacer(1,20))
    ptext='<b>Website</b>: '+'<font color=blue><u>'+comp_link[0]+'</u></font>'
    story.append(Paragraph(ptext,h1_1))
    story.append(Spacer(1,20))
    ptext='<b>Revenue '+str(dates[0][0:4])+"</b>: "+current_revenue
    story.append(Paragraph(ptext,h1_1))
    story.append(PageBreak())

    #CONTENTS
    ptext='<font size=20> Contents </font>'
    story.append(Paragraph(ptext,h1))
    story.append(Spacer(1,10))
    ptext='<b>_________________________________________________________________</b>'
    story.append(Paragraph(ptext,h1))
    story.append(Spacer(1,15))

    destination = sha1("Income Statement- 4 year Trend Analysis".encode('utf-8')).hexdigest()
    ptext='<a href=#{} > <b>Income Statement- 4 year Trend Analysis</b> </a>'.format(destination)
    story.append(Paragraph(ptext,h1_1))
    story.append(Spacer(1,10))

    
    ptext='Vertical Analysis'
    story.append(Paragraph(ptext,h2))
    story.append(Spacer(1,10))
    ptext='Graphical Analysis'
    story.append(Paragraph(ptext,h3))
    story.append(Spacer(1,10))

    destination1 = sha1("Horizontal Analysis of Income Statement ".encode('utf-8')).hexdigest()
    ptext='<a href=#{} > Horizontal Analysis- Year-on-year (Y-O-Y) Trends </a>'.format(destination1)
    story.append(Paragraph(ptext,h2))
    story.append(Spacer(1,10))
    ptext='Graphical Analysis'
    story.append(Paragraph(ptext,h3))
    story.append(Spacer(1,10))

    destination2 = sha1("Balance Sheet- 4 year Trend Analysis".encode('utf-8')).hexdigest()
    ptext='<a href=#{} > <b>Balance Sheet- 4 year Trend Analysis</b></a>'.format(destination2)
    story.append(Paragraph(ptext,h1_1))
    story.append(Spacer(1,10))
    ptext='Vertical Analysis '.format(destination2)
    story.append(Paragraph(ptext,h2))
    story.append(Spacer(1,10))
    ptext='Graphical Analysis of Assets and Liabilities'
    story.append(Paragraph(ptext,h3))
    story.append(Spacer(1,10))

    destination3 = sha1("Horizontal Analysis of Balance Sheet".encode('utf-8')).hexdigest()
    ptext='<a href=#{} > Horizontal Analysis- Year-on-year (Y-O-Y) Trends</a>'.format(destination3)
    story.append(Paragraph(ptext,h2))
    story.append(Spacer(1,10))
    ptext='Graphical Analysis of Assets and Liabilities'
    story.append(Paragraph(ptext,h3))
    story.append(Spacer(1,10))


    destination4 = sha1("Ratio Analysis Four-Year Comparison".encode('utf-8')).hexdigest()
    ptext='<a href=#{} > <b>Ratio Analysis Four-Year Comparison</b></a>'.format(destination4)
    story.append(Paragraph(ptext,h1_1))
    story.append(Spacer(1,10))

    destination5 = sha1("Important Terms and Formulae".encode('utf-8')).hexdigest()
    ptext='<a href=#{} > <b>Important Terms and Formulae</b></a>'.format(destination5)
    story.append(Paragraph(ptext,h1_1))
    story.append(Spacer(1,10))

    if len(articles)!=0:
       destination6 = sha1("Recent Company News".encode('utf-8')).hexdigest()
       ptext='<a href=#{} > <b>Recent Company News</b></a>'.format(destination6)
       story.append(Paragraph(ptext,h1_1))
        
    story.append(PageBreak())

    #Income Statement analysis  
    b=Paragraph('Income Statement- 4 year Trend Analysis' +'<a name="%s"/>' % destination,center)
    b._bookmarkName = destination
    story.append(b)
    story.append(Spacer(1,15))
    ptext='<font size=14> Vertical Analysis of Income Statement</font>'  
    story.append(Paragraph(ptext,h1))
    story.append(Spacer(1,15))
    
    dates1=[]
    for i in dates:
        d=int(i[0:4])
        dates1.append(d)

    dates1.insert(0,"In<br>Billion")
    final_dates=dates1+dates1
    final_dates.pop(5)
    if ebt or sga or net_earnings or gross_profit is not None:
        data=[final_dates,final_revenue,final_cogs,final_gross_profit,final_sga,final_ebt,final_net_earnings]
    

    fig_table1 = ff.create_table(data,height_constant=100)
    for i in range(len(fig_table1.layout.annotations)):
        fig_table1.layout.annotations[i].font.size = 14

    fig_table1.update_layout(template='simple_white')
    pio.write_image(fig_table1,"tmp/vertical.png")

    table_path ="tmp/vertical.png"
    table1_11=open(table_path,"rb")

    story.append(Image(table1_11,480,500))  

    story.append(PageBreak())

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=graph_dates, y=graph_cogs,dx=1,marker=dict(color='#0099ff'),
    mode='lines+markers',name = 'Cost of Goods Sold',showlegend=True))
    fig.add_trace(go.Scatter(x=graph_dates, y=graph_gross_profit,dx=1, marker=dict(color='#000080'),
    mode='lines+markers',name = 'Gross Profit',showlegend=True))
    if 'sellingGeneralAdministrative' in i_transposed:
        fig.add_trace(go.Scatter(x=graph_dates, y=graph_sga,dx=1, marker=dict(color='#A9A9A9'),
        mode='lines+markers',name = 'Selling General and Administrative',showlegend=True))
    if ebt is not None:
        fig.add_trace(go.Scatter(x=graph_dates, y=graph_ebt,dx=1, marker=dict(color='#800000'),
        mode='lines+markers',name = 'Earning before Tax',showlegend=True))
    if net_earnings is not None:
        fig.add_trace(go.Scatter(x=graph_dates, y=graph_net_earnings,dx=1, marker=dict(color='#808000'),
        mode='lines+markers',name = 'Net Earning',showlegend=True))

    fig.update_layout(title="<b> Graphical Analysis </b>",template='simple_white',xaxis_title="Year", yaxis_title="Percentage %")
    pio.write_image(fig,"tmp/vertical-graph.png")
    
    graph1_path='tmp/vertical-graph.png'
    graph1=open(graph1_path,"rb")
    story.append(Image(graph1,500,350))
    story.append(PageBreak())

    b1=Paragraph('<font size=14> Horizontal Analysis of Income Statement</font>'+'<a name="%s"/>' % destination1, h1)
    b1._bookmarkName = destination1  
    story.append(b1)
    story.append(Spacer(1,15))

    if ebt or cogs or gross_profit or sga or net_earnings is not None:
        data1=[final_dates,h_final_revenue,h_final_cogs,h_final_gross_profit,h_final_sga,h_final_ebt,h_final_net_earnings]
        
    fig_table11 = ff.create_table(data1,height_constant=100)
    for i in range(len(fig_table11.layout.annotations)):
        fig_table11.layout.annotations[i].font.size = 14

    fig_table11.update_layout(template='simple_white')
    pio.write_image(fig_table11,"tmp/horizontal.png")
    
    table_path1 = "tmp/horizontal.png"
    table1=open(table_path1,"rb")

    story.append(Image(table1,480,500))  

    story.append(PageBreak())

    fig_h = go.Figure()
    fig_h.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_revenue,dx=1, marker=dict(color='#228B22'),
    mode='lines+markers',name = 'Revenue',showlegend=True))

    if 'costOfRevenue' in i_transposed:
        fig_h.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_cogs,dx=1, marker=dict(color='#0099ff'),
        mode='lines+markers',name = 'Cost of Goods Sold',showlegend=True))
    if gross_profit is not None:
        fig_h.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_gross_profit,dx=1, marker=dict(color='#000080'),
        mode='lines+markers',name = 'Gross Profit',showlegend=True))
    if 'sellingGeneralAdministrative' in i_transposed:
        fig_h.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_sga,dx=1, marker=dict(color='#A9A9A9'),
        mode='lines+markers',name = 'Selling General and Administrative',showlegend=True,))
    if ebt is not None:
        fig_h.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_ebt,dx=1, marker=dict(color='#800000'),
        mode='lines+markers',name = 'Earning before Tax',showlegend=True))
    if net_earnings is not None:
        fig_h.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_net_earnings,dx=1, marker=dict(color='#808000'),
        mode='lines+markers',name = 'Net Earning',showlegend=True))

    fig_h.update_layout(title="<b>Graphical Analysis</b>",template='simple_white',xaxis_title="Year", yaxis_title="Percentage %")
    pio.write_image(fig_h,"tmp/horizontal-graph.png")
    
    graph_h_path= "tmp/horizontal-graph.png"
    graph_h=open(graph_h_path,"rb")
    story.append(Image(graph_h,500,350))
    story.append(PageBreak())

    #balance sheet analysis 
    b2=Paragraph('Balance Sheet - 4 year Trend Analysis'+'<a name="%s"/>' % destination2,center)
    b2._bookmarkName = destination2
    story.append(b2)
    story.append(Spacer(1,15))
    ptext='<font size=14> Vertical Analysis of Balance Sheet</font>'  
    story.append(Paragraph(ptext,h1))
    story.append(Spacer(1,10))

    if 'inventory' in b_transposed:
        data_assets=[final_dates,final_t_assets,final_cash,final_inventory,final_t_current_assets,final_fixed_assets,final_other_assets,
        final_t_liab_she,final_accounts_payable,final_current_liab,final_retained_earnings,final_common_stock]
    else:
        data_assets=[final_dates,final_t_assets,final_cash,final_t_current_assets,final_fixed_assets,final_other_assets,
        final_t_liab_she,final_accounts_payable,final_current_liab,final_retained_earnings,final_common_stock]
    

    fig_table_assets = ff.create_table(data_assets,height_constant=115)
    for i in range(len(fig_table_assets.layout.annotations)):
        if i>8:
            fig_table_assets.layout.annotations[i].font.size = 15
        elif i<=8:
            fig_table_assets.layout.annotations[i].font.size = 15

    fig_table_assets.update_layout(template='simple_white')
    pio.write_image(fig_table_assets,"tmp/assets.png")
    
    table_path_assets = "tmp/assets.png"
    table_assets=open(table_path_assets,"rb")

    story.append(Image(table_assets,450,540))  

    #story.append(PageBreak())


    fig_assets = go.Figure()
    fig_assets.add_trace(go.Scatter(x=graph_dates, y=graph_cash,dx=1, marker=dict(color='#000080'),
    mode='lines+markers',name = 'Cash',showlegend=True))
    if 'inventory' in b_transposed:
        fig_assets.add_trace(go.Scatter(x=graph_dates, y=graph_inventory,dx=1, marker=dict(color='#A9A9A9'),
        mode='lines+markers',name = 'Inventory',showlegend=True))
    fig_assets.add_trace(go.Scatter(x=graph_dates, y=graph_t_current_assets,dx=1, marker=dict(color='#800000'),
    mode='lines+markers',name = 'Total Current Assets',showlegend=True))
    fig_assets.add_trace(go.Scatter(x=graph_dates, y=graph_fixed_assets,dx=1, marker=dict(color='#808000'),
    mode='lines+markers',name = 'Fixed Assets',showlegend=True))
    fig_assets.add_trace(go.Scatter(x=graph_dates, y=graph_other_assets,dx=1, marker=dict(color='#0099ff'),
    mode='lines+markers',name = 'Other Assets',showlegend=True))

    fig_assets.update_layout(title="<b>Graphical Analysis of Assets</b>",template='simple_white',xaxis_title="Year", yaxis_title="Percentage %")
    pio.write_image(fig_assets,"tmp/assets-graph.png")
    
    graph_path_assets= "tmp/assets-graph.png"
    graph_assets=open(graph_path_assets,"rb")
    story.append(Image(graph_assets,400,300))

    fig_liab=go.Figure()
    fig_liab.add_trace(go.Scatter(x=graph_dates, y=graph_accounts_payable,dx=1, marker=dict(color='#000080'),
    mode='lines+markers',name = 'Accounts Payable',showlegend=True))
    fig_liab.add_trace(go.Scatter(x=graph_dates, y=graph_current_liab,dx=1, marker=dict(color='#A9A9A9'),
    mode='lines+markers',name = 'Total Current Liabilities',showlegend=True))
    fig_liab.add_trace(go.Scatter(x=graph_dates, y=graph_t_current_assets,dx=1, marker=dict(color='#800000'),
    mode='lines+markers',name = 'Retained Earnings',showlegend=True))
    fig_liab.add_trace(go.Scatter(x=graph_dates, y=graph_fixed_assets,dx=1, marker=dict(color='#808000'),
    mode='lines+markers',name = 'Common Stock',showlegend=True))
    fig_liab.update_layout(title="<b>Graphical Analysis of Liabilities</b>",template='simple_white',xaxis_title="Year", yaxis_title="Percentage %")
    
    pio.write_image(fig_liab,"tmp/liab-graph.png")

    graph_path_liab= "tmp/liab-graph.png"
    graph_liab=open(graph_path_liab,"rb")
    story.append(Image(graph_liab,400,300))
    story.append(PageBreak()) 

    b3=Paragraph('<font size=14> Horizontal Analysis of Balance Sheet</font>'+'<a name="%s"/>' % destination3,h1)
    b3._bookmarkName = destination3
    story.append(b3)
    story.append(Spacer(1,15))

    if 'inventory' in b_transposed:
        h_balance=[final_dates,h_final_t_assets,h_final_cash,h_final_inventory,h_final_current_assets,h_final_fixed_assets,
        h_final_other_assets,h_final_t_liab_she,h_final_accounts_payable,h_final_current_liab,h_final_retained_earnings,h_final_common_stock]
    else:
        h_balance=[final_dates,h_final_t_assets,h_final_cash,h_final_current_assets,h_final_fixed_assets,
        h_final_other_assets,h_final_t_liab_she,h_final_accounts_payable,h_final_current_liab,h_final_retained_earnings,h_final_common_stock]
    
    fig_balance_h = ff.create_table(h_balance,height_constant=115)
    for i in range(len(fig_balance_h.layout.annotations)):
        if i>8:
            fig_balance_h.layout.annotations[i].font.size = 15
        elif i<=8:
            fig_balance_h.layout.annotations[i].font.size = 15

    fig_balance_h.update_layout(template='simple_white')

    pio.write_image(fig_balance_h,"tmp/horizontal-balance.png")

    table_path_b_v = 'tmp/horizontal-balance.png'
    table_b_v=open(table_path_b_v,"rb")

    story.append(Image(table_b_v,450,550))  

    story.append(PageBreak())
    #horizontal balance graphs
    #assets
    fig_h_assets = go.Figure()
    fig_h_assets.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_t_assets,dx=1, marker=dict(color='#228B22'),
    mode='lines+markers',name = 'Total Assets',showlegend=True))
    fig_h_assets.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_cash,dx=1, marker=dict(color='#0099ff'),
    mode='lines+markers',name = 'Cash',showlegend=True))
    if 'inventory' in b_transposed:
        fig_h_assets.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_inventory,dx=1, marker=dict(color='#000080'),
        mode='lines+markers',name = 'Inventory',showlegend=True))
    
    fig_h_assets.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_t_current_assets,dx=1, marker=dict(color='#800000'),
    mode='lines+markers',name = 'Total Current Assets',showlegend=True))

    fig_h_assets.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_fixed_assets,dx=1, marker=dict(color='#808000'),
    mode='lines+markers',name = 'Fixed Assets',showlegend=True))
    fig_h_assets.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_other_assets,dx=1, marker=dict(color='#A9A9A9'),
    mode='lines+markers',name = 'Other Assets',showlegend=True,))

    fig_h_assets.update_layout(title="<b>Graphical Analysis of Assets</b>",template='simple_white',xaxis_title="Year", yaxis_title="Percentage %")
    
    pio.write_image(fig_h_assets,"tmp/assets-horizontal-graph.png")
    
    graph_h_assets_path="tmp/assets-horizontal-graph.png"
    graph_h_assets=open(graph_h_assets_path,"rb")
    story.append(Image(graph_h_assets,400,300))

    fig_h_liab=go.Figure()

    fig_h_liab.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_t_liab_she,dx=1, marker=dict(color='#228B22'),
    mode='lines+markers',name = "Total Liabilities and Stockholder's Equity",showlegend=True))

    fig_h_liab.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_accounts_payable,dx=1, marker=dict(color='#0099ff'),
    mode='lines+markers',name = 'Accounts Payable',showlegend=True))

    fig_h_liab.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_current_liab,dx=1, marker=dict(color='#000080'),
    mode='lines+markers',name = 'Total Current Liabilities',showlegend=True))

    fig_h_liab.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_retained_earnings,dx=1, marker=dict(color='#800000'),
    mode='lines+markers',name = 'Retained Earnings',showlegend=True,))

    fig_h_liab.add_trace(go.Scatter(x=graph_dates[:-1], y=graph_h_common_stock,dx=1, marker=dict(color='#A9A9A9'),
    mode='lines+markers',name = 'Common Stock',showlegend=True))
    fig_h_liab.update_layout(title="<b>Graphical Analysis of Liabilities</b>",template='simple_white',xaxis_title="Year", yaxis_title="Percentage %")
    
    pio.write_image(fig_h_liab,"tmp/liab-horizontal-graph.png")

    graph_h_liab_path= "tmp/liab-horizontal-graph.png"
    graph_h_liab_path=open(graph_h_liab_path,"rb")
    story.append(Image(graph_h_liab_path,400,300))

    story.append(PageBreak())

    #LIQUIDITY RATIOS:
    
    b4=Paragraph('Ratio Analysis Four-Year Comparison'+'<a name="%s"/>' % destination4,center)
    b4._bookmarkName = destination4
    story.append(b4)
    story.append(Spacer(1,20))

    dates.insert(0," ")
    ratio_dates=dates
    ratio_dates.pop(1)

    if len(quick_ratio)!=0 and 'inventory' in b_transposed:
        quick_ratio.insert(0,'Quick Ratio')
        
        inventory_turnover.insert(0,'Inventory Turnover')
        ratios=[ratio_dates,['LIQUIDITY RATIOS','','','',''],current_ratio,quick_ratio,['','','','',''],
        ['PROFITABILITY RATIOS','','','',''],gross_profit_ratio,profit_margin,roa,roe,['','','','',''],
        ['ACTIVITY RATIOS','','','',''],asset_turnover,inventory_turnover,sales_asset,['','','','',''],
        ['LEVERAGE RATIOS','','','',''],debt_tassets,equity_ratio,debt_equity,equity_multiplier]
    else:
        ratios=[ratio_dates,['LIQUIDITY RATIOS','','','',''],current_ratio,['','','','',''],
        ['PROFITABILITY RATIOS','','','',''],gross_profit_ratio,profit_margin,roa,roe,['','','','',''],
        ['ACTIVITY RATIOS','','','',''],asset_turnover,sales_asset,['','','','',''],
        ['LEVERAGE RATIOS','','','',''],debt_tassets,equity_ratio,debt_equity,equity_multiplier]

    t=Table(ratios,hAlign=TA_LEFT)
    t.setStyle(TableStyle([('FONTNAME', (0,0), (0,-1), 'Times-Bold')]))
    bold_rows=[1,5,11,16]
    for row_num in range(0,len(bold_rows)):
        each_row=bold_rows[row_num]
        t.setStyle(TableStyle([('FONTNAME', (each_row, each_row), (0, each_row), 'Times-Bold')]))

    story.append(t)
    story.append(PageBreak())

    b5=Paragraph('<font size=14> <b><u>Important Terms and Formulae</u></b> </font>'+'<a name="%s"/>' % destination5,center)
    b5._bookmarkName = destination5
    story.append(b5)
    story.append(Spacer(1,20))

    ptext='<font size=14><b> Horizontal Analysis </b></font>'
    story.append(Paragraph(ptext,center))
    story.append(Spacer(1,20))
    ptext='Horizontal analysis compares account balances and ratios over different time periods. For example, you compare a company’s sales in 2014 to its sales in 2015.'
    story.append(Paragraph(ptext,h1_1))
    story.append(Spacer(1,10))

    ptext='<font size=14><b> Vertical Analysis </b></font>'
    story.append(Paragraph(ptext,center))
    story.append(Spacer(1,20))
    ptext='Vertical analysis restates each amount in the income statement as a percentage of sales. This analysis gives the company a heads up if cost of goods sold or any other expense appears to be too high when compared to sales.'
    story.append(Paragraph(ptext,h1_1))
    story.append(Spacer(1,20))


    ptext='<font size=14> <b>Ratios</b> </font>'
    story.append(Paragraph(ptext,center))
    story.append(Spacer(1,20))
    ptext='Liquidity Ratios'
    story.append(Paragraph(ptext,h1))
    story.append(Spacer(1,10))
    ptext='Current Ratio = Current Assets / Current Liabilties'
    story.append(Paragraph(ptext,h1_1))
    ptext='Quick Ratio = (Cash + Market Securities + Accounts Receivable) / Current Liabilities'
    story.append(Paragraph(ptext,h1_1))
    story.append(Spacer(1,10))

    ptext='Profitability Ratios'
    story.append(Paragraph(ptext,h1))
    story.append(Spacer(1,10))
    ptext='Percent Gross Profit = (Gross Profit / Sales)*100'
    story.append(Paragraph(ptext,h1_1))
    ptext='Percent Profit Margin = (Earnings before Taxes / Sales)* 100'
    story.append(Paragraph(ptext,h1_1))
    ptext='Return on Assets = Earnings before Taxes / Total Assets'
    story.append(Paragraph(ptext,h1_1))
    ptext='Return on Equity = Earnings before Taxes / Total Equity'
    story.append(Paragraph(ptext,h1_1))
    story.append(Spacer(1,10))

    ptext='Activity Ratios'
    story.append(Paragraph(ptext,h1))
    story.append(Spacer(1,10))
    ptext='Asset Turnover Ratio = Sales/Total Assets'
    story.append(Paragraph(ptext,h1_1))
    ptext='Inventory Turnover Ratio = Cost of Sales / Inventory'
    story.append(Paragraph(ptext,h1_1))
    ptext='Sales to Assets = Sales / Total Assets'
    story.append(Paragraph(ptext,h1_1))
    story.append(Spacer(1,10))

    ptext='Leverage Ratios'
    story.append(Paragraph(ptext,h1))
    story.append(Spacer(1,10))
    ptext='Debt to Total Assets = Total Liabilities / Total Assets'
    story.append(Paragraph(ptext,h1_1))
    ptext='Equity Ratio = Total Equity / Total Assets'
    story.append(Paragraph(ptext,h1_1))
    ptext='Debt to Equity = Total Liabilities / Total Equity'
    story.append(Paragraph(ptext,h1_1))
    ptext='Equity Multiplier = Total Assets / Total Equity'
    story.append(Paragraph(ptext,h1_1))
    story.append(Spacer(1,10))

    ptext='References: <font color=blue><u> https://cs.thomsonreuters.com </u></font>'
    story.append(Paragraph(ptext,h2))
    ptext='<font color=blue><u>https://www.dummies.com/business/accounting/horizontal-and-vertical-analysis/ </u></font>'
    story.append(Paragraph(ptext,h2))
    story.append(PageBreak())

    #NEWS SECTION

    if len(articles)!=0:
        b6=Paragraph('<font size=14> <b><u>Recent Company News</u></b> </font>'+'<a name="%s"/>' % destination6,center)
        b6._bookmarkName = destination6
        story.append(b6)
        story.append(Spacer(1,20))
        for n in range(0,len(n_title)):
            
            
                ptext="<b>"+n_title[n]+"</b>"
                story.append(Paragraph(ptext,h1_1))
                story.append(Spacer(1,5))

                ptext="<font size=11 color=grey> Published on: "+n_date[n]+"</font>"
                story.append(Paragraph(ptext,h1_1))

                ptext=str(n_desc[n])
                story.append(Paragraph(ptext,h1_1))
                story.append(Spacer(1,10))
                ptext="<font color=blue><u>"+n_url[n]+"</u></font>"
                story.append(Paragraph(ptext,h1_1))
                story.append(Spacer(1,20))
        story.append(PageBreak())
       
        
    

    ptext='<font size=20> Created by: Nikita Pardeshi </font>'
    story.append(Paragraph(ptext,h1))
    story.append(Spacer(1,20))
    
    ptext = "<font size=20> Phone: +91 8879367746 </font>"
    story.append(Paragraph(ptext, h1))
    story.append(Spacer(1, 10))

    address='EMAIL: <font color= blue > <u> nikitapardesi.work@gmail.com</u> </font>'
    story.append(Paragraph(address, h1))
    story.append(Spacer(1, 10))

    address1 = 'LinkedIn: <u> <font color= blue > https://www.linkedin.com/in/nikita-pardeshi-bb411a168 </font></u>'
    story.append(Paragraph(address1, h1))
    story.append(Spacer(1, 10))

    doc = SimpleDocTemplate("{}.pdf".format(company_name),pagesize=letter)
    doc.build(story)


    #merging coverpage pdf and main content pdf into one 

    report=BytesIO()

    pdf2File = open('company.pdf', 'rb')

    
    pdf2Reader = PyPDF2.PdfFileReader(pdf2File)
    pdf_watermark="tmp/watermark.pdf"
    pdf3File=open(pdf_watermark,'rb')
    watermark = PyPDF2.PdfFileReader(pdf3File)

    pdfWriter = PyPDF2.PdfFileWriter()

    for pageNum in range(pdf2Reader.numPages):
        pageObj = pdf2Reader.getPage(pageNum)
        pdfWriter.addPage(pageObj)
        pageObj.mergePage(watermark.getPage(0))


    pdfWriter.write(report)
    my_pdf=report.getvalue()
    
    pdf2File.close()

    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = 'inline; filename="{}.pdf"'.format(company_name)
    response.write(my_pdf)
    return response

    


import numpy as np
import requests
import urllib3
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import filedialog
import pandas as pd
from pandastable import Table, TableModel
import datetime
from datetime import datetime, timedelta
from tkinter import simpledialog
import time
import os
from tkinter import messagebox
import matplotlib.pyplot as plt
from scipy.stats import norm

global cwd

cwd = os.getcwd()

class port_manager:
    def __init__(self):

        app = tk.Tk()

        def page1(app):
            #Opening Page
            page = tk.Frame(app)
            page.grid()
            app.geometry("920x600")
            Welcome_label = Label(app, text = 'Portfolio Manager', width = 45, height = 4, font = "Magneto 20", borderwidth=2, relief='solid', bg='gold2', fg='gray30').grid(row = 0, column = 0, columnspan=1, padx = 3, pady=3)
            tk.Button(app, text = 'Custom Portfolio Tools', command = port_upload, font='Helvitica', width=18, bg='gold2', relief='solid').grid(row = 1, column = 0, columnspan=1, padx = 3, pady = 3)
            tk.Button(app, text = 'Currencies', command=lambda:[which_button(button_text = "Currencies"), self.markets(), view_markets()], font='Helvitica', width=18, bg='gold2', relief='solid').grid(row = 2, column = 0, columnspan=1, padx = 3, pady = 3)
            tk.Button(app, text = 'Cryptocurrencies', command=lambda:[which_button(button_text = "Cryptocurrencies"), self.markets(), view_markets()], font='Helvitica', width=18, bg='gold2', relief='solid').grid(row = 3, column = 0, columnspan=1, padx = 3, pady = 3)
            tk.Button(app, text = 'Global Indices', command=lambda:[which_button(button_text = "Global Indices"), self.markets(), view_markets()], font='Helvitica', width=18, bg='gold2', relief='solid').grid(row = 4, column = 0, columnspan=1, padx = 3, pady = 3)
            tk.Button(app, text = 'Bond Markets', command=lambda:[which_button(button_text = "Bond Markets"), self.markets(), view_markets()], font='Helvitica', width=18, bg='gold2', relief='solid').grid(row = 5, column = 0, columnspan=1, padx = 3, pady = 3)
            tk.Button(app, text = 'User Guide', command = user_guide, font='Helvitica', width=18, bg='gold2', relief='solid').grid(row = 6, column = 0, columnspan=1, padx = 3, pady = 3)
            Label(app, text = "Disclaimer: The information furnished on this application is for informational purposes only. The information provided does not constitute investment ", bg = 'gray30', fg = 'white').grid(row = 8, column = 0)
            Label(app, text = "advice and should not be considered as advice to buy or sell securities. The information should not be relied upon by any person to make an investment decision.", bg = 'gray30', fg = 'white').grid(row = 9, column = 0)
      
        def page2(app):
            #Custom Portfolio Page - Enter a list of tickers
            global height,width,entries
            page = tk.Frame(app)
            page.grid()
            height = 15
            width = 3
            entries = []
            for i in range(height): #Rows
                entries.append([])
                for j in range(width): #Columns
                    entries[i].append(Entry(app, text="", width=41, bg='gray70', relief='solid'))
                    entries[i][j].grid(row=i+3, column=j)
            tk.Button(app, text = '⬅', command = home, font='Helvitica', width = 8, bg='gold2', relief='solid').grid(row = 0, column = 0,pady=2,padx=5, sticky='w')
            tk.Label(app, text = 'Construct a Portfolio or Upload an existing one', bg='gray21', fg='gray70', font= 'Vijaya 13', relief='solid', width=50).grid(row = 0, column = 0, columnspan=3, padx = 120)
            tk.Label(app, text= 'Symbol', bg='gray21', fg='gray70', font=13, relief='solid', width=20).grid(row = 2, column =0)
            tk.Label(app, text= 'Name', bg='gray21', fg='gray70', font=13, relief='solid', width=20).grid(row = 2, column =1)
            tk.Label(app, text= 'Sector', bg='gray21', fg='gray70', font=13, relief='solid', width=20).grid(row = 2, column =2)
            tk.Button(app, text = 'Next', command=lambda:[get_table(), port_tools(), which_button(button_text="Use Table")], font='Helvitica', width = 10, bg='gold2', relief='solid').grid(row = 0, column = 3,pady=2,padx=5, sticky='w')
            tk.Button(app, text = 'Upload csv', command=lambda:[upload_csv(), port_tools(), which_button(button_text="Upload Portfolio")], font='Helvitica', width = 10, bg='gold2', relief='solid').grid(row = 1, column = 3,pady=2,padx=5, sticky='w')
            tk.Button(app, text = 'Help', command=lambda:[help()], font='Helvitica', width = 10, bg='gold2', relief='solid').grid(row =2, column = 3,pady=2,padx=5, sticky='w')

        def page3(app):
            #Choice of functions for your custom portfolio
            page = tk.Frame(app)
            page.grid()
            tk.Button(app, text = '⬅', command = home, font='Helvitica', width = 8, bg='gold2', relief='solid').grid(row = 0, column = 0,pady=2,padx=5, sticky='w')
            tk.Label(app, text = 'Function', bg='gray21', fg='gray70', height = 2, relief = 'solid', font=14).grid(row = 1, column = 0,pady=2,padx=5, sticky='news')
            tk.Label(app, text = 'Description', bg='gray21', fg='gray70', height = 2, relief = 'solid', font=14).grid(row = 1, column = 1,pady=2, sticky='news')
            tk.Button(app, text = 'Raw Data', command=lambda:[self.request(), view_req(), self.plot()], font='Helvitica', width = 20, height = 2, bg='gold2', relief='solid').grid(row = 2, column = 0,pady=2,padx=5, sticky='w')
            tk.Button(app, text = 'Correlation Matrix', command=lambda:[self.request(), self.correl(), view_correl()], font='Helvitica', width = 20, height = 2, bg='gold2', relief='solid').grid(row = 3, column = 0,pady=2,padx=5, sticky='w')
            tk.Button(app, text = 'Daily Returns', command=lambda:[self.request(), self.daily_returns(), view_returns()], font='Helvitica', width = 20, height = 2, bg='gold2', relief='solid').grid(row = 4, column = 0,pady=2,padx=5, sticky='w')
            tk.Button(app, text = 'Performance', command=lambda:[self.request(), self.perf(), view_perf()], font='Helvitica', width = 20, height = 2, bg='gold2', relief='solid').grid(row = 5, column = 0,pady=2,padx=5, sticky='w')
            tk.Button(app, text = 'Plot Returns', command=lambda:[which_button(button_text="Plot Returns"), self.request(), self.daily_returns(), self.plot()], font='Helvitica', width = 20, height = 2, bg='gold2', relief='solid').grid(row = 6, column = 0,pady=2,padx=5, sticky='w')
            tk.Button(app, text = 'Scenarios', command=lambda:[scenario_choices(), which_button(button_text="Scenarios")], font='Helvitica', width = 20, height = 2, bg='gold2', relief='solid').grid(row = 7, column = 0,pady=2,padx=5, sticky='w')
            tk.Button(app, text = 'Statistics', command=lambda:[self.statistics(), view_statistics()], font='Helvitica', width = 20, height = 2, bg='gold2', relief='solid').grid(row = 8, column = 0,pady=2,padx=5, sticky='w')                        
            tk.Label(app, text = 'Price history of all assets listed in the portfolio over a given time period.', bg='gray70', width = 70, height = 3, relief = 'solid').grid(row = 2, column = 1,pady=2, sticky='w')
            tk.Label(app, text = 'Correlation Matrix based on returns over a given time period.', bg='gray70', width = 70, height = 3, relief = 'solid').grid(row = 3, column = 1,pady=2, sticky='w')
            tk.Label(app, text = 'Daily Returns of each asset over each day for a given time period.', bg='gray70', width = 70, height = 3, relief = 'solid').grid(row = 4, column = 1,pady=2, sticky='w')
            tk.Label(app, text = 'Absolute return of assets in the portfolio over a given time period.', bg='gray70', width = 70, height = 3, relief = 'solid').grid(row = 5, column = 1,pady=2, sticky='w')
            tk.Label(app, text = 'Plot the returns of your portfolio to compare assets.', bg='gray70', width = 70, height = 3, relief = 'solid').grid(row = 6, column = 1,pady=2, sticky='w')
            tk.Label(app, text = 'See how your portfolio would perform given some extreme market conditions.', bg='gray70', width = 70, height = 3, relief = 'solid').grid(row = 7, column = 1,pady=2, sticky='w')
            tk.Label(app, text = 'Portfolio Statistics.', bg='gray70', width = 70, height = 3, relief = 'solid').grid(row = 8, column = 1,pady=2, sticky='w')           
                                        
        def page4(app):
            #Scenarios for your portfolio - Make sure for scenarios like Lehman/Tech Bubble that your stocks were public/existed back then
            #For example, applying the Lehaman date range to a new stock like 'Reddit' won't work.
            page = tk.Frame(app)
            page.grid()
            tk.Button(app, text = '⬅', command = home, font='Helvitica', width = 8, bg='gold2', relief='solid').grid(row = 0, column = 0,pady=2,padx=5, sticky='w')
            tk.Label(app, text = 'Function', bg='gray21', fg='gray70', height = 2, relief = 'solid', font=14).grid(row = 1, column = 0,pady=2,padx=5, sticky='news')
            tk.Label(app, text = 'Description', bg='gray21', fg='gray70', height = 2, relief = 'solid', font=14).grid(row = 1, column = 1,pady=2, sticky='news')
            tk.Button(app, text = 'Yen Carry Trade', command=lambda:[which_button(button_text="Yen Carry Trade"),self.request(), self.scenarios(), view_scenarios()], font='Helvitica', width = 20, height = 2, bg='gold2', relief='solid').grid(row = 2, column = 0,pady=2,padx=5, sticky='w')
            tk.Button(app, text = 'FED 50bp Rate Cut', command=lambda:[which_button(button_text="FED 50bp Rate Cut"),self.request(), self.scenarios(), view_scenarios()], font='Helvitica', width = 20, height = 2, bg='gold2', relief='solid').grid(row = 3, column = 0,pady=2,padx=5, sticky='w')
            tk.Button(app, text = 'Lehman Collapses', command=lambda:[which_button(button_text="Lehman Collapses"),self.request(), self.scenarios(), view_scenarios()], font='Helvitica', width = 20, height = 2, bg='gold2', relief='solid').grid(row = 4, column = 0,pady=2,padx=5, sticky='w')
            tk.Button(app, text = 'Dot Com Crash', command=lambda:[which_button(button_text="Dot Com Crash"),self.request(), self.scenarios(), view_scenarios()], font='Helvitica', width = 20, height = 2, bg='gold2', relief='solid').grid(row = 5, column = 0,pady=2,padx=5, sticky='w')
            tk.Label(app, text = 'Aug 2024: Bank of Japan raised interest rates causing a large sell off in stocks to make margin calls.', bg='gray70', width = 75, height = 3, relief = 'solid').grid(row = 2, column = 1,pady=2, sticky='w')
            tk.Label(app, text = 'Sep 2024: Federal Reserve makes a large cut in rates following signs of economic slow down.', bg='gray70', width = 75, height = 3, relief = 'solid').grid(row = 3, column = 1,pady=2, sticky='w')
            tk.Label(app, text = 'Sep 2008: Lehman fails to secure a bailout from the United States government.', bg='gray70', width = 75, height = 3, relief = 'solid').grid(row = 4, column = 1,pady=2, sticky='w')
            tk.Label(app, text = 'Mar 2000: The tech bubble bursts.', bg='gray70', width = 75, height = 3, relief = 'solid').grid(row = 5, column = 1,pady=2, sticky='w')
            Label(app, text = "NOTE: Please ensure selected stocks were public companies during the listed timeframes.", bg = 'gray30', fg = 'white').grid(row = 6, column = 0, columnspan=3)

        def help():
            tk.messagebox.showinfo("Help", "Input Stock Symbol, Name and Sector - For example: AAPL | Apple | Information Technology. \
NOTE: Stock Symbol must be exact, but Name and Sector do not have to match the Symbol - For example: AAPL | APPLE INC | Big Tech. \
Equivalently, you can upload a portfolio using a .csv file - Please ensure the csv has the same headers as the table.")

        def user_guide():
            tk.messagebox.showinfo("User Guide","This app contains two different sections with different functionality. The first section 'Custom Portfolio Tools' allows users to create a hypothetical portfolio to conduct analysis on.\
 This can be done by using the table provided or uploading a csv with the same headers as the table provided.\
 Once a portfolio is created there are multiple tools such as 'Raw Data', 'Correlation Matrix','Daily Returns', 'Performance', 'Plots', 'Scenarios' and 'Statistics.\
 The second section provides an overview of whats happening in the markets, see 'Currencies, ' Cryptocurrencies', 'Global Indices' and 'Bond Yields.'\
")
            
        def get_table():
            #Takes your custom portfolio inputs and stores them in a dataframe and as a .csv in your working directory.
            global comb_df, df_constituents_base
            sym = [entries[i][0].get() for i in range(height)]
            name = [entries[i][1].get() for i in range(height)]
            sect = [entries[i][2].get() for i in range(height)]
            comb = {'Symbol':sym, 'Name':name, 'Sector':sect}
            comb_df = pd.DataFrame(comb)
            comb_df['Symbol']=comb_df['Symbol'].astype(str)
            comb_df['Name']=comb_df['Name'].astype(str)
            comb_df['Sector']=comb_df['Sector'].astype(str)
            comb_df['Symbol'].replace('', np.nan, inplace=True)
            comb_df.dropna(subset=['Symbol'], inplace=True)
            now = format(datetime.now().strftime('%d-%m-%Y %H-%M-%S'))
            comb_df.to_csv(os.path.join(cwd, "Portfolio " + now + ".csv"), lineterminator = '\n', encoding='utf-8', index=False)
            messagebox.showinfo("Portfolio", "Portfolio Holdings file has been saved as: " + os.path.join(cwd, "Portfolio " + now + ".csv"))
            df_constituents_base = comb_df

        def which_button(button_text):
            #keeps track of buttons pressed for certain things like scenarios - this is important because the scenarios functionality just runs the request function below for fixed dates.
            global msg
            msg = f"Button Clicked: {button_text}"
            print(msg)

        def view_req():
            #pop up dataframe
            newWindow = Toplevel(app)
            table = Table(newWindow, dataframe=data_all, showtoolbar=True, showstatusbar=True, height=400, width=800)
            table.show()
            table.redraw()
        
        def view_correl():
            #pop up dataframe
            newWindow = Toplevel(app)
            table = Table(newWindow, dataframe=correlC, showtoolbar=True, showstatusbar=True, height=400, width=800)
            table.show()
            table.redraw()

        def view_returns():
            #pop up dataframe
            newWindow = Toplevel(app)
            table = Table(newWindow, dataframe=data_ror, showtoolbar=True, showstatusbar=True, height=400, width=800)
            table.show()
            table.redraw()

        def view_perf():
            #pop up dataframe
            newWindow = Toplevel(app)
            table = Table(newWindow, dataframe=data_perf_join, showtoolbar=True, showstatusbar=True, height=400, width=800)
            table.show()
            table.redraw()

        def view_scenarios():
            #pop up dataframe
            newWindow = Toplevel(app)
            table = Table(newWindow, dataframe=data_scen_join, showtoolbar=True, showstatusbar=True, height=400, width=800)
            table.show()
            table.redraw()

        def view_statistics():
            #pop up dataframe
            newWindow = Toplevel(app)
            table = Table(newWindow, dataframe=statistics, showtoolbar=True, showstatusbar=True, height=400, width=800)
            table.show()
            table.redraw()   

        def view_markets():
            #pop up dataframe
            newWindow = Toplevel(app)
            table = Table(newWindow, dataframe=markets, showtoolbar=True, showstatusbar=True, height=400, width=800)
            table.show()
            table.redraw()          

        def upload_csv():
            #this is for uploading a custom portfolio via a .csv
            global file, df_constituents_base
            file = pd.read_csv(filedialog.askopenfilename(initialdir='/Desktop',title='Select Portfolio csv',filetypes = (("CSV Files","*.csv"),)))
            if len(file['Symbol']) <= 516:
                df_constituents_base = file
            else:
                messagebox.showerror("Error", "Can not select more than 15 assets")

        def home():
            #next page function
            for widget in app.winfo_children():
                widget.destroy()
            page1(app)

        def port_upload():
            #next page function
            for widget in app.winfo_children():
                widget.destroy()
            page2(app)

        def port_tools():
            #next page function
            for widget in app.winfo_children():
                widget.destroy()
            page3(app)

        def scenario_choices():
            #next page function
            for widget in app.winfo_children():
                widget.destroy()
            page4(app)

        #configuring mainloop
        app.title("Portfolio Manager v1.0")
        app.configure(bg='gray30')
        page1(app)
        app.mainloop()

    def request(self):
        #price history function - for scenarios function it takes the fixed dates, for custom dates there are input boxes that take dates in the format dd/mm/yyyy
        global data, data_app, dates, data_all, start_date, end_date, lst_symbol_list

        apiBase = 'https://query2.finance.yahoo.com'
        headers = { 
        "User-Agent": 
        "Mozilla/5.0 (Windows NT 6.1; Win64; x64)"
        }
        
        def getCredentials(cookieUrl='https://fc.yahoo.com', crumbUrl=apiBase+'/v1/test/getcrumb'):
            cookie = requests.get(cookieUrl).cookies
            crumb = requests.get(url=crumbUrl, cookies=cookie, headers=headers).text
            return {'cookie': cookie, 'crumb': crumb}
        def quote(symbols, credentials):
            url = apiBase + f'/v8/finance/chart/{symbols}?period1={difference_1}&period2={difference_2}&interval=1d&events=history&includeAdjustedClose=true'
            print(url)
            params = {'symbols': ','.join(symbols), 'crumb': credentials['crumb']}
            print(params)
            response = requests.get(url, params=params, cookies=credentials['cookie'], headers=headers)
            quotes = response.json()
            return quotes
        
        if msg == "Button Clicked: Yen Carry Trade":
            start_date = '02/08/2024'
            end_date = '06/08/2024'
        elif msg == "Button Clicked: FED 50bp Rate Cut":
            start_date = '17/09/2024'
            end_date = '19/09/2024'
        elif msg == "Button Clicked: Lehman Collapses":
            start_date = '01/09/2008'
            end_date = '15/09/2008' 
        elif msg == "Button Clicked: Dot Com Crash":
            start_date = '10/03/2000'
            end_date = '23/04/2000' 
        else:
            start_date = simpledialog.askstring("Dates", "Enter Start Date (dd/mm/yyyy):")
            end_date = simpledialog.askstring("Dates", "Enter End Date (dd/mm/yyyy):")
        lst_symbol_list = sorted(sum(df_constituents_base.drop(['Name', 'Sector'], axis = 1).to_numpy().tolist(), [])) #sum [] needed to reduce nested list
        secs = 86400
        date_format = '%d/%m/%Y'
        base_date_format = datetime.strptime('01/01/1970', date_format)
        start_date_format = datetime.strptime(start_date, date_format)
        end_date_format = datetime.strptime(end_date, date_format)
        difference_1 = str((start_date_format - base_date_format).days*secs) #days between start date and base date 01/01/1970 converted to seconds
        difference_2 = str((end_date_format - base_date_format).days*secs+secs)
        lst_url = [f'https://query2.finance.yahoo.com/v8/finance/chart/{symbol}?period1={difference_1}&period2={difference_2}&interval=1d&events=history&includeAdjustedClose=true' 
                for symbol in lst_symbol_list]
        
        credentials = getCredentials()
        data = [quote(lst_symbol_list[i], credentials) for i in range(0, len(lst_symbol_list))]        
        dataC = data.copy()
        dataCC = dataC.copy()
        dates = []
        for i in range(0, len(lst_symbol_list)):
            data[i] = pd.DataFrame(data[i]['chart']['result'][0]['indicators']['quote'][0])
            dates.append(dataC[i]['chart']['result'][0]['timestamp'])
            data[i]['Date'] = dates[i]
            for j in range(0, len(data[i]['Date'])):
                data[i]['Date'][j] = time.strftime('%Y-%m-%d', time.gmtime(data[i]['Date'][j]))

        tickers = []
        for i in range(0, len(lst_symbol_list)):
            tickers.append(list(pd.Series(dataCC[i]['chart']['result'][0]['meta']['symbol']))*len(dataCC[i]['chart']['result'][0]['timestamp']))
            data[i]['Symbol'] = tickers[i]

        data_app = pd.concat(data, ignore_index=False)
        data_all = pd.merge(data_app, df_constituents_base, on="Symbol", how="left")
        data_all = data_all[['Date','Symbol','Name','Sector','open','high','low','close','volume']].rename(columns = {'Symbol':'Ticker','Name':'Asset_Name','open':'Open_Price','high':'High_Price','low':'Low_Price','close':'Close_Price','volume':'Volume'}) 
        print(data_all)
        now = format(datetime.now().strftime('%d-%m-%Y %H-%M-%S'))
        data_all.to_excel(os.path.join(cwd, "Portfolio_Data " + now +".xlsx"), index=False)
        messagebox.showinfo("Raw Data", "Portfolio Raw Data Has Been Saved: " + os.path.join(cwd, "Portfolio_Data " + now +".xlsx"))

    def correl(self):
        #correlation matrix
        global correlC
        data_corr = data_all.drop(columns = ['Asset_Name','Sector','Open_Price','High_Price','Low_Price','Volume'])
        ror = [0]+[round((data_corr.iloc[j+1,2]-data_corr.iloc[j,2])*100/data_corr.iloc[j,2],2)\
                for j in range(0, len(data_corr)-1)]
        data_corr['ROR'] = ror
        earliest_date = data_corr.iloc[0,0]
        data_corr = data_corr.query('Date != @earliest_date')
        data_corr = data_corr.reset_index(drop=True)
        data_corr = data_corr.drop(columns=['Close_Price'])
        data_corr = data_corr.pivot(index='Date', columns='Ticker', values='ROR')
        correl = data_corr.corr().round(2)
        correlC = correl.copy()
        correl = correl.style.background_gradient(cmap='viridis')
        now = format(datetime.now().strftime('%d-%m-%Y %H-%M-%S'))
        correl.to_excel(os.path.join(cwd, "Portfolio_Correlation " + now + ".xlsx"))
        messagebox.showinfo("Correlation Matrix", "Correlation Matrix file has been saved as: " + os.path.join(cwd, "Portfolio_Correlation " + now + ".xlsx"))

    def daily_returns(self):
        #returns day by day
        global data_ror
        data_ror = data_all.drop(columns = ['Asset_Name','Sector','Open_Price','High_Price','Low_Price','Volume'])
        ror = [0]+[round((data_ror.iloc[j+1,2]-data_ror.iloc[j,2])*100/data_ror.iloc[j,2],2)\
                for j in range(0, len(data_ror)-1)]
        data_ror['Return %'] = ror
        earliest_date = data_ror.iloc[0,0]
        data_ror = data_ror.query('Date != @earliest_date')
        data_ror = data_ror.reset_index(drop=True)
        now = format(datetime.now().strftime('%d-%m-%Y %H-%M-%S'))
        data_ror.to_excel(os.path.join(cwd, "Portfolio_Returns " + now + ".xlsx"))
        messagebox.showinfo("Portfolio Returns", "Portfolio Returns file has been saved as: " + os.path.join(cwd, "Portfolio_Returns " + now + ".xlsx"))

    def perf(self):
        #performance for a specified period
        global data_perf_join
        start_date = time.strftime('%Y-%m-%d', time.gmtime(dates[0][0]))
        end_date = time.strftime('%Y-%m-%d', time.gmtime(dates[0][-1]))
        data_perf_0 = data_all.drop(columns = ['Open_Price','High_Price','Low_Price','Volume'])
        data_perf_1 = data_perf_0.query("Date == @start_date")  
        data_perf_2 = data_perf_0.query("Date == @end_date")              
        data_perf_join = pd.merge(data_perf_1, data_perf_2, on=['Ticker', 'Asset_Name','Sector'], how='inner')       
        data_perf_join = data_perf_join[['Date_x','Date_y','Ticker','Asset_Name','Sector','Close_Price_x','Close_Price_y']].rename(columns = {'Date_x':'Date_1','Date_y':'Date_2','Close_Price_x':'Close_1','Close_Price_y':'Close_2'})        
        ror = [round((data_perf_join.iloc[j,6]-data_perf_join.iloc[j,5])*100/data_perf_join.iloc[j,5],2) for j in range(0,len(data_perf_join))]
        data_perf_join['Return %'] = ror
        now = format(datetime.now().strftime('%d-%m-%Y %H-%M-%S'))
        data_perf_join.to_excel(os.path.join(cwd, "Portfolio_Performance " + now + ".xlsx"))
        messagebox.showinfo("Portfolio_Performance", "Portfolio Performance has been saved as: " + os.path.join(cwd, "Portfolio_Performance " + now + ".xlsx"))
    
    def plot(self):
        #plotting portfolio returns
        if msg == "Button Clicked: Plot Returns":  
            data_Q = [data_ror.query(f'Ticker == @lst_symbol_list{[k]}') for k in range(0, len(lst_symbol_list))]
            fig = plt.figure()
            for frame in data_Q:
                plt.plot(frame['Date'], frame['Return %'])
            plt.legend(lst_symbol_list)
            plt.xlabel('Date')
            plt.ylabel('Return %')
            plt.show()

    def scenarios(self):
        #scenarios
        global data_scen_join
        start_date = time.strftime('%Y-%m-%d', time.gmtime(dates[0][0]))
        end_date = time.strftime('%Y-%m-%d', time.gmtime(dates[0][-1]))
        data_scen_0 = data_all.drop(columns = ['Open_Price','High_Price','Low_Price','Volume'])
        data_scen_1 = data_scen_0.query("Date == @start_date")  
        data_scen_2 = data_scen_0.query("Date == @end_date")              
        data_scen_join = pd.merge(data_scen_1, data_scen_2, on=['Ticker', 'Asset_Name','Sector'], how='inner')       
        data_scen_join = data_scen_join[['Date_x','Date_y','Ticker','Asset_Name','Sector','Close_Price_x','Close_Price_y']].rename(columns = {'Date_x':'Date_1','Date_y':'Date_2','Close_Price_x':'Close_1','Close_Price_y':'Close_2'})        
        ror = [round((data_scen_join.iloc[j,6]-data_scen_join.iloc[j,5])*100/data_scen_join.iloc[j,5],2) for j in range(0,len(data_scen_join))]
        data_scen_join['Return %'] = ror
        now = format(datetime.now().strftime('%d-%m-%Y %H-%M-%S'))
        data_scen_join.to_excel(os.path.join(cwd, "Portfolio_Scenarios " + now + ".xlsx"))
        messagebox.showinfo("Portfolio_Scenarios", "Portfolio Scenarios has been saved as: " + os.path.join(cwd, "Portfolio_Scenarios " + now + ".xlsx"))

    def statistics(self):
        #some key metrics and ratios
        global statistics
        apiBase = 'https://query2.finance.yahoo.com'
        headers = { 
        "User-Agent": 
        "Mozilla/5.0 (Windows NT 6.1; Win64; x64)"
        }
        def getCredentials(cookieUrl='https://fc.yahoo.com', crumbUrl=apiBase+'/v1/test/getcrumb'):
            cookie = requests.get(cookieUrl).cookies
            crumb = requests.get(url=crumbUrl, cookies=cookie, headers=headers).text
            return {'cookie': cookie, 'crumb': crumb}
        def quote(symbols, credentials):
            url = apiBase + '/v7/finance/quote'
            params = {'symbols': ','.join(symbols), 'crumb': credentials['crumb']}
            response = requests.get(url, params=params, cookies=credentials['cookie'], headers=headers)
            quotes = response.json()['quoteResponse']['result']
            return quotes
        credentials = getCredentials()
        quotes = quote(sorted(sum(df_constituents_base.drop(['Name', 'Sector'], axis = 1).to_numpy().tolist(), [])), credentials)
        if quotes:
            for quote in quotes:
                stats = [pd.DataFrame({"Ticker": [quotes[i]['symbol']], "Currency": [quotes[i]['currency']],\
                    "52 Week Low": [quotes[i]['fiftyTwoWeekLow']], "52 Week High": [quotes[i]['fiftyTwoWeekHigh']],\
                    "Trailing P/E": [quotes[i]['trailingPE']], "Forward P/E": [quotes[i]['forwardPE']],\
                    "EPS Trailing 12m": [quotes[i]['epsTrailingTwelveMonths']], "EPS Current Year": [quotes[i]['epsCurrentYear']],\
                    "EPS Forward": [quotes[i]['epsForward']], "Book Value": [quotes[i]['bookValue']],\
                    "Price to Book": [quotes[i]['priceToBook']], "Current Price": [quotes[i]['regularMarketPrice']]}) for i in range(0, len(quotes))]
        statistics = pd.concat(stats, ignore_index=False)
        
    def markets(self):
        #a brief overview of whats going on in global indices, currencies, crypto and bond markets.
        global markets
        apiBase = 'https://query2.finance.yahoo.com'
        headers = { 
        "User-Agent": 
        "Mozilla/5.0 (Windows NT 6.1; Win64; x64)"
        }
        def getCredentials(cookieUrl='https://fc.yahoo.com', crumbUrl=apiBase+'/v1/test/getcrumb'):
            cookie = requests.get(cookieUrl).cookies
            crumb = requests.get(url=crumbUrl, cookies=cookie, headers=headers).text
            return {'cookie': cookie, 'crumb': crumb}
        def quote(symbols, credentials):
            url = apiBase + '/v7/finance/quote'
            params = {'symbols': ','.join(symbols), 'crumb': credentials['crumb']}
            response = requests.get(url, params=params, cookies=credentials['cookie'], headers=headers)
            quotes = response.json()['quoteResponse']['result']
            return quotes
        credentials = getCredentials()
        if msg == "Button Clicked: Currencies":
            quotes = quote(['EURUSD=X', 'EURGBP=X', 'EURJPY=X', 'EURCAD=X', 'EURAUD=X', 'EURMXN=X', 'EURSEK=X', 'EURNOK=X', 'EURCHF=X', 'EURHKD=X', 'EURNZD=X'], credentials)
        elif msg == "Button Clicked: Cryptocurrencies":
            quotes = quote(['BTC-USD', 'ETH-USD', 'USDT-USD', 'SOL-USD', 'BNB-USD', 'DOGE-USD'], credentials)
        elif msg == "Button Clicked: Global Indices":
            quotes = quote(['^GSPC', '^DJI', '^IXIC', '^VIX', '^FTSE', '^FCHI', '^N225', '^GDAXI', '^NDX', '000300.SS', '^BVSP'], credentials)
        elif msg == "Button Clicked: Bond Markets":
            quotes = quote(['^IRX', '^FVX', '^TYX'], credentials)
        if quotes:
            for quote in quotes:
                mark = [pd.DataFrame({"Name": [quotes[i]['longName']],\
                    "52 Week Low": [quotes[i]['fiftyTwoWeekLow']], "52 Week High": [quotes[i]['fiftyTwoWeekHigh']],\
                    "Current Price": [quotes[i]['regularMarketPrice']], "% Change": [quotes[i]['regularMarketChangePercent']]}) for i in range(0, len(quotes))]
        markets = pd.concat(mark, ignore_index=False)

port_manager()



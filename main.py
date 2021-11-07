from datetime import datetime
from datetime import timedelta
from dateutil.relativedelta import relativedelta
import pandas as pd
import xlrd
import yfinance as yf
from yahoofinancials import YahooFinancials
import matplotlib.pyplot as plt
import seaborn
import keyboard
import os
from operator import itemgetter

def clear():  
   os.system('cls' if os.name == 'nt' else 'clear')

def SaveFiles(dataframes):
    spy = dataframes["spy"]
    vss =  dataframes["vss"]
    scz =  dataframes["scz"]
    osmax = dataframes["osmax"]
    tlt =  dataframes["tlt"]

    spy.to_excel("spy.xlsx", sheet_name="SPY_Year_Price")
    vss.to_excel("vss.xlsx", sheet_name="VSS_Year_Price")
    scz.to_excel("scz.xlsx", sheet_name="SCZ_Year_Price")
    osmax.to_excel("osmax.xlsx", sheet_name="OSMAX_Year_Price")
    tlt.to_excel("tlt.xlsx", sheet_name="TLT_Year_Price")


def DownloadData():
    finaldate = datetime.today()    
    yearbackdate = (finaldate - timedelta(days=365)).strftime('%Y-%m-%d')
    sixmonthsbackdate = (finaldate - relativedelta(months = 6)).strftime('%Y-%m-%d')
    threemonthsbackdate = (finaldate - relativedelta(months = 3)).strftime('%Y-%m-%d')
    monthbackdate = (finaldate - relativedelta(months = 1)).strftime('%Y-%m-%d')
    #TODO --> Save six, three and one month data in excel and retrieve it
    finaldate = finaldate.strftime('%Y-%m-%d')
    fds = str(finaldate)
    ids = str(yearbackdate)

    spy = yf.download('SPY', period='1y', interval='1d' , start=ids, end=fds)
    vss = yf.download('VSS', period='1y', interval='1d' , start=ids, end=fds)
    scz = yf.download('SCZ', period='1y', interval='1d' , start=ids, end=fds)
    osmax = yf.download('OSMAX', period='1y', interval='1d' , start=ids, end=fds)
    tlt = yf.download('TLT', period='1y', interval='1d' , start=ids, end=fds)

    data = {
        "spy": spy,
        "vss": vss,
        "scz": scz,
        "osmax": osmax,
        "tlt": tlt
    }
    SaveFiles(data)


def ReadFiles():
    spy = pd.read_excel("spy.xlsx")
    vss = pd.read_excel("vss.xlsx")
    scz = pd.read_excel("scz.xlsx")
    osmax = pd.read_excel("osmax.xlsx")
    tlt = pd.read_excel("tlt.xlsx")

    data = {
        "spy": spy,
        "vss": vss,
        "scz": scz,
        "osmax": osmax,
        "tlt": tlt
    }
    return data


def GetReturn(initialprice, finalprice):
    deltaprice = finalprice - initialprice
    result = deltaprice * 100 / initialprice
    return result


def PerformanceList():
    #GetSavedData
    data = ReadFiles();

    #Get Number of columns
    spycolumnsnumber = data['spy'][data['spy'].columns[0]].count()
    vsscolumnsnumber = data['vss'][data['vss'].columns[0]].count()
    sczcolumnsnumber = data['scz'][data['scz'].columns[0]].count()
    osmaxcolumnsnumber = data['osmax'][data['osmax'].columns[0]].count()
    tltcolumnsnumber = data['tlt'][data['tlt'].columns[0]].count()

    #Print data about ETFs
    #SPY
    initialprice = data['spy']['Close'][0]
    finalprice = data['spy']['Close'][spycolumnsnumber - 1]
    result = round(GetReturn(initialprice, finalprice), 2)

    sixmonthsagodate = data['spy']['Date'][spycolumnsnumber - 1] - relativedelta(months = 6)
    threemonthsagodate = data['spy']['Date'][spycolumnsnumber - 1] - relativedelta(months = 3)
    onemonthagodate = data['spy']['Date'][spycolumnsnumber - 1] - relativedelta(months = 1)

    if sixmonthsagodate.weekday() == 5:
        sixmonthsagodate -= relativedelta(days = 1)    
    elif sixmonthsagodate.weekday() == 6:
        sixmonthsagodate + relativedelta(days = 1)
    
    if threemonthsagodate.weekday() == 5:
        threemonthsagodate -= relativedelta(days = 1)    
    elif threemonthsagodate.weekday() == 6:
        sixmonthsagodate + relativedelta(days = 1)
    
    if onemonthagodate.weekday() == 5:
        onemonthagodate -= relativedelta(days = 1)    
    elif onemonthagodate.weekday() == 6:
        onemonthagodate + relativedelta(days = 1)
    
    pricesix = 0
    pricethree = 0
    priceone = 0

    for i in range(0, spycolumnsnumber):
        if data['spy']['Date'][i] == sixmonthsagodate:
            pricesix = data['spy']['Close'][i]

        if data['spy']['Date'][i] == threemonthsagodate:
            pricethree = data['spy']['Close'][i]

        if data['spy']['Date'][i] == onemonthagodate:
            priceone = data['spy']['Close'][i]



    semiannualperformance = round(GetReturn(pricesix, finalprice),2)
    threemonthsperformance = round(GetReturn(pricethree, finalprice),2)
    monthlyperformance = round(GetReturn(priceone, finalprice),2)
    performancesum = semiannualperformance + threemonthsperformance + monthlyperformance
    
    spy = {
        'name': 'SPY',
        'semiannualperformance': semiannualperformance,
        'threemonthsperformance': threemonthsperformance,
        'monthlyperformance': monthlyperformance,
        'performancesum': performancesum
    }

    etfslist = list()
    etfslist.append(spy)


    #VSS
    initialprice = data['vss']['Close'][0]
    finalprice = data['vss']['Close'][vsscolumnsnumber - 1]

    sixmonthsagodate = data['vss']['Date'][vsscolumnsnumber - 1] - relativedelta(months = 6)
    threemonthsagodate = data['vss']['Date'][vsscolumnsnumber - 1] - relativedelta(months = 3)
    onemonthagodate = data['vss']['Date'][vsscolumnsnumber - 1] - relativedelta(months = 1)

    if sixmonthsagodate.weekday() == 5:
        sixmonthsagodate -= relativedelta(days = 1)    
    elif sixmonthsagodate.weekday() == 6:
        sixmonthsagodate += relativedelta(days = 1)
    
    if threemonthsagodate.weekday() == 5:
        threemonthsagodate -= relativedelta(days = 1)    
    elif threemonthsagodate.weekday() == 6:
        threemonthsagodate + relativedelta(days = 1)
    
    if onemonthagodate.weekday() == 5:
        onemonthagodate -= relativedelta(days = 1)    
    elif onemonthagodate.weekday() == 6:
        onemonthagodate + relativedelta(days = 1)
    

    for i in range(0, vsscolumnsnumber):
        if data['vss']['Date'][i] == sixmonthsagodate:
            pricesix = data['vss']['Close'][i]

        if data['vss']['Date'][i] == threemonthsagodate:
            pricethree = data['vss']['Close'][i]

        if data['vss']['Date'][i] == onemonthagodate:
            priceone = data['vss']['Close'][i]


    semiannualperformance = round(GetReturn(pricesix, finalprice),2)
    threemonthsperformance = round(GetReturn(pricethree, finalprice),2)
    monthlyperformance = round(GetReturn(priceone, finalprice),2)
    performancesum = semiannualperformance + threemonthsperformance + monthlyperformance

    vss = {
        'name': 'VSS',
        'semiannualperformance': semiannualperformance,
        'threemonthsperformance': threemonthsperformance,
        'monthlyperformance': monthlyperformance,
        'performancesum': performancesum
    }
    etfslist.append(vss)

    #SCZ
    initialprice = data['scz']['Close'][0]
    finalprice = data['scz']['Close'][sczcolumnsnumber - 1]    
    result = round(GetReturn(initialprice, finalprice), 2)

    sixmonthsagodate = data['scz']['Date'][sczcolumnsnumber - 1] - relativedelta(months = 6)
    threemonthsagodate = data['scz']['Date'][sczcolumnsnumber - 1] - relativedelta(months = 3)
    onemonthagodate = data['scz']['Date'][sczcolumnsnumber - 1] - relativedelta(months = 1)

    if sixmonthsagodate.weekday() == 5:
        sixmonthsagodate -= relativedelta(days = 1)    
    elif sixmonthsagodate.weekday() == 6:
        sixmonthsagodate + relativedelta(days = 1)
    
    if threemonthsagodate.weekday() == 5:
        threemonthsagodate -= relativedelta(days = 1)    
    elif threemonthsagodate.weekday() == 6:
        threemonthsagodate + relativedelta(days = 1)
    
    if onemonthagodate.weekday() == 5:
        onemonthagodate -= relativedelta(days = 1)    
    elif onemonthagodate.weekday() == 6:
        onemonthagodate + relativedelta(days = 1)
    

    for i in range(0, sczcolumnsnumber):
        if data['scz']['Date'][i] == sixmonthsagodate:
            pricesix = data['scz']['Close'][i]

        if data['scz']['Date'][i] == threemonthsagodate:
            pricethree = data['scz']['Close'][i]

        if data['scz']['Date'][i] == onemonthagodate:
            priceone = data['scz']['Close'][i]



    semiannualperformance = round(GetReturn(pricesix, finalprice),2)
    threemonthsperformance = round(GetReturn(pricethree, finalprice),2)
    monthlyperformance = round(GetReturn(priceone, finalprice),2)
    performancesum = semiannualperformance + threemonthsperformance + monthlyperformance

    scz = {
        'name': 'SCZ',
        'semiannualperformance': semiannualperformance,
        'threemonthsperformance': threemonthsperformance,
        'monthlyperformance': monthlyperformance,
        'performancesum': performancesum
    }
    etfslist.append(scz)

    #OSMAX
    initialprice = data['osmax']['Close'][0]
    finalprice = data['osmax']['Close'][osmaxcolumnsnumber - 1]
    result = round(GetReturn(initialprice, finalprice), 2)

    sixmonthsagodate = data['osmax']['Date'][osmaxcolumnsnumber - 1] - relativedelta(months = 6)
    threemonthsagodate = data['osmax']['Date'][osmaxcolumnsnumber - 1] - relativedelta(months = 3)
    onemonthagodate = data['osmax']['Date'][osmaxcolumnsnumber - 1] - relativedelta(months = 1)

    if sixmonthsagodate.weekday() == 5:
        sixmonthsagodate -= relativedelta(days = 1)    
    elif sixmonthsagodate.weekday() == 6:
        sixmonthsagodate + relativedelta(days = 1)
    
    if threemonthsagodate.weekday() == 5:
        threemonthsagodate -= relativedelta(days = 1)    
    elif threemonthsagodate.weekday() == 6:
        threemonthsagodate + relativedelta(days = 1)
    
    if onemonthagodate.weekday() == 5:
        onemonthagodate -= relativedelta(days = 1)    
    elif onemonthagodate.weekday() == 6:
        onemonthagodate + relativedelta(days = 1)
    
    
    for i in range(0, osmaxcolumnsnumber):
        if data['osmax']['Date'][i] == sixmonthsagodate:
            pricesix = data['osmax']['Close'][i]

        if data['osmax']['Date'][i] == threemonthsagodate:
            pricethree = data['osmax']['Close'][i]

        if data['osmax']['Date'][i] == onemonthagodate:
            priceone = data['osmax']['Close'][i]


    semiannualperformance = round(GetReturn(pricesix, finalprice),2)
    threemonthsperformance = round(GetReturn(pricethree, finalprice),2)
    monthlyperformance = round(GetReturn(priceone, finalprice),2)
    performancesum = semiannualperformance + threemonthsperformance + monthlyperformance

    osmax = {
        'name': 'OSMAX',
        'semiannualperformance': semiannualperformance,
        'threemonthsperformance': threemonthsperformance,
        'monthlyperformance': monthlyperformance,
        'performancesum': performancesum
    }
    etfslist.append(osmax)

    #TLT
    initialprice = data['tlt']['Close'][0]
    finalprice = data['tlt']['Close'][tltcolumnsnumber - 1]
    result = round(GetReturn(initialprice, finalprice), 2)
    
    sixmonthsagodate = data['tlt']['Date'][tltcolumnsnumber - 1] - relativedelta(months = 6)
    threemonthsagodate = data['tlt']['Date'][tltcolumnsnumber - 1] - relativedelta(months = 3)
    onemonthagodate = data['tlt']['Date'][tltcolumnsnumber - 1] - relativedelta(months = 1)

    if sixmonthsagodate.weekday() == 5:
        sixmonthsagodate -= relativedelta(days = 1)    
    elif sixmonthsagodate.weekday() == 6:
        sixmonthsagodate + relativedelta(days = 1)
    
    if threemonthsagodate.weekday() == 5:
        threemonthsagodate -= relativedelta(days = 1)    
    elif threemonthsagodate.weekday() == 6:
        threemonthsagodate + relativedelta(days = 1)
    
    if onemonthagodate.weekday() == 5:
        onemonthagodate -= relativedelta(days = 1)    
    elif onemonthagodate.weekday() == 6:
        onemonthagodate + relativedelta(days = 1)
    

    for i in range(0, tltcolumnsnumber):
        if data['tlt']['Date'][i] == sixmonthsagodate:
            pricesix = data['tlt']['Close'][i]

        if data['tlt']['Date'][i] == threemonthsagodate:
            pricethree = data['tlt']['Close'][i]

        if data['tlt']['Date'][i] == onemonthagodate:
            priceone = data['tlt']['Close'][i]

    
    semiannualperformance = round(GetReturn(pricesix, finalprice),2)
    threemonthsperformance = round(GetReturn(pricethree, finalprice),2)
    monthlyperformance = round(GetReturn(priceone, finalprice),2)
    performancesum = semiannualperformance + threemonthsperformance + monthlyperformance

    tlt = {
        'name': 'TLT',
        'semiannualperformance': semiannualperformance,
        'threemonthsperformance': threemonthsperformance,
        'monthlyperformance': monthlyperformance,
        'performancesum': performancesum
    }
    etfslist.append(tlt)

    def mySort(e):
        return e['performancesum']

    etfslist.sort(reverse=True, key=mySort)

    for etf in etfslist:
        print(etf['name'])
        print("Rendimiento semestral: " + str(etf['semiannualperformance']))
        print("Rendimiento trimestral: " + str(etf['threemonthsperformance']))
        print("Rendimiento mensual: " + str(etf['monthlyperformance']))
        print("Sumatoria: " + str(etf['performancesum']))
        print('\n')
    
    print("\n Presione una tecla para continuar")

    keyboard.read_key()

    clear();


isrunning = True;
while isrunning:
    print("0-Salir")
    print("1-Actualizar información")
    print("2-Obtener lista de rendimientos")

    keyboard.read_key()
    
    if keyboard.is_pressed("0"):
        break;

    if keyboard.is_pressed("1"):
        DownloadData();

    if keyboard.is_pressed("2"):
        PerformanceList();
    
    clear()


'''
for i in range(0, columnsnumber):
    date = data['spy'].iloc[i].name
    price = data['spy'].iloc[i]['Close']
    print(str(date) + "  " + str(price))

'''

'''
for s in spycloses:
    print(s.name + " " + s)

'''

'''
vsscloses = vssdata['Close']
sczcloses = sczdata['Close']
osmaxcloses = osmaxdata['Close']
tltcloses = tltdata['Close']


print(spycloses)
print(vsscloses)
print(sczcloses)
print(osmaxcloses)
print(tltcloses)
'''

#closes.plot(figsize=(16,9))
#plt.title('spy')
#plt.show()
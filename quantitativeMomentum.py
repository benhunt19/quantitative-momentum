# calculation, file creation imports
import numpy as np
import pandas as pd
import requests 
import xlsxwriter
import math
from scipy import stats
import yfinance as yf
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import seaborn as sns; sns.set()
from forex_python.converter import CurrencyRates
import os

#email related imports
import ssl
from email.message import EmailMessage
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

### getting the SnP 
import bs4 as bs
import pickle

## ~~~~~~~~~~~~~~~~~~~ GLOBAL FUNCTIONS ~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ##

# this is how much money to invest
def portfolioIput():
    global portfolioSize
    canContinue = False
    portfolioSize = input('Enter the size of your portfolio (£)')
    while not canContinue:
        try:
            portfolioSize = int(portfolioSize)
            print('portfolioSize: ',portfolioSize)
            canContinue = True
        except:
            portfolioSize = input('Enter the size of your portfolio (£) (must be numeric)') 
#portfolioIput()

# this is how many stocks to spread the investment over
def stockLimitInput():
    global stockLimit
    canContinue = False
    stockLimit = input('How many stocks do you wish to invest in?')
    while not canContinue:
        try:
            stockLimit = int(stockLimit)
            print('stockLimit: ',stockLimit)
            canContinue = True
        except:
            stockLimit = input('How many stocks do you wish to invest in? (needs to be an integer)') 
#stockLimitInput()    

# this is depreciated
def exchange(amount, currency):
    try:  
        rate = CurrencyRates().get_rate('GBP', currency)
    except:
        rate = 1
    return amount * rate
exchange(123, 'USD')

def makeTodaysFolder():
    global today
    today = datetime.today().strftime('%Y-%m-%d')
    try:
        current = os.getcwd()
        os.mkdir(current+ f'/{today}')
    except:
        pass 
makeTodaysFolder()

def emailWithAttachment(subject, email_receiver, email_receiver_name,  filename):
    # Define email sender and receiver
    email_sender = '** removed from public repo**'
    email_password =  '** removed from public repo**'

    hello = f"Hi {email_receiver_name}, \n\n"
    body = """
    Here are the findings of this months finance calculations, this is not financial advice

    This was an automated email sent via Python, please see 'About & Audit' tab on excel attachment.
    """
    emailText = hello + body

    
    message = MIMEMultipart()
    message['From'] = email_sender
    message['To'] = email_receiver
    message['Subject'] = subject
    message.attach(MIMEText(emailText, "plain"))
        
    # Open PDF file in binary mode
    
    # We assume that the file is in the directory where you run your Python script from
    with open(filename, "rb") as attachment:
        # The content type "application/octet-stream" means that a MIME attachment is a binary file
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    
    # Encode to base64
    encoders.encode_base64(part)
    
    # Add header 
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )
    
    # Add attachment to your message and convert it to string
    message.attach(part)
    #text = em.as_string()

    # Add SSL (layer of security)
    context = ssl.create_default_context()

    # Log in and send the email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, email_receiver, message.as_string())

# getting current SNP500 
def get_sp500_tickers():
    global wikiLink
    wikiLink = 'https://en.wikipedia.org/wiki/List_of_S%26P_500_companies'
    resp = requests.get(wikiLink)
    soup = bs.BeautifulSoup(resp.text, 'lxml')
    global wikiTableRaw
    wikiTableRaw = soup.find('table', {'class': 'wikitable sortable sticky-header'})
    tickers = []
    wikiCols = [title.text.strip() for title in wikiTableRaw.find('tr').findAll('th')] # array of all headings
    wikiCols.append('Link')
    wikiTable = pd.DataFrame(columns=wikiCols)

    for row in wikiTableRaw.findAll('tr')[1:]:
        tmpArr = []
        for colNum, col in enumerate(wikiCols[0:len(wikiCols) - 1]):
            tmpArr.append(row.findAll('td')[colNum].text.strip())
        link = row.find('a', {"class":"external text"}).get("href")
        tmpArr.append(link)
        tmpSeries = pd.Series(tmpArr, index=wikiCols)
        wikiTable.loc[len(wikiTable)] = tmpSeries

    with open("sp500tickers.pickle","wb") as f:
        pickle.dump(tickers,f)

    return wikiTable

## ~~~~~~~~~~~~~~~~~~~~~~~ ASK QUESTIONS ~~~~~~~~~~~~~~~~~~~~~~~~~ ##

stockLimitInput() # this sets var stockLimit
portfolioIput() # this to set var portfolioSize50

## ~~~~~~~~~~~~~~~~~~~~~~~GET SNP ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~##

df = get_sp500_tickers()

## ~~~~~~~~~~~~~~~ DEFINE DATAFRAMES AND GET DATA ~~~~~~~~~~~~~~~~~~##


columns = [
    'Ticker', 
    'Company Name',
    'Price', 
    'Currency',
    'Link',
    'Number of Shares to Buy',
    'One-Year Price Return', 
    'One-Year Price Percentile',
    'Six-Month Price Retun',
    'Six-Month Price Percentile',
    'Three-Month Price Retun',
    'Three-Month Price Percentile',
    'One-Month Price Retun',
    'One-Month Price Percentile',
    'Average Percentile'
]
finalDf = pd.DataFrame(columns=columns)
naDf = pd.DataFrame(columns=columns)

length = df.shape[0] # set to 30 for testing...
tickerArr = []

for i in range(length):
    tmpSym = df.iloc[i,0]
    tickerArr.append(tmpSym)
stringAgg = ' '.join(tickerArr)

oneMonthAgo = (datetime.now() - timedelta(days=30)).strftime('%Y-%m-%d')
threeMonthsAgo = (datetime.now() - timedelta(days=90)).strftime('%Y-%m-%d')
sixMonthsAgo = (datetime.now() - timedelta(days=180)).strftime('%Y-%m-%d')

timeFrames = [oneMonthAgo, threeMonthsAgo, sixMonthsAgo]

multi = yf.Tickers(stringAgg)

for i in range(length):
    tmpSym = df.iloc[i,0]
    try:
        #dict[tmpSym] = tmpChange
        tmpTic  = multi.tickers[tmpSym]
        curPrice = tmpTic.basic_info.last_price
        tmpSeries = pd.Series([
                        tmpSym, 
                        df.loc[i,'Security'],
                        curPrice,
                        tmpTic.basic_info.currency,
                        df.loc[i,'Link'],
                        'N/A',
                        tmpTic.basic_info.year_change,
                        'X',
                        (curPrice / tmpTic.history(start=sixMonthsAgo).iloc[0]['Open']) - 1,
                        'X',
                        (curPrice / tmpTic.history(start=threeMonthsAgo).iloc[0]['Open']) - 1,
                        'X',
                        (curPrice / tmpTic.history(start=oneMonthAgo).iloc[0]['Open']) - 1,
                        'X',
                        'X'
                    ],
                        index=columns
                    )
        finalDf.loc[len(finalDf)] = tmpSeries
        #print(tmpSeries)
        

    except:
        tmpSeries = pd.Series([
                        tmpSym, 
                        'Price', 
                        'Company Name',
                        'Currency',
                        'Link',
                        'Number of Shares to Buy',
                        'One-Year Price Return',
                        'One-Year Price Percentile',
                        'Six-Month Price Retun',
                        'Six-Month Price Percentile',
                        'Three-Month Price Retun',
                        'Three-Month Price Percentile',
                        'One-Month Price Retun',
                        'One-Month Price Percentile',
                        'Average Percentile'
                    ],
                        index=columns
                    )
        naDf.loc[len(finalDf)] = tmpSeries
        
## ~~~~~~~~~~~~~~~ CALCULATE THE PERCENTILES ~~~~~~~~~~~~~~##

percentileColumnsX = [column for column in columns if 'Percentile' in column.split(' ')] 
percentileColumnsX

percentileColumns = [
    {
        "return": "One-Year Price Return",
        "percentile": "One-Year Price Percentile"
    },
    {
        "return": "Six-Month Price Retun",
        "percentile": "Six-Month Price Percentile"
    },
    {
        "return": "Three-Month Price Retun",
        "percentile": "Three-Month Price Percentile"
    },
    {
        "return": "One-Month Price Retun",
        "percentile": "One-Month Price Percentile"
    }
]

finalDf.loc[3, 'One-Year Price Percentile'] = 10
finalDf

# Using scipy.stats to work out percentiles 
for row in finalDf.index:
    tmpTotalPercentage = 0
    for col in percentileColumns:
        tmpPercentil = stats.percentileofscore(finalDf[col["return"]], finalDf.loc[row , col["return"]])
        finalDf.loc[row , col["percentile"]] = tmpPercentil
        tmpTotalPercentage += tmpPercentil
    finalDf.loc[row ,"Average Percentile"] = tmpTotalPercentage / len(percentileColumns)

## ~~~~~~~~~~~~~~~~~~~~ FILTERING AND ORDERING ~~~~~~~~~~~~~##

finalDfFiltered = finalDf.loc[(finalDf['One-Year Price Return'] != 'N/A') & (finalDf['One-Year Price Return'])]
finalDfFilteredOrdered = finalDfFiltered.sort_values('Average Percentile', ascending = False)

finalDfFilteredOrderedLimited = finalDfFilteredOrdered[:stockLimit]
finalDfFilteredOrderedLimited.reset_index(inplace=True, drop=True)

finalDfFilteredOrderedLimited50 = finalDfFilteredOrdered[:50]
finalDfFilteredOrderedLimited50.reset_index(inplace=True, drop=True)

## ~~~~~~~~~~~~ CALCULATE THE NUMBER OF SHARES TO BUY ~~~~~~~~~~~~##

curLen = len(finalDfFilteredOrderedLimited.index) # incase there is a differnece between stockLimit and the len of df ??
positionSize = float(portfolioSize) / curLen # this is the amount we have to spend on each stock
for i in range(curLen):
    print(finalDfFilteredOrderedLimited.loc[i, 'Currency'])
    positionSizeConverted = exchange(positionSize, finalDfFilteredOrderedLimited.loc[i, 'Currency'])
    tmpAmount = positionSizeConverted / finalDfFilteredOrderedLimited.loc[i, 'Price']
    finalDfFilteredOrderedLimited.loc[i, 'Currency']
    finalDfFilteredOrderedLimited.loc[i, 'Number of Shares to Buy'] = tmpAmount
# this will return fractional shares, add math.floor when a baller with a huge portfolio 


## ~~~~~~~~~~~~ CREATE EXCEL REPORT AND PRICE PLOTS  ~~~~~~~~~~~~~~##

makeTodaysFolder() # also creates global variable of today
fileName = f"./{today}/stock_report_{today}.xlsx"
writer = pd.ExcelWriter(fileName , engine="xlsxwriter")

# creating the top 50 sheet
top50sheetName = 'Top 50'
finalDfFilteredOrderedLimited50.to_excel(writer, sheet_name=top50sheetName)
worksheet = writer.sheets[top50sheetName]
for colnum, column in enumerate(finalDfFilteredOrderedLimited50.columns):
    len(list(column))
    worksheet.set_column(colnum + 1, colnum + 1, len(list(column)) + 1)



# creating the individual stock sheets
for ticker in finalDfFilteredOrderedLimited['Ticker']:
    # iterating throuhg the top 
    mainDf = finalDfFilteredOrderedLimited
    mainDf.loc[mainDf['Ticker'] == ticker].to_excel(writer, sheet_name=ticker)  # Default position, cell A1.
    worksheet = writer.sheets[ticker]

    # setting column widths of the sheet
    for colnum, column in enumerate(mainDf.columns):
        len(list(column))
        worksheet.set_column(colnum + 1, colnum + 1, len(list(column)) + 1)

    worksheet.set_column(0,0,3)

        #link  = str(mainDf.loc[mainDf.Ticker == ticker, 'Link'][0])
        #print(link)
        #worksheet.write(3, 0, link)

    
    for index, date in enumerate(timeFrames):
        # creating, saving and inserting the files into the sheet
        plt.figure(figsize=(10,4))
        yf.Ticker(ticker).history(start=date)['Open'].plot()
        plt.title(f'{ticker} - starting {date}')
        plt.legend([ticker])
        plt.ylabel(yf.Ticker(ticker).basic_info['currency'])
        imageFileName = f"{today}/{ticker}_{date}_{today}.jpg"
        plt.savefig(imageFileName)
        worksheet.insert_image("A" + str(4 + 17 * index), imageFileName, {"x_scale": 0.75, "y_scale": 0.75})
        #plt.show()


#creating the About / Audit
helpAuditSheetName = 'About & Audit'
worksheet = writer.book.add_worksheet(helpAuditSheetName)
worksheet.set_column(0,0,100)

autitMessages = [
    F'This is not financial advice',
    f'Nuber of stocks selected to invest in: {stockLimit}',
    f'Selected Portfolio Size (£): {portfolioSize}',
    f'Tickers evalueted via: {wikiLink}',
    f'Stock data added via the yFinance python library',
    f'This was run on: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}'
    
]

for index, message in enumerate(autitMessages):
    worksheet.write(index,0,message)


writer.close()
print('Workbook is awaiting')


## ~~~~~~~~~~~~~ EMAIL MAILNIG LIST ~~~~~~~~~~~~~~##

emailSubject = f'{today} - Quantitative Momentum Report'
mailingCols = ['Name', 'Email', 'Active']
mailingListData = [
    ['Ben', 'ben.hunt19@hotmail.com', False]
]
#building the dataframe to store the data, does't need to be a DataFrame but could be useful for expansion
mailingDf = pd.DataFrame(columns=mailingCols)
for j in range(len(mailingListData)):
    mailingDf.loc[len(mailingDf)] = pd.Series(mailingListData[j], index=mailingCols)

for i in range(len(mailingDf)):
    if mailingDf.loc[i]['Active']:
        emailWithAttachment(emailSubject, mailingDf.loc[i]['Email'], mailingDf.loc[i]['Name'], fileName)
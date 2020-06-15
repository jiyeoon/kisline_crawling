import os
from datetime import datetime
import pandas as pd 
import numpy as np 
import win32com.client
from copy_of_dcr import GiveMeInfo, SelDriver, SearchAndFind

seldriver = SelDriver()
driver = seldriver.control()


def getDataFrame():
    dir = os.path.dirname(os.path.abspath('__file__'))
    filepath = dir + r'\asset\static\Korea GB Account CRM Extraction File_20200212.xlsx'
    print("Reading ..."+filepath)
    try:
        df = pd.read_excel(filepath)
        return df
    except:
        print("File open error!")


def getDDE(aename):
    if aename == 'Jihun Kang' or aename == 'Steven Park' or aename == 'Jongsu Hwang':
        return 'Diane Kim'
    elif aename == 'Jungrye Jung' or aename == 'Alyssa Lee' or aename == 'Young Ae Kim':
        return 'Erin Lim'
    elif aename == 'Hweejae Yim' or aename == 'Tae Seung Kim' or aename == 'Kaon Kim':
        return 'Jeongha Lim'
    elif aename == 'Junhyung Byun' or aename == 'Sinjung Choi' or aename == 'Woojoon Chu' or aename == 'Yangkyung Lee' or aename == 'Jungsun Yoon':
        return 'Lee Jae'
    else:
        return 'N/A'


def search(search_term, df):
    #df = getDataFrame()
    onecompany = initialize()
    if len(df[df['Tax Number'] == str(search_term)]) >= 1: #BP가 생성되어있으면
        newdf = df[df['Tax Number'] == str(search_term)]
        if(len(newdf) != 1):
            newdf = newdf.iloc[0]
        print(newdf)
        onecompany['Name'] = newdf['ORG Name (KOR)'] #search_term
        onecompany['BP#'] = int(newdf['ORG ID (BP#)'])
        onecompany['2018 turnover'] = int(newdf['Revenue Local Currency (2018)']/1000000)
        onecompany['2018 turnover'] = format(onecompany['2018 turnover'], ',')
        try:
            onecompany['AE'] = newdf['Name_For_Account_Owner'] + " " + newdf['Last_Name_For_Account_Owner']
            onecompany['AE'] = onecompany['AE'].values[0]
        except:
            onecompany['AE'] = 'N/A'

        try:
            onecompany['Industry'] = newdf['SAP_Master_Code_Text'].values[0]
        except:
            onecompany['Industry'] = newdf['SAP_Master_Code_Text']
            
        try:
            onecompany['Buying Classification'] = newdf['Gtm_Regional_Buying_Classification_Text'].values[0]
        except:
            onecompany['Buying Classification'] = newdf['Gtm_Regional_Buying_Classification_Text']
        
        onecompany['DDE'] = getDDE(onecompany['AE'])
        #print(onecompany)
        return onecompany
    else:
        print("There is No " + str(search_term) + " !!!!!!")
        onecompany = getFromKISLINE(str(search_term))
        return onecompany


def initialize():
    return {
        'Name' : 'N/A',
        'BP#' : 'N/A',
        '2018 turnover' : 'N/A',
        'AE' : 'N/A',
        'DDE' : 'N/A',
        'Industry' : 'N/A',
        'Buying Classification' : 'N/A',

        #'News Title' : 'N/A',
        #'News Url' : 'N/A',
        #'Comment' : 'N/A'
    } 

def getInfoFromFile():
    companylist = []
    CRMdf = getDataFrame() #CRM 데이터
    dir = os.path.dirname(os.path.abspath('__file__'))
    filepath = dir + r'\aa.xlsx'
    mydf = pd.read_excel(filepath)

    for i in range(len(mydf)):
        onecompany = initialize()
        onecompany['Account Name (KOR)'] = mydf.iloc[i]['Account Name (KOR)']
        onecompany = search(mydf.iloc[i]['Tax ID'], CRMdf) #search(onecompany['Account Name (KOR)'], CRMdf)
        #onecompany['News Title'] = mydf.iloc[i]['News Title']
        #onecompany['News Url'] = mydf.iloc[i]['News Url']
        #Comment = mydf.iloc[i]['Comment'].split('\n')
        #onecompany['Comment'] = Comment
        companylist.append(onecompany)
    return companylist


def getFromKISLINE(search_term):
    onecompany = initialize()

    information = GiveMeInfo(driver, search_term).solo()

    onecompany['Industry'] = IndustryMappingwithKSCode(information.kscode)
    print(information.moneydict[2018])
    onecompany['2018 turnover'] = information.moneydict[2018]
    try:
        onecompany['AE'] = getAE(onecompany['Industry'], onecompany['2018 turnover'])
    except:
        onecompany['AE'] = 'N/A'
    try:
        onecompany['2018 turnover'] = format(onecompany['2018 turnover'], ',')
    except:
        print(search_term + "\'s 2018 turnover is " + onecompany['2018 turnover'])
    onecompany['DDE'] = getDDE(onecompany['AE'])
    onecompany['Buying Classification'] = 'NNN'
    onecompany['BP#'] = 'NNN'
    onecompany['Name'] = search_term

    return onecompany


def IndustryMappingwithKSCode(kscode):
    dir = os.path.dirname(os.path.abspath('__file__'))
    filepath  = dir + r'\asset\static\★10차산업분류코드 Mapping File_최신버전_Provided by Commercial team_updated.xlsx'
    try:
        df = pd.read_excel(filepath)
        newdf = df[df['10차 KS Code'] == kscode]
        print(newdf)
        industry_text = newdf['Mc_Text'].values[0]
        return industry_text
    except:
        return "N/A"


def main():
    company_list = getInfoFromFile()

    df = pd.DataFrame(company_list)
    
    filepath = os.path.dirname(os.path.abspath('__file__'))
    filename = filepath + r'\aa1.xlsx'
    try:
        df.to_excel(filename, index = False)
    except:
        print("Fail!")


if __name__ == "__main__":
    main()
    #sendMail()
    


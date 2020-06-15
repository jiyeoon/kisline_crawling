import os
from datetime import datetime
import pandas as pd 
import numpy as np 
import win32com.client
from copy_of_dcr import GiveMeInfo, SelDriver, SearchAndFind

now = datetime.now()

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
    if len(df[df['ORG Name (KOR)'] == search_term]) >= 1: #BP가 생성되어있으면
        newdf = df[df['ORG Name (KOR)'] == search_term]
        if(len(newdf) != 1):
            newdf = newdf.iloc[0]
        print(newdf)
        onecompany['Name'] = search_term
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
        print("There is No " + search_term + " !!!!!!")
        onecompany = getFromKISLINE(search_term)
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
        'News Title' : 'N/A',
        'News Url' : 'N/A',
        'Comment' : 'N/A'
    } 


def getCompanyList():
    df = getDataFrame()
    search_term = input("검색할 기업명을 입력하세요::::>>>> \t")
    companylist = []
    onecompany = {}
    while search_term != "":
        onecompany = search(search_term, df)
        news_title = input("\n기사 제목을 입력하세요 :::>>> \t")
        onecompany['News Title'] = news_title
        news_url = input("\n뉴스 url을 입력하세요 :::>>> \t")
        onecompany['News Url'] = news_url
        onecompany['Comment'] = []
        print("Company Comment!!")
        while True:
            input_str = input(">")
            if input_str == "":
                break
            else:
                onecompany['Comment'].append(input_str)
        print(onecompany)
        companylist.append(onecompany)
        search_term = input("\n---------------\n검색할 기업명을 입력하세요:::>>> \t")
    print("Ok Bye..~~~~")
    #print(companylist)
    return companylist


def getInfoFromFile():
    companylist = []
    CRMdf = getDataFrame() #CRM 데이터
    dir = os.path.dirname(os.path.abspath('__file__'))
    filepath = dir + r'\asset\static\dailynewsclipping_search_list.xlsx'
    mydf = pd.read_excel(filepath)

    for i in range(len(mydf)):
        onecompany = initialize()
        onecompany['Name'] = mydf.iloc[i]['Account Name (KOR)']
        onecompany = search(onecompany['Name'], CRMdf)
        onecompany['News Title'] = mydf.iloc[i]['News Title']
        onecompany['News Url'] = mydf.iloc[i]['News Url']
        Comment = mydf.iloc[i]['Comment'].split('\n')
        onecompany['Comment'] = Comment
        companylist.append(onecompany)
    return companylist


def getHTML():
    html_header_text = """
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="utf-8">
        <title>[Daily News Clipping] %s년 %s월 %s일 주요 Keyword Monitoring</title>
        <link href="https://fonts.googleapis.com/css?family=Noto+Sans+KR&display=swap" rel="stylesheet">
        <style>
            body{
                font-family: 'Malgun Gothic';
                font-size: 10pt;
            }
            a{
                color: #039BE5;
            }
        </style>
        </head>

        <body>
        <p>안녕하세요? </br>
        Commercial Sales의 Social Listening 담당 이지연입니다. :)</p>
        <p> %s년 %s월 %s일 주요 Keyword Monitoring 결과를 안내드립니다. <br>
        <a href='/'>이곳</a>을 클릭하시면 더 많은 뉴스를 확인하실 수 있습니다.
        </p>
    """ % (now.year, now.month, now.day, now.year, now.month, now.day)

    #companylist = getCompanyList()
    companylist = getInfoFromFile()

    html_body_text = ""
    for onecompany in companylist:
        html_body_text += "<br>"
        try:
            keyword = onecompany['Comment'][0].split(': ')[1]
            html_body_text += "<p><strong style=\"background-color:#BDBDBD\">Keyword : [%s]</strong>" % (keyword)
        except:
            print("Keyword 불러오는거 실패")
        html_body_text += "<br><h3><a href=\'%s\'>%s</a></h3>" % (onecompany['News Url'], onecompany['News Title'])
        html_body_text += "<strong>* %s / 2018 매출액 : %s / Industry : %s / Buying Classification : %s </strong><br>" % (onecompany['Name'], onecompany['2018 turnover'], onecompany['Industry'], onecompany['Buying Classification'])
        html_body_text += "<strong>* BP# : %s / AE : %s / DDE : %s </strong><br>" % (str(onecompany['BP#']), onecompany['AE'], onecompany['DDE'])
        for onestring in onecompany['Comment']:
            html_body_text += "* " + onestring + "<br>"
        html_body_text += "</p>"
    
    html_footter_text = """
        <br><br><p>감사합니다.<br>이지연 드림</p>
        </body></html>
        """

    html_text = html_header_text + html_body_text + html_footter_text

    return html_text


def main():
    print("func main")
    html_text = getHTML()

    dir = os.path.dirname(os.path.abspath('__file__'))
    filename = dir + r'\\HTML\\' + '%s-%02d-%02d Daily News Clipping Email.html' % (now.year, now.month, now.day)
    html_file = open(filename, 'w')
    html_file.write(html_text)
    html_file.close()


def sendMail():
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "[Daily News Clipping] %s년 %s월 %s일 주요 Keyword News" % (now.year, now.month, now.day)
    newMail.HTMLBody = getHTML()
    #newMail.To = "jiyeon.lee@sap.com"
    newMail.To = "DL_011000358700000330052011E@global.corp.sap"
    #newMail.CC = 'peter.hwang@sap.com'
    #newMail.Send()
    newMail.Display(True) 


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


def getAE(industry_text, turnover):
    if (str(type(turnover)) != "<class 'int'>") or int(turnover) < 386380885200: #매출액이 없거나.. 300M 미만이면
        dir = os.path.dirname(os.path.abspath('__file__'))
        filepath = dir + r'\asset\static\AE_mapping_lower_300M.xlsx'
        df = pd.read_excel(filepath)
        try:
            print(industry_text)
            newdf = df[df['Industry Text'] == industry_text]
            print(newdf)
            print(newdf['AE Name'].values[0])
            aename = newdf['AE Name'].values[0]
        except:
            aename = 'N/A'
        return aename

    else: # this is 300M upper
        dir = os.path.dirname(os.path.abspath('__file__'))
        filepath = dir + r'\asset\static\AE_mapping_upper_300M.xlsx'
        df = pd.read_excel(filepath)
        newdf = df[df['Industry Text'] == industry_text]
        aename = newdf['AE Name'].values[0]
        return aename



def IndustryMappingwithKSCode(kscode):
    dir = os.path.dirname(os.path.abspath('__file__'))
    filepath  = dir + r'\asset\static\★10차산업분류코드 Mapping File_최신버전_Provided by Commercial team_updated'
    try:
        df = pd.read_excel(filepath)
        newdf = df[df['10차 KS Code'] == kscode]
        print(newdf)
        industry_text = newdf['Mc_Text'].values[0]
        return industry_text
    except:
        return "N/A"
        
    

if __name__ == "__main__":
    #main()
    sendMail()
    
import os.path
import time
import json
from bs4 import BeautifulSoup
from collections import OrderedDict, defaultdict
import numpy as np
import pandas as pd
import collections
import traceback
import datetime
import os
import sys
import this

from selenium import webdriver
from selenium.webdriver.ie.options import Options
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait # available since 2.4.0
from selenium.webdriver.support import expected_conditions as EC # available since 2.26.0
from selenium.webdriver.common.keys import Keys

class JsonDump:
    def __init__(self):
        dir = os.path.dirname(os.path.abspath(__file__))
        jsonfile = dir + r'\\asset\\static\\login_infomation.json'
        with open(jsonfile) as f:
            login_data = json.load(f)
        self.login_data = login_data

    def get_login_info(self):
        self.id = self.login_data['ID']
        self.password = self.login_data['Password']

    def get_MU_info(self):
        muinfo = {'MU' : self.login_data['MU'],
                  'Countrycode' : self.login_data['Countrycode'],
                  'inumber' : self.login_data['inumber']}
        return muinfo

#login info

class SelDriver(JsonDump):
    def __init__(self):
        super(SelDriver, self).__init__()
        opts = Options()
        opts.ignore_protected_mode_settings = True
        opts.ignore_zoom_level = True
        dir = os.path.dirname(os.path.abspath(__file__))
        # path = dir + r'\\asset\\selenium\\IEDriverServer.exe'
        # self.driver = webdriver.Ie(executable_path=path, ie_options=opts)

        path = dir + r'\\asset\\selenium\\chromedriver.exe'
        self.driver = webdriver.Chrome(executable_path=path, chrome_options=opts)
        self.access()
        self.login()

    def control(self):
        return self.driver

    def access(self):
        pageurl = "https://www.kisline.com"
        self.driver.get(pageurl)

    def put_id(self):
        self.driver.find_element_by_name('lgnuid').click()
        self.driver.find_element_by_name('lgnuid').send_keys(self.id)

    def put_password(self):
        self.driver.find_element_by_class_name('pwType').click()
        self.driver.find_element_by_class_name('pwType').send_keys(self.password)

    def login(self):
        self.get_login_info()
        try:
            try:
                self.driver.find_element_by_class_name('btn_ico_close').click()
            except:
                pass
            print('****LOGIN****')
            self.put_id()

            self.put_password()
            self.driver.find_element_by_xpath('//*[@id="loginForm"]/div/div/fieldset/a').click()
        except Exception as e:
            print(e)
            try:
                time.sleep(1)
                self.driver.find_element_by_xpath('//*[@id="loginForm"]/div/div/fieldset/a').click()
            except:
                print('Please check Log-in Status')
                input("Looks like already Logged in, or please Log-in my yourself, then press any key")
#login
                

class SearchAndFind:
    def __init__(self, driver, search_object):
        self.driver = driver
        self.setup_search_object(search_object)
        oncemore_search = True
        while oncemore_search != False:
            # Until oncemore_search is False, perform below.
            self.search()
            self.is_there_company()
            oncemore_search, become_empty = self.do_by_company_number()
        else:
        #when Oncemoresearch is off, this will perform.
            if become_empty == True:
                self.search_term = 'Fail|' + str(self.original_search_term)



    def setup_search_object(self,search_object):
        # Handling search_object. and perform search.
        if type(search_object) == list:  # when massive search mode
            self.search_term = search_object[0]
            if search_object[1] != '-9999':  # when org_id is null(-9999)
                self.orgid = search_object[1]
            else:
                self.orgid = ''
        else:  # When single search mode
            self.search_term = search_object
            self.orgid = ''
        self.original_search_term = self.search_term  # for put in output file
        return

    def search(self):
        try:
            self.driver.find_element_by_id('q').clear()
            self.driver.find_element_by_id('q').send_keys(self.search_term)
            self.driver.find_element_by_id('searchView').click()
            self.search_page = self.driver.page_source
        except:
            try:
                time.sleep(1)
                self.search()
            except:
                traceback.print_exception(*sys.exc_info())
        finally:
            return

    def is_there_company(self):
        soup = self.get_soup(self.search_page)
        table = soup.find(id="eprTable")
        search_row = table.find_all('tr')
        length = len(search_row)
        if (length == 2):
            a = table.find_all("tr", class_="last")
            if len(a) != 0:
                print("Only 1 Company Found. Coutinue to process")
                self.numofcompany = '1'
            else:
                print("NOTTING.")
                self.numofcompany = '0'
        else:
            self.numofcompany = '1+'


    def do_by_company_number(self):
        if self.numofcompany == "0":
            print("No Company Found. Do you want to try by other search_term?")
            new_search_term1 = input("If you want to search by other term, please put search_term, Or to pass this search, press 0   ")
            new_search_term2 = input("Once more!   ")
            # For prevent error,
            if new_search_term1 != "" and new_search_term2 == "":
                new_search_term = new_search_term1
            elif new_search_term1 == "" and new_search_term2 != "":
                new_search_term = new_search_term2
            elif new_search_term1 == "" and new_search_term2 == "":
                new_search_term = ''
            elif new_search_term1 == "0" or new_search_term2 == "0":
                new_search_term = "0"
            else:
                new_search_term = new_search_term2

            if new_search_term == '0': #When You just want to pass this search_term
                oncemore_search, become_empty = False, True
            elif new_search_term == '':
                oncemore_search, become_empty = False, False
            else: #When you want to modify this search_term
                self.search_term = new_search_term
                oncemore_search, become_empty = True, False

        elif self.numofcompany == "1": #If there is only company, automatically continue to main page
            self.driver.find_element_by_id('eprTable').find_elements_by_name('overViewOpen')[0].click()
            oncemore_search, become_empty = False, False

        else: #If there are more than one company, click manually and continue
            print("More than one company found: PLEASE CLICK COMPANY THAT YOU WANT, Then Press any Key")
            many1 = input("If you put more than 3 Character, it perform new search")
            many2 = input("To Prevent Error, Please press any key once more, to pass, PRESS 0")
            if many1 == '0' or many2 == '0': #When You just want to pass this search_term
                oncemore_search, become_empty = False, True
            elif len(many1) > 2 or len(many2) > 2:
                if len(many1) >= len(many2):
                    new_search_term = many1
                else:
                    new_search_term = many2
                self.search_term = new_search_term
                oncemore_search, become_empty = True, False
            else:
                oncemore_search, become_empty = False, False
        return oncemore_search, become_empty

    def get_soup(self, page_source):
        return BeautifulSoup(page_source, "lxml")



class GiveMeInfo(SearchAndFind):
    def __init__(self, driver, search_object):
        super().__init__(driver, search_object)

    def japan_ramen(self):
        if self.search_term[0:5] != 'Fail|':
            self.get_overview_info()
            self.driver.find_element_by_class_name("lnb").find_elements_by_tag_name("ul")[1].click()
            self.driver.find_element_by_class_name("lnb").find_elements_by_tag_name("ul")[1].click()
            self.kor_page_source = self.driver.page_source
            self.driver.find_element_by_id("langTsm").click()  # 영어페이지 클릭
            self.eng_page_source = self.driver.page_source
            self.get_info_from_both_page()
        else:
            self.assign_empty()
            return 0
#get information
            
    def solo(self):
        self.japan_ramen()
        return self

    def juju_club(self):
        result = self.japan_ramen()
        if result == 0:
            return self
        self.analyze_group_table()
        return self

    def i_love_ugb(self):
        result = self.japan_ramen()
        if result == 0:
            return self
        self.analyze_group_table()
        time.sleep(0.2)
        self.analyze_employee_tab()
        time.sleep(0.2)
        self.analyze_business_tab()
        time.sleep(0.2)
        self.analyze_thankyou_tab()
        time.sleep(0.2)
        self.analyze_boss()
        return self


    def assign_empty(self):
        self.company_name_eng = np.nan
        self.company_name_kor = np.nan
        self.kor_group_main_company = np.nan
        self.groupname_eng = np.nan
        self.companytype = np.nan
        self.address_kor = np.nan
        self.eng_add_1 = np.nan
        self.eng_add_2 = np.nan
        self.eng_add_3 = np.nan
        self.youngnam = np.nan
        self.postcode = np.nan
        self.phonenumber = np.nan
        self.fax_number = np.nan
        self.homepage = np.nan
        self.taxnumber = np.nan
        self.kscode = np.nan
        self.kscode_kor = np.nan
        self.kscode_eng = np.nan
        self.moneydict = {2015 : np.nan, 2016 : np.nan, 2017 : np.nan}  #매출액인듯
        self.eieikdict = {2015 : np.nan, 2016 : np.nan, 2017 : np.nan}
        self.sokdict = {2015 : np.nan, 2016 : np.nan, 2017 : np.nan}
        self.pskdict = {2015 : np.nan, 2016 : np.nan, 2017 : np.nan}
        self.sooneieikdict = {2015: np.nan, 2016: np.nan, 2017: np.nan}
        jujulist = []
        groupdict = {}
        for i in range(3):
            groupdict["name"], groupdict["portion"], groupdict["status"] = np.nan, np.nan, np.nan
            jujulist.append(groupdict)
        self.juju = jujulist
        self.number_of_employee = np.nan
        self.year_of_employee = np.nan
        self.ceo_name_eng_last = np.nan
        self.ceo_name_eng_first = np.nan
        self.ceo_name_kor = np.nan
        self.companystatus = np.nan
        self.items = ''
        self.employee_dict = ''
        self.businessline = ''
        self.whole_business_address = ''
        self.thankyou = ''
        self.boss_information = ''
        self.bossbirth = ''
        self.bossname = ''
        self.deepsearch = ''
        self.notable_information = ''
        self.notable_event = ''
        return

    def get_deep_search(self, soup):
        deepsearch = []
        try:
            lists = soup.find("div", class_="deep_inner").find("div", class_="right").find_all("li", class_="deepSearch_call")
            for item in lists:
                themes = item.find_all("a")
                for realitem in themes:
                    deepsearch.append(realitem.get_text().strip())
            deepsearch = "\n".join(deepsearch)
        except:
            deepsearch = 'N/A'
        self.deepsearch = deepsearch



    def get_overview_info(self):
        self.overview_page = self.driver.page_source
        soup = self.get_soup(self.overview_page)
        self.get_deep_search(soup)
        # 첫번째 테이블 선택
        basic_info_table = soup.find_all("tbody")[0]
        basic_info_table_row = basic_info_table.find_all("tr")

        # KISCODE Handle
        self.kiscode = basic_info_table_row[0].find("td").get_text().split(")", -1)[-2].split("(", -1)[-1]

        #이름 Handle
        self.company_name_kor = basic_info_table_row[0].find("td").get_text().rsplit("(", 1)[0].strip()

        #사업자등록번호 Handle
        taxnumber = basic_info_table_row[3].find('td').get_text().strip()
        if taxnumber == "-":
            self.taxnumber = "N/A"
        else:
            self.taxnumber = int(taxnumber.replace("-", "").strip())

        #기업 형태 Handle
        self.companytype = basic_info_table_row[5].find('td').get_text().strip()

        #KSCODE Handle
        row_kscode = basic_info_table_row[6].find('td').get_text().strip().split(")")
        self.kscode = row_kscode[0].replace("(", "").strip()
        self.kscode_kor = row_kscode[1].strip()
        peopleinfo = basic_info_table_row[8].get_text().strip().split("(")
        
        try:
            if ("-" in peopleinfo) or (len(peopleinfo) == 0):
                raise
            else:
                self.number_of_employee = int(peopleinfo[0].split("\n")[1].strip().replace(",","").split("명")[0].strip())
                self.year_of_employee = int(peopleinfo[1].strip().split("년")[0])
        except:
            self.number_of_employee = "N/A"
            self.year_of_employee = "N/A"

        self.items = basic_info_table_row[10].get_text().strip().replace("주요상품\n","")

        #휴폐업상태 Handle
        if "폐업" in self.companytype:
            self.companystatus = "Inactive"
        elif "피흡수합병" in self.companytype:
            self.companystatus = "Inactive - 피흡수"
        else:
            self.companystatus = "Active"

        # 전화번호 Handle
        try:
            phonenumber = basic_info_table_row[11].find('td').get_text().strip()
            if len(phonenumber) == 0:
                raise
            elif phonenumber == '-':
                raise
            elif phonenumber == 'NULL':
                raise
            else:
                self.phonenumber = phonenumber
        except:
            self.phonenumber = "N/A"


        address_row = basic_info_table_row[12].find('td').get_text()
        # 우편번호 Handle
        self.postcode = address_row.strip().split("\n")[0].replace("(", "").replace(")", "").zfill(5)

        #주소 Handle
        self.address_kor = address_row.strip().split("\n")[1].strip()

        turnovertable = soup.find_all("tbody")[1]  # 매출액테이블 선택

        moneyandeieik_list = self.get_moneydict(turnovertable)

        moneydict = moneyandeieik_list[0]
        if moneydict == "N/A":
            moneydict = {}
            moneydict[2015], moneydict[2016], moneydict[2017] = "N/A", "N/A", "N/A"
            moneydict['lastyear'] = "N/A"
        else:
            moneydictyears = sorted(list(moneydict.keys()))
            lastyear = moneydictyears[-1]
            moneydict['lastyear'] = moneydict[lastyear]
            if lastyear < 2015:
                moneydict[2015] = str(lastyear) + "년   " + str(moneydict[lastyear])
                moneydict[2015], moneydict[2016], moneydict[2017] = "N/A", "N/A", "N/A"
            else:
                for i in [2015, 2016, 2017]:
                    if i not in moneydictyears:
                        moneydict[i] = "N/A"
        self.moneydict = moneydict

        eieikdict = moneyandeieik_list[1]
        if eieikdict == "N/A":
            eieikdict = {}
            eieikdict[2015], eieikdict[2016], eieikdict[2017] = "N/A", "N/A", "N/A"
        else:
            eieikdictyears = sorted(list(eieikdict.keys()))
            lastyear = eieikdictyears[-1]
            if lastyear < 2015:
                eieikdict[2015] = str(lastyear) + "년   " + str(eieikdict[lastyear])
                eieikdict[2015], eieikdict[2016], eieikdict[2017] = "N/A", "N/A", "N/A"
            else:
                for i in [2015, 2016, 2017]:
                    if i not in eieikdictyears:
                        eieikdict[i] = "N/A"
        self.eieikdict = eieikdict


        sokdict = moneyandeieik_list[2]
        if sokdict == "N/A":
            sokdict = {}
            sokdict[2015], sokdict[2016], sokdict[2017] = "N/A", "N/A", "N/A"
        else:
            sodictyears = sorted(list(sokdict.keys()))
            lastyear = sodictyears[-1]
            if lastyear < 2015:
                sokdict[2015] = str(lastyear) + "년   " + str(sokdict[lastyear])
                sokdict[2015], sokdict[2016], sokdict[2017] = "N/A", "N/A", "N/A"
            else:
                for i in [2015, 2016, 2017]:
                    if i not in sodictyears:
                        sokdict[i] = "N/A"
        self.sokdict = sokdict

        pskdict = moneyandeieik_list[3]
        if pskdict == "N/A":
            pskdict = {}
            pskdict[2015], pskdict[2016], pskdict[2017] = "N/A", "N/A", "N/A"
        else:
            psdictyears = sorted(list(pskdict.keys()))
            lastyear = psdictyears[-1]
            if lastyear < 2015:
                pskdict[2015] = str(lastyear) + "년   " + str(pskdict[lastyear])
                pskdict[2015], pskdict[2016], pskdict[2017] = "N/A", "N/A", "N/A"
            else:
                for i in [2015, 2016, 2017]:
                    if i not in psdictyears:
                        pskdict[i] = "N/A"
        self.pskdict = pskdict
        
        sooneieikdict = moneyandeieik_list[4]
        if sooneieikdict == "N/A":
            sooneieikdict = {}
            sooneieikdict[2015], sooneieikdict[2016], sooneieikdict[2017] = "N/A", "N/A", "N/A"
        else:
            sooneieikyears = sorted(list(sooneieikdict.keys()))
            lastyear = sooneieikyears[-1]
            if lastyear < 2015:
                sooneieikdict[2015] = str(lastyear) + "년   " + str(sooneieikdict[lastyear])
                sooneieikdict[2015], sooneieikdict[2016], sooneieikdict[2017] = "N/A", "N/A", "N/A"
            else:
                for i in [2015, 2016, 2017]:
                    if i not in sooneieikyears:
                        sooneieikdict[i] = "N/A"
        self.sooneieikdict = sooneieikdict


    def get_moneydict(self, turnovertable):
        turnover_row = turnovertable.find_all("tr")
        moneydict = OrderedDict()
        eieikdict = OrderedDict()
        sokdict = OrderedDict()
        pskdict = OrderedDict()
        sooneieikdict = OrderedDict()
        number = 0
        firstyear = turnover_row[0].find_all("td")[0].get_text()[0:4]
        if firstyear == "조회된 ":
            moneydict = "N/A"
            eieikdict = "N/A"
            sokdict = "N/A"
            pskdict = "N/A"
            sooneieikdict = "N/A"
        else:
            for i in turnover_row:
                year = int(i.find_all("td")[0].get_text()[0:4])
                try:
                    moneyofyear = int(i.find_all("td")[4].get_text().replace(",", "")) * 1000
                    eieikofyear = int(i.find_all("td")[5].get_text().replace(",", "")) * 1000
                    sofyear = float(i.find_all("td")[6].get_text().strip().replace("%","")) 
                    psofyear = float(i.find_all("td")[7].get_text().strip().replace("%","")) 
                    
#밑에게 영업이익 eieikofyear
                except:
                    moneyofyear = "N/A"
                    eieikofyear = "N/A"
                    sofyear = "N/A"
                    psofyear = "N/A"
                finally:
                    moneydict[year] = moneyofyear
                    eieikdict[year] = eieikofyear
                    sokdict[year] = sofyear
                    pskdict[year] = psofyear


                try:
                    sooneieikofyear = float(i.find_all("td")[7].get_text().strip().replace("%", "").replace(",", "")) * moneyofyear
                    sooneieikdict[year] = sooneieikofyear
                except:
                    sooneieikofyear = "N/A"
                finally:
                    sooneieikdict[year] = sooneieikofyear

                number = number + 1
        moneyandeieik_list = [moneydict, eieikdict, sokdict, pskdict, sooneieikdict]   #변경함
        return moneyandeieik_list

    def get_notable_event(self, tables):
        try:
            rows = tables[2].find_all("tr")
            temp_list = []
            for i in rows:
                information = i.find_all("td")
                date = information[0].get_text()
                text = information[1].get_text()
                string = f'{date} : {text}'
                temp_list.append(string)
            self.notable_event = "\n".join(temp_list)

        except:
            self.notable_event = "N/A"




    def get_info_from_both_page(self):
        kor_soup = self.get_soup(self.kor_page_source)
        kor_tables = kor_soup.find_all("tbody")
        basic_info = kor_tables[0].find_all("tr")
        self.get_notable_event(kor_tables)
        eng_soup = self.get_soup(self.eng_page_source)
        eng_table = eng_soup.find("tbody")
        eng_basic_info = eng_table.find_all("tr")

        self.company_name_eng = eng_basic_info[0].get_text().strip().split("\n")[-1].strip()


        company_group = basic_info[5].get_text().strip().split("\n")

        if company_group[1] != '-':
            self.kor_group_main_company = company_group[-1].split(")", 1)[-1]
            self.groupname_eng = eng_basic_info[13].get_text().strip().split("\n")[-1].split(")", 1)[-1]
        else:
            self.kor_group_main_company = "N/A"
            self.groupname_eng = "N/A"



        address_eng = eng_basic_info[5].get_text().replace("MAP Land-lot Num.", "").strip().split("\n")[-1].split(" ")
        if len(address_eng) == 1:
            address_eng = []
            temp_address_eng = eng_basic_info[5].get_text().replace("MAP Land-lot Num.", "").strip().split("\n")[
                -1].split(",")
            for i in temp_address_eng:
                address_eng.append(i.capitalize())

        lastadd = "".join(address_eng[-1])
        try:
            if lastadd in ["Busan", "Seoul", "Ulsan", "Incheon", "Daegu", "Daejeon", "Gwangju", 'Sejong']:
                self.eng_add_1 = " ".join(address_eng)
                self.eng_add_2 = "".join(address_eng[-1])
                self.eng_add_3 = "".join(address_eng[-1])
            else:
                self.eng_add_1 = " ".join(address_eng)
                self.eng_add_2 = "".join(address_eng[-2])
                self.eng_add_3 = "".join(address_eng[-1])

            if lastadd in ["Gyeongbuk", "Busan", "Daegu", "Ulsan", "Gyeongnam"]:
                self.youngnam = "yes"
            else:
                self.youngnam = 'no'
        except:
            print("Someting is wrong with address. Please check it later")
            pass

        faxnumber = eng_basic_info[6].get_text().split("\n")[1].split(":")[2].strip()
        try:
            if len(faxnumber) == 0:
                raise
            else:
                self.fax_number = faxnumber
        except:
            self.fax_number = "N/A"


        homepage = basic_info[15].get_text().split("홈페이지")[-1].strip()
        if homepage == '-' or homepage == 'http://':
            self.homepage = "N/A"
        else:
            self.homepage = homepage

        self.ceo_name_kor = basic_info[1].get_text().strip().split("\n ")[1].strip()
        ceo_name_eng = eng_basic_info[1].get_text().strip().split("\n")[1].split("/")[0]
        if "," in ceo_name_eng:  # 한국인인경우
            self.ceo_name_eng_last = ceo_name_eng.split(",")[0].strip()
            self.ceo_name_eng_first = ceo_name_eng.split(",")[1].strip()
        else:  # 외국인인경우
            self.ceo_name_eng_last = ceo_name_eng.split(" ")[-1].strip()
            self.ceo_name_eng_first = " ".join(ceo_name_eng.split(" ")[0:-1]).strip()

        self.kscode_eng = eng_basic_info[11].get_text().split("\n")[-2].split(")")[-1]
        self.notable_information = basic_info[17].get_text().strip()

    def get_fitst_click(self):
        firstclick = self.driver.find_element_by_class_name("lnb").find_elements_by_tag_name("ul")[2]
        firstclick.click()
        whattoclick = firstclick.find_element_by_tag_name("ul").find_elements_by_tag_name("li")
        return whattoclick

    def analyze_group_table(self):
        self.get_fitst_click()[1].click()
        group_page = self.driver.page_source
        group_soup = self.get_soup(group_page)
        thead = group_soup.find("thead").find_all("tr")[0].find("th").get_text().strip()

        if thead == "주주명":
            tabletype = 6
        else:
            tabletype = 8

        table = group_soup.find("tbody")
        rowlist = table.find_all("tr")
        temp_list = []

        if tabletype == 8:
            for rows in range(3):
                groupdict = {}
                try:
                    findtd = rowlist[rows * 2].find_all("td")
                    groupdict["name"] = findtd[2].get_text().strip()
                    groupdict["status"] = findtd[6].get_text().strip()
                    groupdict["portion"] = float(findtd[5].get_text().replace("%", "").strip())
                    temp_list.append(groupdict)
                except:
                    groupdict["name"], groupdict["portion"], groupdict["status"] = np.nan, np.nan, np.nan
                    temp_list.append(groupdict)

        elif tabletype == 6:
            for rows in range(3):
                groupdict = {}
                try:
                    findtd = rowlist[rows * 2].find_all("td")
                    groupdict["name"] = findtd[0].get_text().strip()
                    groupdict["status"] = findtd[1].get_text().strip()
                    groupdict["portion"] = float(findtd[6].get_text().replace("%", "").strip())
                    temp_list.append(groupdict)
                except:
                    groupdict["name"], groupdict["portion"], groupdict["status"] = np.nan, np.nan, np.nan
                    temp_list.append(groupdict)
        self.juju = temp_list


    def analyze_employee_tab(self):
        driver.find_element_by_xpath('//*[@id="content"]/div/ul[3]/li/ul/li[3]/a').click()
        driver.find_element_by_xpath('//*[@id="cont"]/ul/li[4]/a').click()

        employee_page = self.driver.page_source
        employee_soup = self.get_soup(employee_page)

        try:
            employeetable = employee_soup.find("div", id="empPrsIfrs").find("tbody")
            typeofemployee = "ifrs"
        except:
            employeetable = employee_soup.find("div", id="empPrsGaap").find("tbody")
            typeofemployee = "gaap"

        # 사업명구하기
        line_of_business =[]
        for i in employeetable.find_all("th"):
            business = i.get_text()
            line_of_business.append(business)

        employrows = employeetable.find_all("tr")
        employee_dict = {}
        try:
            if typeofemployee == "ifrs":
                temp_number_space = 0
                for index, row in enumerate(employrows):
                    numberofemployee = int(row.find_all("td")[5].get_text())
                    if index % 2 == 0:
                        temp_number_space = 0
                        temp_number_space = temp_number_space + numberofemployee
                    else:
                        temp_number_space = temp_number_space + numberofemployee
                        order_of_business = index // 2
                        business_name = line_of_business[order_of_business]
                        employee_dict[business_name] = temp_number_space
            else:
                for index, row in enumerate(employrows):
                    classofemployee = row.find("th").get_text()
                    numberofemployee = int(row.find("td").get_text())
                    employee_dict[classofemployee] = numberofemployee

        except:
            employee_dict = 'N/A'
        self.employee_dict = str(employee_dict).replace("{","").replace("}", "").replace("'","")

    def analyze_business_tab(self):
        driver.find_element_by_xpath('//*[@id="content"]/div/ul[3]/li/ul/li[5]/a').click()

        business_page = self.driver.page_source
        business_soup = self.get_soup(business_page)

        lineofbusiness = business_soup.find("div", class_="tbl_type03").find("tbody").find_all("tr")
        business_text_list = []
        try:
            for i in lineofbusiness:
                itemofbusiness = i.find_all("td")[1].get_text()
                business_text_list.append(itemofbusiness)
            text = "\n".join(business_text_list)
            self.businessline = text
        except:
            self.businessline = "N/A"

        driver.find_element_by_xpath('//*[@id="cont"]/ul/li[3]/a').click()
        gongjang_page = self.driver.page_source
        gongjang_soup = self.get_soup(gongjang_page)
        gongjang_list = gongjang_soup.find("div", class_="tbl_type03").find("tbody").find_all("tr")
        whole_list = []
        add_type_dict = {}
        for i in gongjang_list:
            row_infomation = i.find_all("td")
            for index, information in enumerate(row_infomation):
                if index == 1:
                    name = information.get_text().strip()
                elif index == 2:
                    address = information.get_text().strip().replace("\n","").replace("지도보기","").split("지번주소")[0]
                elif index == 5:
                    phone = information.get_text().strip()
                elif index == 6:
                    addtype = information.get_text().strip()
                    if addtype in list(add_type_dict.keys()):
                        add_type_dict[addtype] = add_type_dict[addtype] + 1
                    else:
                        add_type_dict[addtype] = 1
            convertstring = f'{name} : {address}, phone: {phone}'
            whole_list.append(convertstring)
        addcount = str(add_type_dict).replace('{', "").replace('}', "").replace("'","")
        whole_address = "\n".join(whole_list)
        result_str = addcount + '\n' + whole_address
        self.whole_business_address = result_str

    def analyze_thankyou_tab(self):
        driver.find_element_by_xpath('//*[@id="content"]/div[1]/ul[3]/li/ul/li[12]/a').click()

        thankyou_soup = self.get_soup(self.driver.page_source)
        tablerows = thankyou_soup.find("div", "tbl_type03").find("tbody").find_all("tr")
        bupin = []
        try:
            for rows in tablerows:
                head = rows.find_all("td")
                for index, columns in enumerate(head):
                    information = columns.get_text()
                    if index == 0:
                        season = information
                    elif index == 2:
                        thankyoubupin = information
                bupin.append(thankyoubupin)
                break
            thankyoutext = bupin
        except:
            thankyoutext = 'N/A'

        self.thankyou = thankyoutext


    def analyze_boss(self):
        try:
            driver.find_element_by_xpath('//*[@id="content"]/div/ul[3]/li/ul/li[3]/a').click()
            driver.find_element_by_name("peDtlSrch").click()
        except:
            self.bossbirth = 'N/A'
            self.bossname = 'N/A'
            self.boss_information = 'N/A'
            return

        boss_soup = self.get_soup(self.driver.page_source)

        boss_rows = boss_soup.find("div", class_="tbl_type02").find("tbody").find_all("tr")
        textlist = []
        for row in boss_rows:
            header = row.find_all("th", scope="row")
            values = row.find_all("td")
            for lable, value in zip(header, values):
                lable = lable.get_text()
                value = value.get_text().replace("\n", "").strip()
                if lable != "근무지":
                    if lable == "성명":
                        self.bossname = value
                    if lable == "생년월일":
                        self.bossbirth = value
                    if lable == "대학교" or lable == "전공" or lable == "대학원 이상":
                        string = f'{lable} : {value}'
                        textlist.append(string)
                else:
                    break
        text = "\n".join(textlist)
        self.boss_information = text
        # except:
        #     print("cannot get any detailed information")





class DataFrameOperator:
    def __init__(self, muinfo):
        self.dir = os.path.dirname(os.path.abspath(__file__))
        self.prepare_tax_excel()
        self.prepare_kscode_excel()
        self.prepare_AE_excel()
        self.check_path()
        self.get_today()
        self.make_dataframe()
        self.mu = muinfo['MU']
        self.countrycode = muinfo['Countrycode']
        self.inumber = muinfo['inumber']

    def prepare_tax_excel(self):
        path = self.dir + r'\\asset\\static\\tax_finder.xlsx'
        taxfile = pd.read_excel(path)
        taxfile["Tax Number"] = taxfile["Tax Number"].apply(lambda x: str(x)).fillna('-9999')
        self.taxfile = taxfile

    def compare_tax(self, taxnumber):
        if taxnumber == "N/A":
            return
        elif np.isnan(taxnumber) == True:
            return
        else:
            taxdict = self.taxfile[self.taxfile["Tax Number"] == f'{taxnumber}'].reset_index().T.to_dict()
            try:
                taxlist = []
                for i in range(len(taxdict)):
                    print(f"동일한 사업자번호를 가진 BP#  {taxdict[i]['Business Partner']} | {taxdict[i]['Eng_name']} | {taxdict[i]['korea_name']} 가 발견되었습니다.")
                    taxlist.append(taxdict[i]['Business Partner'])
                if len(taxlist) == 0:
                    return ''
                else:
                    return str(taxlist).replace("[","").replace("]","")
            except:
                pass

    def prepare_kscode_excel(self):
        path = self.dir + r'\\asset\\static\\ks_code_mapper.xlsx'
        kscodefile = pd.read_excel(path)
        self.kscodefile = kscodefile

    def prepare_AE_excel(self):
        path = self.dir + r'\\asset\\static\\mc_ae_mapper.xlsx'
        mcaefile = pd.read_excel(path)
        mcaefile["MC_Code"] = mcaefile["MC_Code"].apply(str)
        self.mcaefile = mcaefile

    def match_kscode(self, kscode):
        if kscode == "N/A":
            return np.nan
        else:
            try:
                mccode = self.kscodefile[self.kscodefile["KS_Code(9th)"] == f'{kscode}']["MC_Code"].get_values()[0]
                mccode = str(mccode).replace(".0", "")
                return mccode
            except:
                return np.nan

    def match_ae(self, mccode, revenue, youngnam):

        try:
            if revenue >= 325000000000:
                segment = 'GBU'
                IAC = 'GB - Upper'
            elif revenue >= 100000000000:
                segment = 'Upper'
                IAC = 'GB - Lower'
            else:
                segment = 'Lower'
                IAC = 'GB - Lower'

        except:
            segment = 'Lower'
            IAC = 'GB - Lower'

        try:
            aelist = self.mcaefile[(self.mcaefile["MC_Code"] == f'{mccode}') &
                                   (self.mcaefile["Upper/Lower"] == f'{segment}')].reset_index().T.to_dict()[0]
            if youngnam == "yes":
                aelist['AE_inumber'], aelist['AE_LastName'], aelist['AE_FirstName'] = 'I345534','Hwang','Jongsu'
            aelist['IAC'] = IAC

        except:
            aelist = {'IAC' : IAC,
                      'MC_Code': np.nan,
                      'MC_Code_description': np.nan,
                      'AE_inumber': np.nan,
                      'AE_LastName': np.nan,
                      'AE_FirstName': np.nan}

        return aelist


    def make_dataframe(self):
        self.maindataframe = pd.DataFrame()

    def concat(self):
        self.maindataframe = pd.concat([self.maindataframe, self.adddataframe])

    def get_today(self):
        self.today = str(datetime.date.today())

    def check_path(self):
        try:
            if not (os.path.isdir("EXCEL")):
                os.makedirs(os.path.join("EXCEL"))
        except OSError as e:
            if e.errno != errno.EEXIST:
                print("Failed to create directory!")
            raise

    def get_df_row_size(self):
        return self.maindataframe.shape[0]

    def plush_to_excel(self, mode=0):
        try:
            now = datetime.datetime.today()
            nowstr = str(now)[:16].replace(":", "")
            rownum = self.get_df_row_size()
            file_basic = f'{nowstr}_({rownum})'
            if mode == 0:
                filename = f'{file_basic}.xlsx'
            elif mode == 1:
                filename = f'{file_basic}_autosave.xlsx'
            elif mode == 2:
                filename = f'{file_basic}_closesave.xlsx'
            elif mode == 3:
                filename = f'{file_basic}_errorsave.xlsx'
            elif mode == 4:
                filename = f'{file_basic}_finalsave.xlsx'
            elif mode == 10:
                filename = f'{file_basic}_socialsave.xlsx'

            filenamewithPath = ".\\EXCEL\\" + filename
            print("File is created at EXCEL folder as " + filename)
            self.maindataframe.to_excel(filenamewithPath, index=False)
        except:
            try:
                print("Attempt Once more")
                self.plush_to_excel(4)
            except KeyboardInterrupt:
                print("Attempt Final attempt")
                self.plush_to_excel(4)
        finally:
            if mode == 0:
                os.startfile(filenamewithPath)

    def make_infomation(self, information):
    
        basic_form = collections.OrderedDict({
            'Campaign Name': information.search_term,
            'BP #\n(Account ID)': str(information.taxlist),
            'DDE Mapping(KISLINE기준)' : '',
            'CRM기준_AE_NAME' : '',
            'KISLINE_AE_NAME' : str(information.AElast) + ', ' + str(information.AEfirst),
            'NNN여부(CRM)' : '',
            'AE Match(CRM-KISLINE)' : '',
            'Organisation Name - Kor': information.company_name_kor,
            'Organisation Name - Eng': information.company_name_eng,
            'KOSDAQ\n상장 여부': information.companytype,
            'Business registration, National/Tax ID \n(if DUNS not available)': information.taxnumber,
            'Master Code Text': information.mccode_discription,
            '9차_KS_Code\n(KISLINE data)': information.kscode,
            'Industry Description KOR\n(표준산업분류)\n(KISLINE data)': information.kscode_kor,
            'Industry Description ENG\n(표준산업분류)\n(KISLINE data)': information.kscode_eng,
            'Local Sales Turnover\n2015 (K won)': information.moneydict[2015],  #고친부분
            'Local Sales Turnover\n2016 (K won)': information.moneydict[2016],
            'Local Sales Turnover\n2017 (K won)': information.moneydict[2017],
            'Street Address\n(Korean)': information.address_kor,
            'Local Employee Size': information.number_of_employee,
            'Contact Person Korean Name': information.ceo_name_kor,
            'Org Telephone Number (no Country code indicator necessary)': information.phonenumber,
            'Website Address': information.homepage,
            'Master DB Contact 여부 (O,X)' : ''
        })
        self.adddataframe = pd.DataFrame(basic_form, index=[0])
        self.concat()

    def make_social_infomation(self, information):
        if len(information.taxlist) > 0:
            creation = "Yes"
            RBC = ""
        else:
            creation = "No"
            RBC = "Net New"
        basic_form = collections.OrderedDict({
            'Request Receive Date\ndd/mm/yyyy': self.today,
            'Account Creator(Name)': self.inumber,
            'Account Created' : creation,
            'Account ID (CRM BP#)' : str(information.taxlist),
            'Tax # (KISLINE)' : information.taxnumber,
            'Organisation Name - full legal version, no abbreviations or marketing or brand names\n(35 characters)': information.company_name_eng,
            'KOR_NAME(KISLINE)': information.company_name_kor,
            'Local Sales Turnover\n2017 (K won)': information.moneydict[2017],
            'AE Name(based on CRM)': '',
            'AE Name(based on KISLINE)' : str(information.AElast) + ', ' + str(information.AEfirst),
            'AE MATCH' : '',
            '2019 Internal Account Classification': information.IAC,
            'Regional Buying Classification' : RBC,
            'KISLINE KS CODE' : information.kscode,
            'KISLINE KS Descrtion': information.kscode_kor,
            'SIC MC Description - KISLINE': information.mccode_discription,            
            'SIC MC Description - CRM': '',
            'CRM MC_KIS MATCH' : '',
            'Website Address': information.homepage,
            'Org Telephone Number (no Country code indicator necessary)': information.phonenumber,            
            'Young Nam': information.youngnam,
            'Local Employee Size': information.number_of_employee,
            'Social Platform Lead (URL link)' : '',
            'Lead Description' : ''
        })
        self.adddataframe = pd.DataFrame(basic_form, index=[0])
        self.concat()


    def make_infomation_with_juju(self, information):
        #If you want with juju
        basic_form = collections.OrderedDict({
            'State of BP# Request': 'In Progress',
            'Campaign Name': information.search_term,
            'Requested Person': np.nan,
            'Request Receive Date\ndd/mm/yyyy': self.today,
            "BP Creation Requested Date\ndd/mm/yyyy": np.nan,
            "BP Creation Completed Date\ndd/mm/yyyy": np.nan,
            'MU': self.mu,
            'BP #\n(Account ID)': information.orgid,
            'Organisation Name - full legal version, no abbreviations or marketing or brand names\n(35 characters)': information.company_name_eng,
            'Organisation Name 2 - continued \n(35 characters)': information.company_name_kor,
            'PE Name_KOR\n(모기업)\n(KISLINE Data)': information.kor_group_main_company,
            'PE Name_ENG\n(KISLINE Data)': information.groupname_eng,
            'KOSDAQ\n상장 여부': information.companytype,
            'Street Address\n(Korean)': information.address_kor,
            'Street Address \n(35 characters)': information.eng_add_1,
            "Street Address 2 - continued (information appearing below 'Street Address'\n(35 characters)": np.nan,
            'City': information.eng_add_2,
            'PO Box': np.nan,
            'Postal code for PO Box': np.nan,
            'City of PO Box (if different to Street City)': np.nan,
            'State/Province/Territory': information.eng_add_3,
            'Young Nam province': information.youngnam,
            'Postal Code': information.postcode,
            'Country code ': self.countrycode,
            'Org Telephone Number (no Country code indicator necessary)': information.phonenumber,
            'Org Fax Number (no Country code indicator necessary)': information.fax_number,
            'Website Address': information.homepage,
            'DUNS Number\n(if provided in original data source)': '9999',
            'Pre BP Number\n': np.nan,
            'Business registration, National/Tax ID \n(if DUNS not available)': information.taxnumber,
            'Industry Code (SIC)\n(text name min.)': np.nan,
            'Master Code': information.mccode,
            'Master Code Text': information.mccode_discription,
            '9차_KS_Code\n(KISLINE data)': information.kscode,
            'Industry Description KOR\n(표준산업분류)\n(KISLINE data)': information.kscode_kor,
            'Industry Description ENG\n(표준산업분류)\n(KISLINE data)': information.kscode_eng,
            'Local Sales Turnover\n2016 (K won)': information.moneydict[2016],
            'Local Sales Turnover\n2017 (K won)': information.moneydict[2017],
            'EUR Sales Turnover 2016': np.nan,
            'Local Employee Size': information.number_of_employee,
            'Local Employee Size Year': information.year_of_employee,
            'Global Employee Size': np.nan,
            'AE_ID \n(I Number)': information.AEinumber,
            'AE_Last Name': information.AElast,
            'AE_First Name': information.AEfirst,
            'Proposed_2018_RBC': 'Net New',
            'Proposed_2018_IAC': information.IAC,
            'Contact Person \nSalutation ': np.nan,
            'Contact Person \nLast Name \n(40 characters)': information.ceo_name_eng_last,
            'Contact Person \nFirst Name \n(40 characters)': information.ceo_name_eng_first,
            'Contact Person Korean Name': information.ceo_name_kor,
            'Contact Email Address\n(40 characters)': np.nan,
            'E-mail Opt-In\n2018 \n(drop-down list)': np.nan,
            'Contact Person \nTelephone Number(no Country code indicator necessary)': information.phonenumber,
            'Contact Person\n Fax Number(no Country code indicator necessary)': np.nan,
            'Contact Person \nMobile Phone Number(no Country code indicator necessary)': np.nan,
            'Contact Person Linkedin \n(URL Link)': np.nan,
            'Miscellaneous Link': np.nan,
            'Job Title \n(free text-60 characters)': 'CEO',
            'Job Title\n(free text/Korean)': '대표이사',
            'Job Function(drop-down list)': 'Chief Exec Officer',
            'Department*\n(drop-down list)': 'Management',
            'Department*\n(free text/Korean)': np.nan,
            'Contact ID\n(from Data Team)': np.nan,
            'CCAP Team Member\n(I-Number)': self.inumber,
            'Company Status': information.companystatus,
            'Contact Status': 'New Contact Discovered',
            'Profile Status': 'Closed',
            'duplicated_tax': str(information.taxlist),
            '1st JUJU': information.juju[0]['name'],
            '1st status': information.juju[0]['status'],
            '1st Portion': information.juju[0]['portion'],
            '2nd JUJU': information.juju[1]['name'],
            '2nd status': information.juju[1]['status'],
            '2nd Portion': information.juju[1]['portion'],
            '3rd JUJU': information.juju[2]['name'],
            '3rd status': information.juju[2]['status'],
            '3rd Portion': information.juju[2]['portion']
        })
        self.adddataframe = pd.DataFrame(basic_form, index=[0])
        self.concat()

    def make_infomation_I_love_UGB(self, information):
        #If you want with UGB
        basic_form = collections.OrderedDict({
            'State of BP# Request': 'In Progress',
            'Campaign Name': information.search_term,
            'Requested Person': np.nan,
            'Request Receive Date\ndd/mm/yyyy': self.today,
            "BP Creation Requested Date\ndd/mm/yyyy": np.nan,
            "BP Creation Completed Date\ndd/mm/yyyy": np.nan,
            'MU': self.mu,
            'BP #\n(Account ID)': information.orgid,
            'Organisation Name - full legal version, no abbreviations or marketing or brand names\n(35 characters)': information.company_name_eng,
            'Organisation Name 2 - continued \n(35 characters)': information.company_name_kor,
            'PE Name_KOR\n(모기업)\n(KISLINE Data)': information.kor_group_main_company,
            'PE Name_ENG\n(KISLINE Data)': information.groupname_eng,
            'KOSDAQ\n상장 여부': information.companytype,
            'Street Address\n(Korean)': information.address_kor,
            'Street Address \n(35 characters)': information.eng_add_1,
            "Street Address 2 - continued (information appearing below 'Street Address'\n(35 characters)": np.nan,
            'City': information.eng_add_2,
            'PO Box': np.nan,
            'Postal code for PO Box': np.nan,
            'City of PO Box (if different to Street City)': np.nan,
            'State/Province/Territory': information.eng_add_3,
            'Young Nam province': information.youngnam,
            'Postal Code': information.postcode,
            'Country code ': self.countrycode,
            'Org Telephone Number (no Country code indicator necessary)': information.phonenumber,
            'Org Fax Number (no Country code indicator necessary)': information.fax_number,
            'Website Address': information.homepage,
            'DUNS Number\n(if provided in original data source)': '9999',
            'Pre BP Number\n': np.nan,
            'Business registration, National/Tax ID \n(if DUNS not available)': information.taxnumber,
            'Industry Code (SIC)\n(text name min.)': np.nan,
            'Master Code': information.mccode,
            'Master Code Text': information.mccode_discription,
            '9차_KS_Code\n(KISLINE data)': information.kscode,
            'Industry Description KOR\n(표준산업분류)\n(KISLINE data)': information.kscode_kor,
            'Industry Description ENG\n(표준산업분류)\n(KISLINE data)': information.kscode_eng,
            'Local Sales Turnover\n2016 (K won)': information.moneydict[2015],
            'Local Sales Turnover\n2016 (K won)': information.moneydict[2016],
            'Local Sales Turnover\n2017 (K won)': information.moneydict[2017],
            'EUR Sales Turnover 2016': np.nan,
            'Local Employee Size': information.number_of_employee,
            'Local Employee Size Year': information.year_of_employee,
            'Global Employee Size': np.nan,
            'AE_ID \n(I Number)': information.AEinumber,
            'AE_Last Name': information.AElast,
            'AE_First Name': information.AEfirst,
            'Proposed_2018_RBC': 'Net New',
            'Proposed_2018_IAC': information.IAC,
            'Contact Person \nSalutation ': np.nan,
            'Contact Person \nLast Name \n(40 characters)': information.ceo_name_eng_last,
            'Contact Person \nFirst Name \n(40 characters)': information.ceo_name_eng_first,
            'Contact Person Korean Name': information.ceo_name_kor,
            'Contact Email Address\n(40 characters)': np.nan,
            'E-mail Opt-In\n2018 \n(drop-down list)': np.nan,
            'Contact Person \nTelephone Number(no Country code indicator necessary)': information.phonenumber,
            'Contact Person\n Fax Number(no Country code indicator necessary)': np.nan,
            'Contact Person \nMobile Phone Number(no Country code indicator necessary)': np.nan,
            'Contact Person Linkedin \n(URL Link)': np.nan,
            'Miscellaneous Link': np.nan,
            'Job Title \n(free text-60 characters)': 'CEO',
            'Job Title\n(free text/Korean)': '대표이사',
            'Job Function(drop-down list)': 'Chief Exec Officer',
            'Department*\n(drop-down list)': 'Management',
            'Department*\n(free text/Korean)': np.nan,
            'Contact ID\n(from Data Team)': np.nan,
            'CCAP Team Member\n(I-Number)': self.inumber,
            'Company Status': information.companystatus,
            'Contact Status': 'New Contact Discovered',
            'Profile Status': 'Closed',
            'duplicated_tax': str(information.taxlist),
            '1st JUJU': information.juju[0]['name'],
            '1st status': information.juju[0]['status'],
            '1st Portion': information.juju[0]['portion'],
            '2nd JUJU': information.juju[1]['name'],
            '2nd status': information.juju[1]['status'],
            '2nd Portion': information.juju[1]['portion'],
            '3rd JUJU': information.juju[2]['name'],
            '3rd status': information.juju[2]['status'],
            '3rd Portion': information.juju[2]['portion'],
            'Turnover2015 (K won)': information.moneydict[2015],
            'OperationProfiturnover2015 (K won)': information.eieikdict[2015],
            'NetProfit2015 (K won)': information.sooneieikdict[2015],
            'Turnover2016 (K won)': information.moneydict[2016],
            'OperationProfiturnover2016 (K won)': information.eieikdict[2016],
            'NetProfit2016 (K won)': information.sooneieikdict[2016],
            'Turnover2017 (K won)': information.moneydict[2017],
            'OperationProfiturnover2017 (K won)': information.eieikdict[2017],
            'NetProfit2017 (K won)': information.sooneieikdict[2017],
            'main_item' : information.items,
            'employstatus' : information.employee_dict,
            'line_of_business' : information.businessline,
            'business_address' : information.whole_business_address,
            'kamsabupin' : information.thankyou,
            'boss_name' : information.bossname,
            'boss_birth' : information.bossbirth,
            'boss_information' : information.boss_information,
            'deepsearch' : information.deepsearch,
            'additional_information' : information.notable_information,
            'notable_event' : information.notable_event
        })
        self.adddataframe = pd.DataFrame(basic_form, index=[0])
        self.concat()


def massive_search():
    #ready for massive_search.
    path = '.\\asset\\static\\massive.xlsx'
    searchfile = pd.read_excel(path)
    search_data = searchfile.columns[0]
    org_data = searchfile.columns[1]
    searchfile[search_data] = searchfile[search_data].apply(lambda x: str(x).replace(".0", ""))
    searchfile[org_data] = searchfile[org_data].fillna(-9999).apply(lambda x: str(x).replace(".0", ""))
    info_additional = []
    info_len = len(searchfile.columns)
    for i in range(info_len):
        if i == 0:
            pass
        else:
            if searchfile[searchfile.columns[i]].unique().shape[0] != 1:
                info_additional.append(searchfile.columns[i])
    output_additional_info_list = []
    for i in info_additional:
        output_additional_info_list.append(searchfile[i].fillna(-9999).tolist())

    return searchfile[search_data].tolist(), searchfile[org_data].tolist(), output_additional_info_list


if __name__ == "__main__":
    seldriver = SelDriver()
    driver = seldriver.control()
    muinfo = seldriver.get_MU_info()
    maindataframe = DataFrameOperator(muinfo)
    print("**Main Dataframe initiated.")

    mode = True
    print("Please Put your 'SEARCH TERM', SEARCH TERM may Company's name or Tax number      : ")
    print("!!!!!!!If you want massive search: Press <<<<9999<<<<<<<<<, Social Listening: 1212")
    search_term = ""
    while search_term == "":
        search_term = input("SEARCH TERM:::::>>>>>>>>              ")

    if search_term == str(9999):
        mode = 9999
        print("""
        !*!*!**!!*!*!*!*!*!!*!*!*!
        WELCOMETOMASSIVESEARCHMODE
        !*!*!**!!*!*!*!*!*!!*!*!*!
                """)
        print("If you want to stop during loop, just close the brower. Python will automactically save the result")
        massivemode = None
        while (massivemode != "0") and (massivemode != "1") and (massivemode != "88"):
            massivemode = input("Please select massive search mode.        0: normal massive / 1:juju massive / 88:UGB  ")

        basic_searchlist, org_id, additioanl_info_list = massive_search()
        try:
            start_num = int(input("몇부터 시작하시겠습니까? 숫자만 가능. 시작 : 0 입력"))
        except ValueError:
            traceback.print_exception(*sys.exc_info())
            print("Once more!")
            start_num = int(input("몇부터 시작하시겠습니까? 숫자만 가능. 시작 : 0 입력"))
        searchlist = basic_searchlist[start_num:]
        org_id_list = org_id[start_num:]
        for i in range(len(additioanl_info_list)):
            additioanl_info_list[i] = additioanl_info_list[i][start_num:]

        while mode != 0:
            try:
                for i in range(len(searchlist)):
                    search_object = [searchlist[i], org_id_list[i]]
                    print(searchlist[i])
                    for info_list in additioanl_info_list:
                        if info_list[i] == -9999:
                            pass
                        else:
                            print(info_list[i])

                    if massivemode == "0":
                        information = GiveMeInfo(driver, search_object).solo()
                        information.taxlist = maindataframe.compare_tax(information.taxnumber)
                        mccode = maindataframe.match_kscode(information.kscode)
                        aelist = maindataframe.match_ae(mccode, information.moneydict[2017], information.youngnam)
                        information.mccode, information.mccode_discription, information.IAC = aelist['MC_Code'], aelist[
                            'MC_Code_description'], aelist['IAC']
                        information.AEinumber, information.AElast, information.AEfirst = aelist['AE_inumber'], aelist[
                            'AE_LastName'], aelist['AE_FirstName']

                        maindataframe.make_infomation(information)
                        del information

                    elif massivemode == "1":
                        information = GiveMeInfo(driver, search_object).juju_club()
                        information.taxlist = maindataframe.compare_tax(information.taxnumber)
                        mccode = maindataframe.match_kscode(information.kscode)
                        aelist = maindataframe.match_ae(mccode, information.moneydict[2017], information.youngnam)
                        information.mccode, information.mccode_discription, information.IAC = aelist['MC_Code'], aelist[
                            'MC_Code_description'], aelist['IAC']
                        information.AEinumber, information.AElast, information.AEfirst = aelist['AE_inumber'], aelist[
                            'AE_LastName'], aelist['AE_FirstName']

                        maindataframe.make_infomation_with_juju(information)
                        del information

                    elif massivemode == "88":
                        information = GiveMeInfo(driver, search_object).i_love_ugb()
                        information.taxlist = maindataframe.compare_tax(information.taxnumber)
                        mccode = maindataframe.match_kscode(information.kscode)
                        aelist = maindataframe.match_ae(mccode, information.moneydict[2017], information.youngnam)
                        information.mccode, information.mccode_discription, information.IAC = aelist['MC_Code'], aelist[
                            'MC_Code_description'], aelist['IAC']
                        information.AEinumber, information.AElast, information.AEfirst = aelist['AE_inumber'], aelist[
                            'AE_LastName'], aelist['AE_FirstName']

                        maindataframe.make_infomation_I_love_UGB(information)
                        del information

                    number = start_num + i
                    print(f"{number}번째 작업완료. EXCEL상 {number+2}번째 계정이 완료되었습니다. 재시작시 {number+1}을 입력해주세요.")
                    time.sleep(0.4)

                    if (i%5 == 0) & (i != 0):
                        print("잠시 10초 쉬어갑니다. ")
                        time.sleep(10)

                    elif (i%100 == 0) & (i != 0):
                        print("For ensure your data, Autosave is performed")
                        maindataframe.plush_to_excel(1)

                        print("Don't worry, I'm only swimming. Let's begin again! :)")
                mode = 0
                print("Completed!")
                print("Be Happy!!!!")

            except Exception as e:
                maindataframe.plush_to_excel(3)
                print(e)
                traceback.print_exception(*sys.exc_info())
                print('Error is occured. Please report to wunhan.park@sap.com')
                break

            except KeyboardInterrupt:
                try:
                    print('You must want to close this program. Okay... bye...★')
                    maindataframe.plush_to_excel(2)
                    traceback.print_exception(*sys.exc_info())
                    break

                except EOFError:
                    maindataframe.plush_to_excel(2)
                    break

                except KeyboardInterrupt:
                    maindataframe.plush_to_excel(2)
                    break
        else:
            print("Done! NOW EXPORTING YOUR DATA......")
            maindataframe.plush_to_excel()

    elif search_term == str(1212):
        submaindataframe = DataFrameOperator(muinfo)
        print("""!**!*!*!*!*!*!*
        Social listening Mode!
              !**!*!**!*!*!!""")
        search_term = ""
        while search_term == "":
            search_term = input("SEARCH TERM:::::>>>>>>>>              ")

        while (mode != "0") & (mode != 9999):
            try:
                information = GiveMeInfo(driver, search_term).solo()
                information.taxlist = maindataframe.compare_tax(information.taxnumber)
                mccode = maindataframe.match_kscode(information.kscode)
                aelist = maindataframe.match_ae(mccode, information.moneydict[2017], information.youngnam)
                information.mccode, information.mccode_discription, information.IAC = aelist['MC_Code'], aelist[
                    'MC_Code_description'], aelist['IAC']
                information.AEinumber, information.AElast, information.AEfirst = aelist['AE_inumber'], aelist[
                    'AE_LastName'], aelist['AE_FirstName']

                maindataframe.make_infomation(information)
                submaindataframe.make_social_infomation(information)

                del information

                print("If you want to continue, please put other SEARCH TERM.         ")
                search_term = ''
                while search_term == '':
                    search_term = str(input("ELSE, EXPORT TO EXCEL >:>:>:>:>:>       PRESS 0                    "))
                if search_term == "0":
                    mode = "0"
                else:
                    mode = 99

            except Exception as e:
                maindataframe.plush_to_excel(3)
                submaindataframe.plush_to_excel(10)
                print(e)
                traceback.print_exception(*sys.exc_info())
                print('Error is occured. Please report to wunhan.park@sap.com')
                break

            except KeyboardInterrupt:
                try:
                    print('You must want to close this program. Okay... bye...★')
                    maindataframe.plush_to_excel(2)
                    submaindataframe.plush_to_excel(10)
                    break

                except EOFError:
                    maindataframe.plush_to_excel(2)
                    submaindataframe.plush_to_excel(10)
                    break

                except KeyboardInterrupt:
                    maindataframe.plush_to_excel(2)
                    submaindataframe.plush_to_excel(10)
                    break

        else:
            print("Completed!")
            print("NOW EXPORTING YOUR DATA......")
            maindataframe.plush_to_excel()
            submaindataframe.plush_to_excel(10)
            print("Be Happy!!!!")

    else:
        while (mode != "0") & (mode != 9999):
            try:
                information = GiveMeInfo(driver, search_term).solo()
                information.taxlist = maindataframe.compare_tax(information.taxnumber)
                mccode = maindataframe.match_kscode(information.kscode)
                aelist = maindataframe.match_ae(mccode, information.moneydict[2017], information.youngnam)
                information.mccode, information.mccode_discription, information.IAC = aelist['MC_Code'], aelist['MC_Code_description'], aelist['IAC']
                information.AEinumber, information.AElast, information.AEfirst = aelist['AE_inumber'], aelist[
                    'AE_LastName'], aelist['AE_FirstName']

                maindataframe.make_infomation(information)
                del information

                print("If you want to continue, please put other SEARCH TERM.         ")
                search_term = ''
                while search_term == '':
                    search_term = str(input("ELSE, EXPORT TO EXCEL >:>:>:>:>:>       PRESS 0                    "))
                if search_term == "0":
                    mode = "0"
                else:
                    mode = 99


            except Exception as e:
                maindataframe.plush_to_excel(3)
                print(e)
                traceback.print_exception(*sys.exc_info())
                print('Error is occured. Please report to wunhan.park@sap.com')
                break

            except KeyboardInterrupt:
                try:
                    print('You must want to close this program. Okay... bye...★')
                    maindataframe.plush_to_excel(2)
                    break

                except EOFError:
                    maindataframe.plush_to_excel(2)
                    break

                except KeyboardInterrupt:
                    maindataframe.plush_to_excel(2)
                    break

        else:
            print("Completed!")
            print("NOW EXPORTING YOUR DATA......")
            maindataframe.plush_to_excel()
            print("Be Happy!!!!")
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.touch_actions import TouchActions
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from xlwt import Workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from fake_useragent import UserAgent
import smtplib, ssl
import pandas





def configure_driver():


    # Add additional Options to the webdriver
    chrome_options = Options()
    ua = UserAgent()
    userAgent = ua.random                                     #THIS IS FAKE AGENT IT WILL GIVE YOU NEW AGENT EVERYTIME
    print(userAgent)
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-infobars")
    #chrome_options.add_argument("start-maximized")
    #chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument("--disable-extensions")
    #chrome_options.add_argument('--proxy-server=%s' % PROXY)
    #chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
    return driver

def emailtome():
    port = 587  # For starttls
    smtp_server = "smtp.gmail.com"
    sender_email = "propertyownershipscrapper"
    receiver_email = "asadikram53@gmail.com"
    password = "property@123"
    message = """\
    Subject: Solve The Captha

    Hey get to your system and kindly solve the captcha for code to continue Thank!."""

    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_server, port) as server:
        server.ehlo()  # Can be omitted
        server.starttls(context=context)
        server.ehlo()  # Can be omitted
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message)


def getCourses(driver, search_keyword):

    start_time = time.time()


    kb = Workbook()
    # add_sheet is used to create sheet.
    sheet1 = kb.add_sheet('Sheet 1', cell_overwrite_ok=True)
    print(" WORKSHEET CREATED SUCCESSFULLY!")
    print(" ")
    print(" ")
    print(" ")
    # INITIALIZING THE COLOUMN NAMES NOW
    sheet1.write(0, 0, "LINK")
    sheet1.write(0, 1, "Id. f??d??ral de b??timent (EGID)")
    sheet1.write(0, 2, "Abr??viation du canton")
    sheet1.write(0, 3, "N?? OFS de la commune")
    sheet1.write(0, 4, "Nom le la commune")
    sheet1.write(0, 5, "Id. f??d??ral d'immeuble (EGRID)")
    sheet1.write(0, 6, "N?? de secteur du RF")
    sheet1.write(0, 7, "N?? d'immeuble")
    sheet1.write(0, 8, "Suffixe du n?? d'immeuble")
    sheet1.write(0, 9, "Type d'immeuble")
    sheet1.write(0, 10, "N?? officiel de b??timent")
    sheet1.write(0, 11, "Nom du b??timent")
    sheet1.write(0, 12, "Coordonn??e E du b??timent")
    sheet1.write(0, 13, "Coordonn??e N du b??timent")
    sheet1.write(0, 14, "Provenance des coordonn??es")
    sheet1.write(0, 15, "Statut du b??timent")
    sheet1.write(0, 16, "Cat??gorie de b??timent")
    sheet1.write(0, 17, "Classe de b??timent")
    sheet1.write(0, 18, "Ann??e de construction du b??timent")
    sheet1.write(0, 19, "Mois de construction du b??timent")
    sheet1.write(0, 20, "Epoque de construction")
    sheet1.write(0, 21, "Ann??e de d??molition du b??timent")
    sheet1.write(0, 22, "Surface du b??timent [m2]")
    sheet1.write(0, 23, "Volume du b??timent [m3]")
    sheet1.write(0, 24, "Volume du b??timent : norme")
    sheet1.write(0, 25, "Volume du b??timent : indication sur la donn??e")
    sheet1.write(0, 26, "Nombre de niveaux")
    heading1 = "Nombre d'" + 'enregistrements "logements"'
    sheet1.write(0, 27, heading1)
    sheet1.write(0, 28, "Nombre de pi??ces d???hab. ind??p.")
    sheet1.write(0, 29, "Etat des donn??es")
    sheet1.write(0, 30, "Id. f??d??ral d???entr??e (EDID)")
    sheet1.write(0, 31, "Id. f??d??ral d???adresse de b??timent (EGAID)")
    sheet1.write(0, 32, "N?? d???entr??e du b??timent")
    sheet1.write(0, 33, "Id. f??d??ral de rue (ESID)")
    sheet1.write(0, 34, "D??signation de la rue FR")
    sheet1.write(0, 35, "D??signation abr??g??e de la rue FR")
    sheet1.write(0, 36, "R??f??rence de l'index FR")
    sheet1.write(0, 37, "Langue du nom de rue FR")
    sheet1.write(0, 38, "D??signation officielle")
    sheet1.write(0, 39, "NPA")
    sheet1.write(0, 40, "Chiffre compl. du NPA")
    sheet1.write(0, 41, "Localit??")
    sheet1.write(0, 42, "Coordonn??e E de l'entr??e")
    sheet1.write(0, 43, "Coordonn??e N de l'entr??e")
    sheet1.write(0, 44, "Adresse officielle")
    sheet1.write(0, 45, "??tat des donn??es")
    sheet1.write(0, 46, "Link To Object")
    sheet1.write(0, 47, "Id. f??d??ral de logement (EWID)")
    sheet1.write(0, 48, "N?? administratif de logement")
    sheet1.write(0, 49, "Etage")
    sheet1.write(0, 50, "Logement sur plusieurs ??tages")
    sheet1.write(0, 51, "N?? physique de logement")
    sheet1.write(0, 52, "Situation sur l'??tage")
    sheet1.write(0, 53, "Statut du logement")
    sheet1.write(0, 54, "??tat des donn??es")
    sheet1.write(0, 55, "Rfinfo Link")
    sheet1.write(0, 56, "Name Step(2)")
    sheet1.write(0, 57, "LocalCh Link")
    sheet1.write(0, 58, "Officer Link")
    sheet1.write(0, 59, "Name Step(3)")
    sheet1.write(0, 60, "Address_local")
    sheet1.write(0, 61, "HouseNumber_local")
    sheet1.write(0, 62, "NPA_local")
    sheet1.write(0, 63, "Localit??_local")
    sheet1.write(0, 64, "Phone_numberlocal")
    sheet1.write(0, 65, "Data_LocalCh_Availaible")
    kb.save('propertyownership.xls')
    mi = 1
    #chosen = int(input("Enter The Number You Want To Start From 1-END: "))
    excel_data_df = pandas.read_csv('VIDI.csv')
    ALLCHE = excel_data_df['EGID'].tolist()
    print("Total Entries: ",len(ALLCHE))
    #ALLCHE =['280092392']
    lkno=1
    for ci in range(len(ALLCHE)):
        try:
            print("LINK NO: ",lkno)
            lkno=lkno+1
            choosenurl = ''
            try:
                choosenurl = 'https://api.geo.admin.ch/rest/services/ech/MapServer/ch.bfs.gebaeude_wohnungs_register/' + str(ALLCHE[ci]) + '_0/extendedHtmlPopup?lang=fr'
                driver.get(choosenurl)
                WebDriverWait(driver, 4).until(
                    expected_conditions.visibility_of_element_located((By.TAG_NAME, 'table')))
            except Exception:
                choosenurl = 'https://api.geo.admin.ch/rest/services/ech/MapServer/ch.bfs.gebaeude_wohnungs_register/' + str(ALLCHE[ci]) + '_1/extendedHtmlPopup?lang=fr'
                driver.get(choosenurl)
                WebDriverWait(driver, 4).until(expected_conditions.visibility_of_element_located((By.TAG_NAME, 'tr')))
            print("CHOSEN URL: ", choosenurl)
            WebDriverWait(driver, 40).until(expected_conditions.visibility_of_all_elements_located((By.TAG_NAME, 'tr')))
            datastateaccomodationinfo = ''
            housingstatus = ''
            situationonfloor = ''
            physicalhousingnumber = ''
            multistoryaccomodation = ''
            stage = ''
            administrativehousingnumber = ''
            federalhousingidewid = ''
            datastateentryinfo = ''
            linktowhite = ''
            officialaddress = ''
            ncordinateinput = ''
            ecordinateinput = ''
            locality = ''
            digitcomplpostcode = ''
            postcode = ''
            officialdesignation = ''
            languagestreetnamefr = ''
            indexreferencefr = ''
            shortdesignationfr = ''
            streetdesignationfr = ''
            federalstreetid = ''
            buildingentrynumber = ''
            federalbuildingaddressidegaid = ''
            federalentryidegid = ''
            datastate = ''
            numberofroomsofhabit = ''
            numberofaccommodationrecords = ''
            numberoflevels = ''
            buildingvolumeindicationofdata = ''
            buildingvolumestandard = ''
            buildingvolumem3 = ''
            buildingaream2 = ''
            yearofdemolitionofbuilding = ''
            constructionperiod = ''
            monthofbuildingconstruction = ''
            yearofconstructionofbuilding = ''
            buildingclass = ''
            buildingcategory = ''
            buildingstatus = ''
            sourceofcontactdetails = ''
            mcordinateofbuilding = ''
            ecordinateofbuilding = ''
            nameofbuilding = ''
            officialbuildingnumber = ''
            typeofbuilding = ''
            buildingnumbersuffix = ''
            buildingnumber = ''
            rfsectornumber = ''
            federalbuildingid = ''
            nameoftown = ''
            ofsnumberofmunicipality = ''
            cantonabbreviation = ''
            federalbuildingidegid = ''
            container = driver.find_elements_by_tag_name('tr')
            tabl = driver.find_element_by_tag_name('table')
            try:
                linktowhite = tabl.find_element_by_tag_name('a').get_attribute('href')
            except Exception:
                linktowhite = ''
                pass
            datastatecount = 1
            for contain in container:
                try:
                    alltd = contain.find_elements_by_tag_name('td')
                    heading = alltd[0].text

                    if heading == "Id. f??d??ral de b??timent (EGID)":
                        federalbuildingidegid = alltd[1].text
                    elif heading == "Abr??viation du canton":
                        cantonabbreviation = alltd[1].text
                    elif heading == "N?? OFS de la commune":
                        ofsnumberofmunicipality = alltd[1].text
                    elif heading == "Nom le la commune":
                        nameoftown = alltd[1].text
                    elif heading == "Id. f??d??ral d'immeuble (EGRID)":
                        federalbuildingid = alltd[1].text
                    elif heading == "N?? de secteur du RF":
                        rfsectornumber = alltd[1].text
                    elif heading == "N?? d'immeuble":
                        buildingnumber = alltd[1].text
                    elif heading == "Suffixe du n?? d'immeuble":
                        buildingnumbersuffix = alltd[1].text
                    elif heading == "Type d'immeuble":
                        typeofbuilding = alltd[1].text
                    elif heading == "N?? officiel de b??timent":
                        officialbuildingnumber = alltd[1].text
                    elif heading == "Nom du b??timent":
                        nameofbuilding = alltd[1].text
                    elif heading == "Coordonn??e E du b??timent":
                        ecordinateofbuilding = alltd[1].text
                    elif heading == "Coordonn??e N du b??timent":
                        mcordinateofbuilding = alltd[1].text
                    elif heading == "Provenance des coordonn??es":
                        sourceofcontactdetails = alltd[1].text
                    elif heading == "Statut du b??timent":
                        buildingstatus = alltd[1].text
                    elif heading == "Cat??gorie de b??timent":
                        buildingcategory = alltd[1].text
                    elif heading == "Classe de b??timent":
                        buildingclass = alltd[1].text
                    elif heading == "Ann??e de construction du b??timent":
                        yearofconstructionofbuilding = alltd[1].text
                    elif heading == "Mois de construction du b??timent":
                        monthofbuildingconstruction = alltd[1].text
                    elif heading == "Epoque de construction":
                        constructionperiod = alltd[1].text
                    elif heading == "Ann??e de d??molition du b??timent":
                        yearofdemolitionofbuilding = alltd[1].text
                    elif heading == "Surface du b??timent [m2]":
                        buildingaream2 = alltd[1].text
                    elif heading == "Volume du b??timent [m3]":
                        buildingvolumem3 = alltd[1].text
                    elif heading == "Volume du b??timent : norme":
                        buildingvolumestandard = alltd[1].text
                    elif heading == "Volume du b??timent : indication sur la donn??e":
                        buildingvolumeindicationofdata = alltd[1].text
                    elif heading == "Nombre de niveaux":
                        numberoflevels = alltd[1].text
                    elif heading == "Nombre d'" + 'enregistrements "logements"':
                        numberofaccommodationrecords = alltd[1].text
                    elif heading == "Nombre de pi??ces d???hab. ind??p.":
                        numberofroomsofhabit = alltd[1].text
                    elif heading == "Etat des donn??es":
                        if datastatecount == 1:
                            datastate = alltd[1].text
                            datastatecount = datastatecount + 1
                        else:
                            pass
                    elif heading == "Id. f??d??ral d???entr??e (EDID)":
                        federalentryidegid = alltd[1].text
                    elif heading == "Id. f??d??ral d???adresse de b??timent (EGAID)":
                        federalbuildingaddressidegaid = alltd[1].text
                    elif heading == "N?? d???entr??e du b??timent":
                        buildingentrynumber = alltd[1].text
                    elif heading == "Id. f??d??ral de rue (ESID)":
                        federalstreetid = alltd[1].text
                    elif heading == "D??signation de la rue FR":
                        streetdesignationfr = alltd[1].text
                    elif heading == "D??signation abr??g??e de la rue FR":
                        shortdesignationfr = alltd[1].text
                    elif heading == "R??f??rence de l'index FR":
                        indexreferencefr = alltd[1].text
                    elif heading == "Langue du nom de rue FR":
                        languagestreetnamefr = alltd[1].text
                    elif heading == "D??signation officielle":
                        officialdesignation = alltd[1].text
                    elif heading == "NPA":
                        postcode = alltd[1].text
                    elif heading == "Chiffre compl. du NPA":
                        digitcomplpostcode = alltd[1].text
                    elif heading == "Localit??":
                        locality = alltd[1].text
                    elif heading == "Coordonn??e E de l'entr??e":
                        ecordinateinput = alltd[1].text
                    elif heading == "Coordonn??e N de l'entr??e":
                        ncordinateinput = alltd[1].text
                    elif heading == "Adresse officielle":
                        officialaddress = alltd[1].text
                    elif heading == "??tat des donn??es":
                        if datastatecount == 2:
                            datastateentryinfo = alltd[1].text
                            datastatecount = datastatecount + 1
                        else:
                            pass
                    elif heading == "Id. f??d??ral de logement (EWID)":
                        federalhousingidewid = alltd[1].text
                    elif heading == "N?? administratif de logement":
                        administrativehousingnumber = alltd[1].text
                    elif heading == "Etage":
                        stage = alltd[1].text
                    elif heading == "Logement sur plusieurs ??tages":
                        multistoryaccomodation = alltd[1].text
                    elif heading == "N?? physique de logement":
                        physicalhousingnumber = alltd[1].text
                    elif heading == "Situation sur l'??tage":
                        situationonfloor = alltd[1].text
                    elif heading == "Statut du logement":
                        housingstatus = alltd[1].text
                    elif heading == "??tat des donn??es":
                        if datastatecount == 3:
                            datastateaccomodationinfo = alltd[1].text
                            datastatecount = datastatecount + 1
                        else:
                            pass
                except Exception:
                    pass

            print("Data State Accomodation: ", datastateaccomodationinfo)
            print("Housing Status: ", housingstatus)
            print("Situation Floor: ", situationonfloor)
            print("Physical Housing Number: ", physicalhousingnumber)
            print("Multi Story Accomation: ", multistoryaccomodation)
            print("Stage: ", stage)
            print("Administrative Housing Number: ", administrativehousingnumber)
            print("Federal Housing Id EWID: ", federalhousingidewid)
            print("Data State Entry Info: ", datastateentryinfo)
            print("White Space Link: ", linktowhite)
            print("Official Address: ", officialaddress)
            print("N Cordinate Input: ", ncordinateinput)
            print("E cordinate Input: ", ecordinateinput)
            print("Locality: ", locality)
            print("Digit Compl Postcode: ", digitcomplpostcode)
            print("Postcode: ", postcode)
            print("Official Designation: ", officialdesignation)
            print("Languages Street Name Fr: ", languagestreetnamefr)
            print("Index Reference Fr: ", indexreferencefr)
            print("Short Designation Fr: ", shortdesignationfr)
            print("Street Designation Fr: ", streetdesignationfr)
            print("Federal Street Id: ", federalstreetid)
            print("Building Entry Number: ", buildingentrynumber)
            print("Federal Building Address Id EGAID: ", federalbuildingaddressidegaid)
            print("Federal Entry Id EGID: ", federalentryidegid)
            print("Data State : ", datastate)
            print("Number Of Rooms Of Habit: ", numberofroomsofhabit)
            print("Number Of Accomation Records: ", numberofaccommodationrecords)
            print("Number Of Levels: ", numberoflevels)
            print("Building Volume Indication Of Data: ", buildingvolumeindicationofdata)
            print("Building Volume Standard: ", buildingvolumestandard)
            print("Building Volume m3: ", buildingvolumem3)
            print("Building Area m2: ", buildingaream2)
            print("Year Of Demolition Of Building: ", yearofdemolitionofbuilding)
            print("Construction Period: ", constructionperiod)
            print("Month Of Building Construction: ", monthofbuildingconstruction)
            print("Year Of Construction Of Building: ", yearofconstructionofbuilding)
            print("Building Class: ", buildingclass)
            print("Building Category: ", buildingcategory)
            print("Building Status: ", buildingstatus)
            print("Source Of Contact Details: ", sourceofcontactdetails)
            print("M Cordinate Of Building: ", mcordinateofbuilding)
            print("E Cordinate Of Building: ", ecordinateofbuilding)
            print("Name Of Building: ", nameofbuilding)
            print("Official Building Number: ", officialbuildingnumber)
            print("Type Of Building: ", typeofbuilding)
            print("Building Number Suffix: ", buildingnumbersuffix)
            print("Building Number: ", buildingnumber)
            print("Rf Sector Number: ", rfsectornumber)
            print("Federal Building ID: ", federalbuildingid)
            print("Name Of Town: ", nameoftown)
            print("OFS Number Of Municipality: ", ofsnumberofmunicipality)
            print("Canton Abbreviation: ", cantonabbreviation)
            print("Federal Building Id EGID: ", federalbuildingidegid)
            print("*********************************************")

            officers = []
            rfinfolink = ''
            try:
                rfinfolink = "http://www.rfinfo.vd.ch/rfinfo.php?no_commune=" + ofsnumberofmunicipality + "&no_immeuble=" + buildingnumber
                print("RF INFO LINK: ",rfinfolink)
                driver.get(rfinfolink)
                try:
                    WebDriverWait(driver, 10).until(
                        expected_conditions.visibility_of_element_located((By.TAG_NAME, 'tbody')))
                except Exception:
                    try:
                        ik = 1
                        WebDriverWait(driver, 10).until(
                            expected_conditions.visibility_of_element_located((By.ID, 'capchaContainer')))
                        emailtome()
                        print("CAPCTHA DETECTED! MAIL HAS BEEN SENT")
                        while ik != 2:
                            try:
                                WebDriverWait(driver, 10).until(
                                    expected_conditions.visibility_of_element_located((By.ID, 'capchaContainer')))
                            except Exception:
                                break
                            time.sleep(10)
                    except Exception:
                        pass
                tb = driver.find_element_by_tag_name('tbody')
                containerrf = tb.find_elements_by_tag_name('tr')
                idex = 0
                for containrf in containerrf:
                    td = containrf.find_element_by_tag_name('td').text
                    if "Propri??taire(s)" in td:
                        for jk in range(idex + 1, len(containerrf)):
                            officers.append(containerrf[jk].find_element_by_tag_name('td').text)
                    else:
                        pass
                    idex = idex + 1

                print("OFFICERS: ", officers)

                for ofc in officers:
                    print("Scrapping ", ofc, " Details")
                    try:
                        ofc1 = ofc.replace(" ", "+")
                    except Exception:
                        ofc1 = ofc
                        pass
                    try:
                        street = streetdesignationfr.replace(" ", "+")
                    except Exception:
                        street = streetdesignationfr
                        pass
                    localchlinkoff = "https://www.local.ch/fr/q?what=" + ofc1 + "&where=" + street + "+" + postcode + "+" + locality
                    print("LOCAL CH LINK MADE: ", localchlinkoff)

                    try:
                        driver.get(localchlinkoff)
                        WebDriverWait(driver, 5).until(
                            expected_conditions.visibility_of_element_located(
                                (By.XPATH,
                                 "//h1[@class='search-header-results-title lui-margin-top-zero lui-margin-bottom-s']")))
                        WebDriverWait(driver, 3).until(
                            expected_conditions.visibility_of_all_elements_located(
                                (By.XPATH,
                                 "//div[@class='js-entry-card-container row lui-margin-vertical-xs lui-sm-margin-vertical-m']")))
                        linkno = 1
                        pgno = 1
                        x = 1
                        try:
                            time.sleep(2)
                            driver.find_element_by_xpath('//*[@id="onetrust-accept-btn-handler"]').click()
                        except Exception:
                            pass
                        for i in range(1000):
                            try:
                                WebDriverWait(driver, 8).until(
                                    expected_conditions.visibility_of_all_elements_located(
                                        (By.XPATH,
                                         "//div[@class='js-entry-card-container row lui-margin-vertical-xs lui-sm-margin-vertical-m']")))
                                container = driver.find_elements_by_xpath(
                                    "//div[@class='js-entry-card-container row lui-margin-vertical-xs lui-sm-margin-vertical-m']")
                                for contain in container:
                                    locallink = contain.find_element_by_tag_name('a').get_attribute('href')
                                    print("Page No: ", pgno)
                                    print("Link No: ", linkno)
                                    print("Local Ch Link: ", locallink)
                                    linkno = linkno + 1
                                    namelocalch = contain.find_element_by_xpath(
                                        ".//h2[@class='lui-margin-vertical-zero card-info-title']").text

                                    try:
                                        address = contain.find_element_by_xpath(
                                            ".//div[@class='card-info-address']").text
                                    except Exception:
                                        address = ''
                                        pass

                                    try:
                                        address = address.split(", ")
                                        addresslocalch1 = address[0]
                                        addresslocalch = addresslocalch1
                                        housenumber1 = addresslocalch1.split(" ")
                                        housenumber = housenumber1[-1]
                                        address1 = address[1].split(" ")
                                        postalcodelocalch = address1[0]
                                        localitylocalch = address1[1]
                                        try:
                                            addresslocalch.replace(housenumber, "")
                                        except Exception:
                                            pass
                                    except Exception:
                                        addresslocalch = ''
                                        housenumber = ''
                                        postalcodelocalch = ''
                                        localitylocalch = ''
                                        pass

                                    try:
                                        phonelocalch = contain.find_element_by_xpath(
                                            ".//a[@title='Appeler']").get_attribute(
                                            'href')
                                    except Exception:
                                        phonelocalch = ''
                                        pass

                                    try:
                                        phonelocalch = phonelocalch.replace("tel:", "")
                                    except Exception:
                                        pass

                                    print("Officer Name Local Ch: ", namelocalch)
                                    print("Address Local Ch: ", addresslocalch)
                                    print("House Number Local Ch: ", housenumber)
                                    print("Postal Code Local Ch: ", postalcodelocalch)
                                    print("Locality Local Ch: ", localitylocalch)
                                    print("Phone Local Ch: ", phonelocalch)
                                    sheet1.write(mi, 65, "True")
                                    sheet1.write(mi, 64, phonelocalch)
                                    sheet1.write(mi, 63, localitylocalch)
                                    sheet1.write(mi, 62, postalcodelocalch)
                                    sheet1.write(mi, 61, housenumber)
                                    sheet1.write(mi, 60, addresslocalch)
                                    sheet1.write(mi, 59, namelocalch)
                                    sheet1.write(mi, 58, locallink)
                                    sheet1.write(mi, 57, localchlinkoff)
                                    sheet1.write(mi, 56, ofc)
                                    sheet1.write(mi, 55, rfinfolink)
                                    sheet1.write(mi, 54, datastateaccomodationinfo)
                                    sheet1.write(mi, 53, housingstatus)
                                    sheet1.write(mi, 52, situationonfloor)
                                    sheet1.write(mi, 51, physicalhousingnumber)
                                    sheet1.write(mi, 50, multistoryaccomodation)
                                    sheet1.write(mi, 49, stage)
                                    sheet1.write(mi, 48, administrativehousingnumber)
                                    sheet1.write(mi, 47, federalhousingidewid)
                                    sheet1.write(mi, 46, linktowhite)
                                    sheet1.write(mi, 45, datastateentryinfo)
                                    sheet1.write(mi, 44, officialaddress)
                                    sheet1.write(mi, 43, ncordinateinput)
                                    sheet1.write(mi, 42, ecordinateinput)
                                    sheet1.write(mi, 41, locality)
                                    sheet1.write(mi, 40, digitcomplpostcode)
                                    sheet1.write(mi, 39, postcode)
                                    sheet1.write(mi, 38, officialdesignation)
                                    sheet1.write(mi, 37, languagestreetnamefr)
                                    sheet1.write(mi, 36, indexreferencefr)
                                    sheet1.write(mi, 35, shortdesignationfr)
                                    sheet1.write(mi, 34, streetdesignationfr)
                                    sheet1.write(mi, 33, federalstreetid)
                                    sheet1.write(mi, 32, buildingentrynumber)
                                    sheet1.write(mi, 31, federalbuildingaddressidegaid)
                                    sheet1.write(mi, 30, federalentryidegid)
                                    sheet1.write(mi, 29, datastate)
                                    sheet1.write(mi, 28, numberofroomsofhabit)
                                    sheet1.write(mi, 27, numberofaccommodationrecords)
                                    sheet1.write(mi, 26, numberoflevels)
                                    sheet1.write(mi, 25, buildingvolumeindicationofdata)
                                    sheet1.write(mi, 24, buildingvolumestandard)
                                    sheet1.write(mi, 23, buildingvolumem3)
                                    sheet1.write(mi, 22, buildingaream2)
                                    sheet1.write(mi, 21, yearofdemolitionofbuilding)
                                    sheet1.write(mi, 20, constructionperiod)
                                    sheet1.write(mi, 19, monthofbuildingconstruction)
                                    sheet1.write(mi, 18, yearofconstructionofbuilding)
                                    sheet1.write(mi, 17, buildingclass)
                                    sheet1.write(mi, 16, buildingcategory)
                                    sheet1.write(mi, 15, buildingstatus)
                                    sheet1.write(mi, 14, sourceofcontactdetails)
                                    sheet1.write(mi, 13, mcordinateofbuilding)
                                    sheet1.write(mi, 12, ecordinateofbuilding)
                                    sheet1.write(mi, 11, nameofbuilding)
                                    sheet1.write(mi, 10, officialbuildingnumber)
                                    sheet1.write(mi, 9, typeofbuilding)
                                    sheet1.write(mi, 8, buildingnumbersuffix)
                                    sheet1.write(mi, 7, buildingnumber)
                                    sheet1.write(mi, 6, rfsectornumber)
                                    sheet1.write(mi, 5, federalbuildingid)
                                    sheet1.write(mi, 4, nameoftown)
                                    sheet1.write(mi, 3, ofsnumberofmunicipality)
                                    sheet1.write(mi, 2, cantonabbreviation)
                                    sheet1.write(mi, 1, federalbuildingidegid)
                                    sheet1.write(mi, 0, choosenurl)
                                    kb.save('propertyownership.xls')
                                    mi = mi + 1
                                    print("")
                                    print("********************")
                                    print("")
                                cururl = driver.current_url
                                pgno = pgno + 1
                                try:
                                    if x == 1:
                                        driver.find_element_by_xpath("//a[@rel='next']").click()
                                        time.sleep(2)
                                    else:
                                        togo = ''
                                        togo = cururl.split("page=")
                                        togo1 = togo[1]
                                        togo1 = togo1[1:]
                                        pgno = str(pgno)
                                        togo1 = pgno + togo1
                                        urltogo = togo[0] + "page=" + togo1
                                        pgno = int(pgno)
                                        print("New Page Url: ", urltogo)
                                        driver.get(urltogo)
                                    x = x + 1
                                except Exception:
                                    print("Page Not Formed")
                                    break
                            except Exception:
                                print("No Page More")
                                break
                    except Exception:
                        try:
                            print("Trying New Link!")
                            locallink = driver.current_url
                            print("Local Ch Link: ", locallink)
                            namelocalch = driver.find_element_by_xpath(
                                '//h1[@class="lui-margin-vertical-zero title-card-title lui-display-flex lui-display-flex-center-aligned"]').text
                            try:
                                addresslocalch = driver.find_element_by_xpath("//span[@itemprop='streetAddress']").text
                                addre = addresslocalch.split(" ")
                                housenumber = addre[-1]
                                try:
                                    addresslocalch = addresslocalch.replace(housenumber, "")
                                except Exception:
                                    pass
                            except Exception:
                                addresslocalch = ''
                                housenumber = ''
                                pass
                            try:
                                postalcodelocalch = driver.find_element_by_xpath("//span[@itemprop='postalCode']").text
                            except Exception:
                                postalcodelocalch = ''
                                pass
                            try:
                                localitylocalch = driver.find_element_by_xpath(
                                    "//span[@itemprop='addressLocality']").text
                            except Exception:
                                localitylocalch = ''
                                pass
                            try:
                                phonelocalch = driver.find_element_by_xpath(
                                    "//meta[@itemprop='telephone']").get_attribute(
                                    'content')
                            except Exception:
                                phonelocalch = ''
                                pass

                            print("Officer Name Local Ch: ", namelocalch)
                            print("Address Local Ch: ", addresslocalch)
                            print("House Number: ", housenumber)
                            print("Postal Code Local Ch: ", postalcodelocalch)
                            print("Locality Local Ch: ", localitylocalch)
                            print("Phone Local Ch: ", phonelocalch)
                            sheet1.write(mi, 65, "True")
                            sheet1.write(mi, 64, phonelocalch)
                            sheet1.write(mi, 63, localitylocalch)
                            sheet1.write(mi, 62, postalcodelocalch)
                            sheet1.write(mi, 61, housenumber)
                            sheet1.write(mi, 60, addresslocalch)
                            sheet1.write(mi, 59, namelocalch)
                            sheet1.write(mi, 58, locallink)
                            sheet1.write(mi, 57, localchlinkoff)
                            sheet1.write(mi, 56, ofc)
                            sheet1.write(mi, 55, rfinfolink)
                            sheet1.write(mi, 54, datastateaccomodationinfo)
                            sheet1.write(mi, 53, housingstatus)
                            sheet1.write(mi, 52, situationonfloor)
                            sheet1.write(mi, 51, physicalhousingnumber)
                            sheet1.write(mi, 50, multistoryaccomodation)
                            sheet1.write(mi, 49, stage)
                            sheet1.write(mi, 48, administrativehousingnumber)
                            sheet1.write(mi, 47, federalhousingidewid)
                            sheet1.write(mi, 46, linktowhite)
                            sheet1.write(mi, 45, datastateentryinfo)
                            sheet1.write(mi, 44, officialaddress)
                            sheet1.write(mi, 43, ncordinateinput)
                            sheet1.write(mi, 42, ecordinateinput)
                            sheet1.write(mi, 41, locality)
                            sheet1.write(mi, 40, digitcomplpostcode)
                            sheet1.write(mi, 39, postcode)
                            sheet1.write(mi, 38, officialdesignation)
                            sheet1.write(mi, 37, languagestreetnamefr)
                            sheet1.write(mi, 36, indexreferencefr)
                            sheet1.write(mi, 35, shortdesignationfr)
                            sheet1.write(mi, 34, streetdesignationfr)
                            sheet1.write(mi, 33, federalstreetid)
                            sheet1.write(mi, 32, buildingentrynumber)
                            sheet1.write(mi, 31, federalbuildingaddressidegaid)
                            sheet1.write(mi, 30, federalentryidegid)
                            sheet1.write(mi, 29, datastate)
                            sheet1.write(mi, 28, numberofroomsofhabit)
                            sheet1.write(mi, 27, numberofaccommodationrecords)
                            sheet1.write(mi, 26, numberoflevels)
                            sheet1.write(mi, 25, buildingvolumeindicationofdata)
                            sheet1.write(mi, 24, buildingvolumestandard)
                            sheet1.write(mi, 23, buildingvolumem3)
                            sheet1.write(mi, 22, buildingaream2)
                            sheet1.write(mi, 21, yearofdemolitionofbuilding)
                            sheet1.write(mi, 20, constructionperiod)
                            sheet1.write(mi, 19, monthofbuildingconstruction)
                            sheet1.write(mi, 18, yearofconstructionofbuilding)
                            sheet1.write(mi, 17, buildingclass)
                            sheet1.write(mi, 16, buildingcategory)
                            sheet1.write(mi, 15, buildingstatus)
                            sheet1.write(mi, 14, sourceofcontactdetails)
                            sheet1.write(mi, 13, mcordinateofbuilding)
                            sheet1.write(mi, 12, ecordinateofbuilding)
                            sheet1.write(mi, 11, nameofbuilding)
                            sheet1.write(mi, 10, officialbuildingnumber)
                            sheet1.write(mi, 9, typeofbuilding)
                            sheet1.write(mi, 8, buildingnumbersuffix)
                            sheet1.write(mi, 7, buildingnumber)
                            sheet1.write(mi, 6, rfsectornumber)
                            sheet1.write(mi, 5, federalbuildingid)
                            sheet1.write(mi, 4, nameoftown)
                            sheet1.write(mi, 3, ofsnumberofmunicipality)
                            sheet1.write(mi, 2, cantonabbreviation)
                            sheet1.write(mi, 1, federalbuildingidegid)
                            sheet1.write(mi, 0, choosenurl)
                            kb.save('propertyownership.xls')
                            mi = mi + 1
                            print("")
                            print("********************")
                            print("")
                        except Exception:
                            print("NO DATA FOR THIS")
                            sheet1.write(mi, 65, "False")
                            sheet1.write(mi, 64, "")
                            sheet1.write(mi, 63, "")
                            sheet1.write(mi, 62, "")
                            sheet1.write(mi, 61, "")
                            sheet1.write(mi, 60, "")
                            sheet1.write(mi, 59, "")
                            sheet1.write(mi, 58, "")
                            sheet1.write(mi, 57, localchlinkoff)
                            sheet1.write(mi, 56, ofc)
                            sheet1.write(mi, 55, rfinfolink)
                            sheet1.write(mi, 54, datastateaccomodationinfo)
                            sheet1.write(mi, 53, housingstatus)
                            sheet1.write(mi, 52, situationonfloor)
                            sheet1.write(mi, 51, physicalhousingnumber)
                            sheet1.write(mi, 50, multistoryaccomodation)
                            sheet1.write(mi, 49, stage)
                            sheet1.write(mi, 48, administrativehousingnumber)
                            sheet1.write(mi, 47, federalhousingidewid)
                            sheet1.write(mi, 46, linktowhite)
                            sheet1.write(mi, 45, datastateentryinfo)
                            sheet1.write(mi, 44, officialaddress)
                            sheet1.write(mi, 43, ncordinateinput)
                            sheet1.write(mi, 42, ecordinateinput)
                            sheet1.write(mi, 41, locality)
                            sheet1.write(mi, 40, digitcomplpostcode)
                            sheet1.write(mi, 39, postcode)
                            sheet1.write(mi, 38, officialdesignation)
                            sheet1.write(mi, 37, languagestreetnamefr)
                            sheet1.write(mi, 36, indexreferencefr)
                            sheet1.write(mi, 35, shortdesignationfr)
                            sheet1.write(mi, 34, streetdesignationfr)
                            sheet1.write(mi, 33, federalstreetid)
                            sheet1.write(mi, 32, buildingentrynumber)
                            sheet1.write(mi, 31, federalbuildingaddressidegaid)
                            sheet1.write(mi, 30, federalentryidegid)
                            sheet1.write(mi, 29, datastate)
                            sheet1.write(mi, 28, numberofroomsofhabit)
                            sheet1.write(mi, 27, numberofaccommodationrecords)
                            sheet1.write(mi, 26, numberoflevels)
                            sheet1.write(mi, 25, buildingvolumeindicationofdata)
                            sheet1.write(mi, 24, buildingvolumestandard)
                            sheet1.write(mi, 23, buildingvolumem3)
                            sheet1.write(mi, 22, buildingaream2)
                            sheet1.write(mi, 21, yearofdemolitionofbuilding)
                            sheet1.write(mi, 20, constructionperiod)
                            sheet1.write(mi, 19, monthofbuildingconstruction)
                            sheet1.write(mi, 18, yearofconstructionofbuilding)
                            sheet1.write(mi, 17, buildingclass)
                            sheet1.write(mi, 16, buildingcategory)
                            sheet1.write(mi, 15, buildingstatus)
                            sheet1.write(mi, 14, sourceofcontactdetails)
                            sheet1.write(mi, 13, mcordinateofbuilding)
                            sheet1.write(mi, 12, ecordinateofbuilding)
                            sheet1.write(mi, 11, nameofbuilding)
                            sheet1.write(mi, 10, officialbuildingnumber)
                            sheet1.write(mi, 9, typeofbuilding)
                            sheet1.write(mi, 8, buildingnumbersuffix)
                            sheet1.write(mi, 7, buildingnumber)
                            sheet1.write(mi, 6, rfsectornumber)
                            sheet1.write(mi, 5, federalbuildingid)
                            sheet1.write(mi, 4, nameoftown)
                            sheet1.write(mi, 3, ofsnumberofmunicipality)
                            sheet1.write(mi, 2, cantonabbreviation)
                            sheet1.write(mi, 1, federalbuildingidegid)
                            sheet1.write(mi, 0, choosenurl)
                            kb.save('propertyownership.xls')
                            mi = mi + 1
                            pass
                    pass
            except Exception:
                print("NO OFFICER STEP 2")
                sheet1.write(mi, 65, "False")
                sheet1.write(mi, 64, "")
                sheet1.write(mi, 63, "")
                sheet1.write(mi, 62, "")
                sheet1.write(mi, 61, "")
                sheet1.write(mi, 60, "")
                sheet1.write(mi, 59, "")
                sheet1.write(mi, 58, "")
                sheet1.write(mi, 57, "")
                sheet1.write(mi, 56, "")
                sheet1.write(mi, 55, rfinfolink)
                sheet1.write(mi, 54, datastateaccomodationinfo)
                sheet1.write(mi, 53, housingstatus)
                sheet1.write(mi, 52, situationonfloor)
                sheet1.write(mi, 51, physicalhousingnumber)
                sheet1.write(mi, 50, multistoryaccomodation)
                sheet1.write(mi, 49, stage)
                sheet1.write(mi, 48, administrativehousingnumber)
                sheet1.write(mi, 47, federalhousingidewid)
                sheet1.write(mi, 46, linktowhite)
                sheet1.write(mi, 45, datastateentryinfo)
                sheet1.write(mi, 44, officialaddress)
                sheet1.write(mi, 43, ncordinateinput)
                sheet1.write(mi, 42, ecordinateinput)
                sheet1.write(mi, 41, locality)
                sheet1.write(mi, 40, digitcomplpostcode)
                sheet1.write(mi, 39, postcode)
                sheet1.write(mi, 38, officialdesignation)
                sheet1.write(mi, 37, languagestreetnamefr)
                sheet1.write(mi, 36, indexreferencefr)
                sheet1.write(mi, 35, shortdesignationfr)
                sheet1.write(mi, 34, streetdesignationfr)
                sheet1.write(mi, 33, federalstreetid)
                sheet1.write(mi, 32, buildingentrynumber)
                sheet1.write(mi, 31, federalbuildingaddressidegaid)
                sheet1.write(mi, 30, federalentryidegid)
                sheet1.write(mi, 29, datastate)
                sheet1.write(mi, 28, numberofroomsofhabit)
                sheet1.write(mi, 27, numberofaccommodationrecords)
                sheet1.write(mi, 26, numberoflevels)
                sheet1.write(mi, 25, buildingvolumeindicationofdata)
                sheet1.write(mi, 24, buildingvolumestandard)
                sheet1.write(mi, 23, buildingvolumem3)
                sheet1.write(mi, 22, buildingaream2)
                sheet1.write(mi, 21, yearofdemolitionofbuilding)
                sheet1.write(mi, 20, constructionperiod)
                sheet1.write(mi, 19, monthofbuildingconstruction)
                sheet1.write(mi, 18, yearofconstructionofbuilding)
                sheet1.write(mi, 17, buildingclass)
                sheet1.write(mi, 16, buildingcategory)
                sheet1.write(mi, 15, buildingstatus)
                sheet1.write(mi, 14, sourceofcontactdetails)
                sheet1.write(mi, 13, mcordinateofbuilding)
                sheet1.write(mi, 12, ecordinateofbuilding)
                sheet1.write(mi, 11, nameofbuilding)
                sheet1.write(mi, 10, officialbuildingnumber)
                sheet1.write(mi, 9, typeofbuilding)
                sheet1.write(mi, 8, buildingnumbersuffix)
                sheet1.write(mi, 7, buildingnumber)
                sheet1.write(mi, 6, rfsectornumber)
                sheet1.write(mi, 5, federalbuildingid)
                sheet1.write(mi, 4, nameoftown)
                sheet1.write(mi, 3, ofsnumberofmunicipality)
                sheet1.write(mi, 2, cantonabbreviation)
                sheet1.write(mi, 1, federalbuildingidegid)
                sheet1.write(mi, 0, choosenurl)
                kb.save('propertyownership.xls')
                mi = mi + 1
                pass
        except Exception:
            print("NOTHING WORKED MOVING TO NEXT LINK")
            pass








    print("time elapsed: {:.2f}s".format(time.time() - start_time))

# create the driver object.
search_keyword = "Web Scraping"
driver= configure_driver()
getCourses(driver, search_keyword)

# close the driver.#3driver.close()














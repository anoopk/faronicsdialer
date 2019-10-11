from __future__ import print_function

import base64
import datetime
import time
from datetime import datetime
from tkinter import *

import pandas as pd
import win32api
from googleapiclient import discovery
from bs4 import BeautifulSoup
from dateutil.parser import parse
from httplib2 import Http
from oauth2client import file, client, tools
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import flask

import common
import google_sheet_common as gs


def getvalue(message):
    root = Tk()
    root.title('Faronics')
    mystring = StringVar()

    def close_window(root):
        root.destroy()

    def getvalue():
        close_window(root)

    Label(root, text=message).grid(row=0)  # label
    Entry(root, textvariable=mystring).grid(row=0, column=1, sticky=E)  # entry textbox
    WSignUp = Button(root, text="Submit", command=getvalue).grid(row=3, column=0, sticky=W)  # button
    root.mainloop()
    return mystring.get()


def getvcode(email, storagefile):
    # Creating a storage.JSON file with authentication details
    time.sleep(5)
    SCOPES = 'https://www.googleapis.com/auth/gmail.modify'
    store = file.Storage(storagefile)
    creds = store.get()

    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
        creds = tools.run_flow(flow, store)
    GMAIL = discovery.build('gmail', 'v1', http=creds.authorize(Http()))

    user_id = 'me'
    vcode = GMAIL.users().messages().list(userId='me', labelIds=["UNREAD"]).execute()
    for mssg in vcode['messages']:
        try:
            temp_dict = {}
            m_id = mssg['id']
            message = GMAIL.users().messages().get(userId=user_id, id=m_id).execute()  # fetch the message using API
            payld = message['payload']
            headr = payld['headers']

            for two in headr:
                if two['name'] == 'Date':
                    msg_date = two['value']
                    date_parse = (parse(msg_date))
                    m_date = (date_parse.date())

                else:
                    pass
            diff = (datetime.now() - date_parse.replace(tzinfo=None)).total_seconds() / 60
            temp_dict['Snippet'] = message['snippet']  # fetching message snippet
            # print(str(BeautifulSoup((base64.b64decode(bytes(payld['body']['data'], 'UTF-8'))), "lxml")))
            verificationcode = str(BeautifulSoup((base64.b64decode(bytes(payld['body']['data'], 'UTF-8'))), "lxml"))
            verificationcode = (verificationcode.split("Code:")[1].split("\r")[0]).strip()
            # print(verificationcode)
            break
        except:
            print("except")
            pass

    return verificationcode


try:
    driver = webdriver.Firefox()
    wait = WebDriverWait(driver, 60)
    driver.get(url='https://faronicsna.my.salesforce.com/')
    baseurl = 'https://faronicsna.my.salesforce.com/'

    usrdf = gs.getsheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'control')
    uemail = getvalue("Please Enter Your Email ID :")
    df1 = usrdf[usrdf['sfemail'] == uemail].reset_index(drop=True)
    if df1.shape[0] > 0:
        email = df1['sfemail'].item()
        password = df1['sfpassword'].item()
        storagefile = df1['storage'].item()
        # print(email, password, storagefile)
    else:
        print("Email id does not exists. Please contact administrator.")
        sys.exit()

    sfUserTb = driver.find_element_by_id('username')
    sfPwdTb = driver.find_element_by_id('password')
    loginBtn = driver.find_element_by_id('Login')

    sfUserTb.send_keys(email)
    # print("entered email")
    sfPwdTb.send_keys(password)
    # print("entered pwd")
    loginBtn.click()
    # print("Login button click")
    time.sleep(5)
    identity = driver.find_elements_by_id("emc")
    # print(len(identity))
    if len(identity) > 0:
        # print("Enter verificationcode")
        # # time.sleep(5)
        driver.find_element_by_id("emc").send_keys(getvcode(email, storagefile))
        # verificationcode = getvalue("Please Enter verification code :")
        # driver.find_element_by_id("emc").send_keys(verificationcode)
        driver.find_element_by_id("save").click()
        time.sleep(3)
    wait.until(EC.visibility_of_element_located((By.ID, "userNavLabel")))
    usrname = driver.find_element_by_id("userNavLabel").text
    win32api.MessageBox(0, 'Welcome ' + usrname + '!, Click on OK button to proceed with calling', 'Faronics',
                        0x00001000)

    rendatadf = gs.getsheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'control')
    thisusr = rendatadf[rendatadf['username'] == usrname].reset_index(drop=True)
    oppsdf = gs.getsheet(thisusr['sheetid'].item(), thisusr['sheetname'].item())
    oppsdf['tzno'] = ''
    oppsdf.loc[oppsdf['Timezone'] == 'Eastern', 'tzno'] = 1
    oppsdf.loc[oppsdf['Timezone'] == 'Central', 'tzno'] = 2
    oppsdf.loc[oppsdf['Timezone'] == 'Mountain', 'tzno'] = 3
    oppsdf.loc[oppsdf['Timezone'] == 'Pacific', 'tzno'] = 4
    oppsdf.loc[oppsdf['Timezone'] == 'Time Zone Not Found', 'tzno'] = 0
    oppsdf.loc[oppsdf['Timezone'] == 'Hawaii', 'tzno'] = 0
    oppsdf['tzno'] = pd.to_numeric(oppsdf['tzno'], downcast='signed')
    oppsdf.drop_duplicates('Account Name', inplace=True)
    donedf = gs.getsheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'call_log')
    oppsdf['Done'] = oppsdf['SageCRMid'].apply(lambda x: donedf[donedf['SageCRMid'] == (x)]['alertclosetime'].any())
    oppsdf.loc[oppsdf['Done'] == False, 'Done'] = 0
    oppsdf['Done'] = pd.to_datetime(oppsdf['Done'])
    oppsdf.sort_values(['Done', 'tzno'], ascending=[True, False], inplace=True)
    # oppsdf.sort_values('tzno', ascending=False, inplace=True)
    oppsdf.reset_index(drop=True, inplace=True)

    cdata = [[str(datetime.now()), "Login", "", '', usrname]]
    gs.append_sheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'call_log', cdata)


    def startcalling():
        print("Starting loop / data refresh again")
        if oppsdf.shape[0] == 0:
            win32api.MessageBox(0, 'Opps not available. Please reset the report filter / ', 'Faronics', 0x00001000)
        for i in range(oppsdf.shape[0]):

            dial = common.gettimenow(oppsdf['Timezone'][i])
            if dial:
                print("Call to ", oppsdf['Account Name'][i], "with Maintenance Amount", oppsdf['Maintenance Amount'][i],
                      'and licences covered', oppsdf['Licenses Covered'][i])
                driver.find_element_by_id("phSearchInput").send_keys(oppsdf['Account Name'][i] + Keys.ENTER)
                time.sleep(3)
                aclist = driver.find_elements_by_xpath(
                    "//div[@id='Account_body']/table//tr[contains(@class,'dataRow')]")
                # print(len(aclist))
                for j in range(len(aclist)):
                    acname = driver.find_element_by_xpath(
                        "//div[@id='Account_body']/table//tr[contains(@class,'dataRow')][" + str(j + 1) + "]/th").text
                    acsage = driver.find_element_by_xpath(
                        "//div[@id='Account_body']/table//tr[contains(@class,'dataRow')][" + str(
                            j + 1) + "]/td[2]").text
                    # print(acname.lower(), "--", oppsdf['Account Name'][i].lower())
                    # print(acsage, "--", oppsdf['SageCRMid'][i])
                    # if acname.lower() == oppsdf['Account Name'][i].lower():
                    if acsage == oppsdf['SageCRMid'][i]:
                        # print("Found")
                        selectaccount = True
                        driver.find_element_by_xpath(
                            "//div[@id='Account_body']/table//tr[contains(@class,'dataRow')][" + str(
                                j + 1) + "]/th/a").click()
                        break
                    else:
                        selectaccount = False

                activityhistory = len(driver.find_elements_by_xpath(
                    "//div[contains(@id,'RelatedHistoryList_body')]/table//tr[contains(@class,'dataRow')]"))
                # print("activityhistory",activityhistory)
                if activityhistory > 0:
                    actdate = pd.to_datetime(driver.find_element_by_xpath(
                        "//div[contains(@id,'RelatedHistoryList_body')]/table//tr[contains(@class,'dataRow')][1]/td[7]").text)
                    assignedto = driver.find_element_by_xpath(
                        "//div[contains(@id,'RelatedHistoryList_body')]/table//tr[contains(@class,'dataRow')][1]/td[6]").text
                    subject = driver.find_element_by_xpath(
                        "//div[contains(@id,'RelatedHistoryList_body')]/table//tr[contains(@class,'dataRow')][1]/th").text
                else:
                    actdate, assignedto, subject = None, None, None

                if selectaccount:
                    callmsg = "Account Details\n" + "Account Name - " + oppsdf['Account Name'][i] + "\nSageCRMid - " + \
                              oppsdf['SageCRMid'][i] + "\nMaintenance Amount - " + str(
                        oppsdf['Maintenance Amount'][i]) + "\nlicences covered- " + str(
                        oppsdf['Licenses Covered'][i]) + '\n\nLast Activity \nDate: ' + str(
                        actdate) + "\nAssigned To: " + assignedto + "\nSubject: " + subject
                else:
                    callmsg = "Account Details\n" + "Account Name - " + oppsdf['Account Name'][i] + "\nSageCRMid - " + \
                              oppsdf['SageCRMid'][i] + "\nMaintenance Amount - " + str(
                        oppsdf['Maintenance Amount'][i]) + "\nlicences covered- " + str(
                        oppsdf['Licenses Covered'][i])
                win32api.MessageBox(0, callmsg, 'Faronics', 0x00001000)

                msgclose = datetime.now()
                cdata = [
                    [str(datetime.now()), oppsdf['Account Name'][i], oppsdf['Maintenance Amount'][i], str(msgclose),
                     usrname, oppsdf['Opportunity'][i], oppsdf['Maintenance End'][i], oppsdf['Licenses Covered'][i],
                     oppsdf['SageCRMid'][i]]]
                gs.append_sheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'call_log', cdata)

                # nextcallmsg = "Next account is \n" + oppsdf['Account Name'][i + 1] + "\nMaintenance Amount - " + str(
                #     oppsdf['Maintenance Amount'][i + 1]) + '\nlicences covered- ' + str(oppsdf['Licenses Covered'][i + 1])
                # win32api.MessageBox(0, nextcallmsg, 'Faronics', 0x00001000)
                # msgclose = datetime.now()
            else:
                print('Timezeone issue')
        startcalling()
    startcalling()
except:
    cdata = [[str(datetime.now()), "Logout", "", '', usrname]]
    gs.append_sheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'call_log', cdata)
    sys.exit()

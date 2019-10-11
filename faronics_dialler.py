from __future__ import print_function

import base64
import datetime
import time
from datetime import datetime
from tkinter import *

import pandas as pd
import win32api
from apiclient import discovery
from bs4 import BeautifulSoup
from dateutil.parser import parse
from httplib2 import Http
from oauth2client import file, client, tools
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

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
            print(str(BeautifulSoup((base64.b64decode(bytes(payld['body']['data'], 'UTF-8'))), "lxml")))
            verificationcode = str(BeautifulSoup((base64.b64decode(bytes(payld['body']['data'], 'UTF-8'))), "lxml"))
            verificationcode = (verificationcode.split("Code:")[1].split("\r")[0]).strip()
            print(verificationcode)
            break
        except:
            print("except")
            pass

    return verificationcode


driver = webdriver.Firefox()
wait = WebDriverWait(driver, 60)
driver.get(url='https://faronicsna.my.salesforce.com/')

usrdf = gs.getsheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'control')
uemail = getvalue("Please Enter Email ID :")
usrpwd = getvalue("Please Enter Salesforce Password :")
# df1 = usrdf[usrdf['sfemail'] == uemail].reset_index(drop=True)
# if df1.shape[0] > 0:
#     email = df1['sfemail'].item()
#     password = df1['sfpassword'].item()
#     storagefile = df1['storage'].item()
#     print(email, password, storagefile)
# else:
#     print("Email id does not exists. Please contact administrator.")
#     sys.exit()

sfUserTb = driver.find_element_by_id('username')
sfPwdTb = driver.find_element_by_id('password')
loginBtn = driver.find_element_by_id('Login')

sfUserTb.send_keys(uemail)
print("entered email")
sfPwdTb.send_keys(usrpwd)
print("entered pwd")
loginBtn.click()
print("Login button click")
time.sleep(5)
identity = driver.find_elements_by_id("emc")
print(len(identity))

if uemail == 'apakhare@faronics.com':
    fromid = "Abby Pakhare <apakhare@faronics.com>"
    bcc = 'emailtosalesforce@a-2bojn735ku0ldq4i4zbb2p5gcqpje6gnlxbzjjc7691l9zoldn.d-dbjream.na14.le.salesforce.com'
    storage = "storage_apakhare.json"
if uemail == 'afernandes@faronics.com':
    fromid = "Alan Fernandes <afernandes@faronics.com>"
    bcc = 'emailtosalesforce@2e1zvfyskjufpdwwxqueozws24hbprf6bsjgiu4ljiijjmohg6.d-dbjream.dl.le.salesforce.com'
    storage = "storage_afernandes.json"
if uemail == 'andy.singh@faronics.com':
    fromid = "Andy Singh <andy.singh@faronics.com>"
    bcc = 'emailtosalesforce@kbzaivdsdgmnftjtd8ny48xjxl9911860raqmjajt23ba9f8i.d-dbjream.na14.le.salesforce.com'
    storage = "storage_andy.singh.json"
if uemail == 'joshua.domnic@faronics.com':
    fromid = "Joshua Domnic <joshua.domnic@faronics.com>"
    bcc = 'emailtosalesforce@5-cigdfc1ahklnf9cmm4ym2240u4s7wo40jk2zqi694p21q3eef.d-dbjream.na52.le.salesforce.com'
    storage = "storage_joshua.json"
if uemail == 'rtashewale@faronics.com':
    fromid = "Robert Tashewale <rtashewale@faronics.com>"
    bcc = 'emailtosalesforce@xyilvymt48ifj402puw49c6zjluz2uvpnm4y053lypnizh11p.d-dbjream.dl.le.salesforce.com'
    storage = "storage_robert.json"

if uemail == 'tgrewal@faronics.com':
    fromid = "TJ Grewal <tgrewal@faronics.com>"
    bcc = 'emailtosalesforce@kw2alxzvn3vg9hxs1nmve588x5uade1x6lquqcwx3w2i6ixx7.d-dbjream.na14.le.salesforce.com'
    storage = "storage_tj.json"
if uemail == 'cswanson@faronics.com':
    fromid = "Catherine Swanson <cswanson@faronics.com>"
    bcc = 'emailtosalesforce@321ok2ida0d2646eb6evx4iielpailc45qr2zo5iih5o8xg9m9.d-dbjream.dl.le.salesforce.com'
    storage = "storage_catherine.json"
if uemail == 'dennis@faronics.com':
    fromid = "Dennis Winkelmans <dennis@faronics.com>"
    bcc = 'emailtosalesforce@kvhrvio9hi5ohse7r64qic7xrp5tl4u2tg7khi73mz34wr7o6.d-dbjream.dl.le.salesforce.com'
    storage = "storage_dennis.json"
if uemail == 'daniel.gelinas@faronics.com':
    fromid = "Daniel Gelinas <daniel.gelinas@faronics.com>"
    bcc = 'emailtosalesforce@7-122ld17owba2wbfeho5lfzrer3amdj6r1xe2twbec80y5ng2u8.d-dbjream.na14.le.salesforce.com'
    storage = "storage_daniel.json"
if uemail == 'elu@faronics.com':
    fromid = "Effort Lu <elu@faronics.com>"
    bcc = 'emailtosalesforce@h-180wlrlfkjg4ojhfqu2vd9jdtoun5iwuvhyezhgpc2s99lcxeq.d-dbjream.na14.le.salesforce.com'
    storage = "storage_effort.json"
if uemail == 'gjones@faronics.com':
    fromid = "Gary Jones <gjones@faronics.com>"
    bcc = 'emailtosalesforce@e-31ozr6dl6ths5qyn4xgnfzyb300iw5f3x5y7r7h46gmbtxxrg1.d-dbjream.na14.le.salesforce.com'
    storage = "storage_gary.json"
if uemail == 'jshearing@faronics.com':
    fromid = "Judy Shearing <jshearing@faronics.com>"
    bcc = 'emailtosalesforce@b-bv657d700dsr98qym2s4syp7wwi8j9wbd7victudu35hvmesm.d-dbjream.dl.le.salesforce.com'
    storage = "storage_judy.json"
if uemail == 'margarita.angel@faronics.com':
    fromid = "Margarita Angel <margarita.angel@faronics.com>"
    bcc = 'emailtosalesforce@fwtr3w0j5szkvng7yxedf6ui6rhetsn0vd43bvzv9olf9t3jr.d-dbjream.dl.le.salesforce.com'
    storage = "storage_margarita.json"
if uemail == 'michael.weaver@faronics.com':
    fromid = "Michael Weaver <michael.weaver@faronics.com>"
    bcc = 'emailtosalesforce@9-krwm1ru1itew5u43qmq6p35qu5odaremtig36x0l5u4bvhs5o.d-dbjream.na52.le.salesforce.com'
    storage = "storage_michael.json"
if uemail == 'rdhanji@faronics.com':
    fromid = "Rahim Dhanji <rdhanji@faronics.com>"
    bcc = 'emailtosalesforce@q6f18genecoumkoynoyvctgg04f3p37mwbyk5r6b7rxw6afpk.d-dbjream.dl.le.salesforce.com'
    storage = "storage_rahim.json"
if uemail == 'ssalamian@faronics.com':
    fromid = "Shahrzad Salamian <ssalamian@faronics.com>"
    bcc = 'emailtosalesforce@s-oveujyvbpno1s9fuf8xrgq77lczv3gq0cj2u2v0b9eoeyssnt.d-dbjream.dl.le.salesforce.com'
    storage = "storage_ssalamian.json"
if uemail == 'sdejesus@faronics.com':
    fromid = "Sophia De Jesus <sdejesus@faronics.com>"
    bcc = 'emailtosalesforce@imxupkfr0tgqmzn5eu68y2ob3jks8d7rwlnqv8pupsjgahc9.d-dbjream.dl.le.salesforce.com'
    storage = "storage_sophia.json"
if uemail == 'victor.reyes@faronics.com':
    fromid = "Victor Reyes <victor.reyes@faronics.com>"
    bcc = 'emailtosalesforce@a-2qjswaa85h83huxcni7oj87rm96r6chf3k4642fqhxhonif4iv.d-dbjream.na52.le.salesforce.com'
    storage = "storage_victor.json"
if uemail == 'amidha@faronics.com':
    fromid = "Abhay Midha <amidha@faronics.com>"
    bcc = 'emailtosalesforce@203kuinf7ynxlkggv8nhwnhp1dxdtejoksxl45drk1xqjx18i7.d-dbjream.dl.le.salesforce.com'
    storage = "storage_abhay.json"

storagefile = storage

if len(identity) > 0:
    driver.find_element_by_id("emc").send_keys(getvcode(uemail, storagefile))
    driver.find_element_by_id("save").click()
    time.sleep(3)
wait.until(EC.visibility_of_element_located((By.ID, "userNavLabel")))
usrname = driver.find_element_by_id("userNavLabel").text
win32api.MessageBox(0, 'Welcome ' + usrname + '!, Click on OK button to proceed with calling', 'Faronics', 0x00001000)

rendatadf = gs.getsheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'control')
thisusr = rendatadf[rendatadf['username'] == usrname].reset_index(drop=True)


#           GET SHEETID AND SHEETNAME FROM USER AS INPUT
sheetid = getvalue("Please Enter Email ID :")
sheetname = getvalue("Please Enter Email ID :")

# oppsdf = gs.getsheet(thisusr['sheetid'].item(), thisusr['sheetname'].item())
oppsdf = gs.getsheet(sheetid, sheetname)
#
#
# oppsdf['tzno'] = ''
# oppsdf.loc[oppsdf['Timezone'] == 'Eastern', 'tzno'] = 1
# oppsdf.loc[oppsdf['Timezone'] == 'Central', 'tzno'] = 2
# oppsdf.loc[oppsdf['Timezone'] == 'Mountain', 'tzno'] = 3
# oppsdf.loc[oppsdf['Timezone'] == 'Pacific', 'tzno'] = 4
# oppsdf.loc[oppsdf['Timezone'] == 'Time Zone Not Found', 'tzno'] = 0
# oppsdf.loc[oppsdf['Timezone'] == 'Hawaii', 'tzno'] = 0
# oppsdf['tzno'] = pd.to_numeric(oppsdf['tzno'], downcast='signed')
oppsdf.drop_duplicates('Account Name', inplace=True)
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
            callmsg = "Call to \n" + "Account Name - " + oppsdf['Account Name'][i] + "\nSageCRMid - " + \
                      oppsdf['SageCRMid'][i] + "\nMaintenance Amount - " + str(
                oppsdf['Maintenance Amount'][i]) + "\nlicences covered- " + str(oppsdf['Licenses Covered'][i])
            win32api.MessageBox(0, callmsg, 'Faronics', 0x00001000)
            msgclose = datetime.now()
            cdata = [[str(datetime.now()), oppsdf['Account Name'][i], oppsdf['Maintenance Amount'][i], str(msgclose),
                      usrname, oppsdf['Opportunity'][i], oppsdf['Maintenance End'][i], oppsdf['Licenses Covered'][i],
                      oppsdf['SageCRMid'][i]]]
            gs.append_sheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'call_log', cdata)

            nextcallmsg = "Next account is \n" + oppsdf['Account Name'][i + 1] + "\nMaintenance Amount - " + str(
                oppsdf['Maintenance Amount'][i + 1]) + '\nlicences covered- ' + str(oppsdf['Licenses Covered'][i + 1])
            win32api.MessageBox(0, nextcallmsg, 'Faronics', 0x00001000)
            msgclose = datetime.now()
        else:
            print('Timezeone issue')
    startcalling()


startcalling()
cdata = [[str(datetime.now()), "Logout", "", '', usrname]]
gs.append_sheet('16iM4AVWs9LvRVtFdrZuUaphXH-QKH79kCIHWDcgM8Ws', 'call_log', cdata)

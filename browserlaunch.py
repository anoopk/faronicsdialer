from __future__ import print_function
import base64
import datetime
import re
import sys
import time
from datetime import datetime
from datetime import timedelta
import pandas as pd
from IPython.core.display import display, HTML
from apiclient import discovery
from bs4 import BeautifulSoup
from dateutil.parser import parse
from httplib2 import Http
from oauth2client import file, client, tools
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.wait import WebDriverWait
display(HTML("<style>.container { width:98% !important; }</style>"))
pd.options.mode.chained_assignment = None  # default='warn'


def launchFirefox(binary, fp, dpath):
    opts = Options()
    opts.profile = fp
    # opts.add_argument("--headless")
    opts.add_argument("download.default_directory="+dpath)
    firefox_capabilities = DesiredCapabilities.FIREFOX
    firefox_capabilities['marionette'] = True
    driver = webdriver.Firefox(capabilities=firefox_capabilities, firefox_binary=binary, options=opts)
    return driver


def launchFFandLoginToSF(binary, fp, email, password, locpath, dpath, storagefile,labelid):
    opts = Options()
    opts.profile = fp
#     firefoxProfile.setPreference("browser.download.folderList",2);
#     firefoxProfile.setPreference("browser.download.manager.showWhenStarting",false);
#     firefoxProfile.setPreference("browser.download.dir",downloadPath);
#     firefoxProfile.setPreference("browser.helperApps.neverAsk.saveToDisk","application/pdf");
    # opts.add_argument("--headless")
    opts.add_argument("download.default_directory="+dpath)
    firefox_capabilities = DesiredCapabilities.FIREFOX
#     firefox_capabilities['marionette'] = True
    driver = webdriver.Firefox(capabilities=firefox_capabilities, firefox_binary=binary, options=opts)
    driver.get(url='https://faronicsna.my.salesforce.com/')

    sfUserTb = driver.find_element_by_id('username')
    sfPwdTb = driver.find_element_by_id('password')
    loginBtn = driver.find_element_by_id('Login')

    sfUserTb.send_keys(email)
    sfPwdTb.send_keys(password)
    loginBtn.click()
    time.sleep(5)
    identity = driver.find_elements_by_id("emc")
    if len(identity) > 0:
        time.sleep(5)
        # Creating a storage.JSON file with authentication details
        SCOPES = 'https://www.googleapis.com/auth/gmail.modify'
        store = file.Storage(locpath + storagefile)
        creds = store.get()

        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets(locpath + 'client_secret.json', SCOPES)
            creds = tools.run_flow(flow, store)
        GMAIL = discovery.build('gmail', 'v1', http=creds.authorize(Http()))

        user_id = 'me'
        vcode = GMAIL.users().messages().list(userId='me', labelIds=[labelid]).execute()
        for mssg in vcode['messages']:
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
            verificationcode = str(BeautifulSoup((base64.b64decode(bytes(payld['body']['data'], 'UTF-8'))), "lxml"))
            verificationcode = (verificationcode.split("Code:")[1].split("\r")[0]).strip()
            break

        driver.find_element_by_id("emc").send_keys(verificationcode)
        driver.find_element_by_id("save").click()
        time.sleep(2)
    return driver


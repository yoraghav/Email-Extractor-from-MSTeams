

import json
import random
import re
import time
from datetime import datetime
from threading import Timer

from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.utils import ChromeType
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from msedge.selenium_tools import Edge, EdgeOptions

browser: webdriver.Chrome = None
config = None
already_joined_ids = []
active_correlation_id = ""
conversation_link = "https://teams.microsoft.com/_#/conversations/a"
mode = 3
uuid_regex = r"\b[0-9a-f]{8}\b-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-\b[0-9a-f]{12}\b"

def load_config():
    global config
    with open('config.json', encoding='utf-8') as json_data_file:
        config = json.load(json_data_file)


def init_browser():
    global browser

    if "chrome_type" in config and config['chrome_type'] == "msedge":
        chrome_options = EdgeOptions()
        chrome_options.use_chromium = True

    else:
        chrome_options = webdriver.ChromeOptions()

    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_argument('--ignore-ssl-errors')
    chrome_options.add_argument('--use-fake-ui-for-media-stream')
    chrome_options.add_experimental_option('prefs', {
        'credentials_enable_service': False,
        'profile.default_content_setting_values.media_stream_mic': 1,
        'profile.default_content_setting_values.media_stream_camera': 1,
        'profile.default_content_setting_values.geolocation': 1,
        'profile.default_content_setting_values.notifications': 1,
        'profile': {
            'password_manager_enabled': False
        }
    })
    chrome_options.add_argument('--no-sandbox')

    chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])

    if 'chrome_type' in config:
        if config['chrome_type'] == "chromium":
            browser = webdriver.Chrome(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install(),
                                       options=chrome_options)
        elif config['chrome_type'] == "msedge":
            browser = Edge(EdgeChromiumDriverManager().install(), options=chrome_options)
        else:
            browser = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
    else:
        browser = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)

    # make the window a minimum width to show the meetings menu
    window_size = browser.get_window_size()
    if window_size['width'] < 1200:
        print("Resized window width")
        browser.set_window_size(1200, window_size['height'])

    if window_size['height'] < 850:
        print("Resized window height")
        browser.set_window_size(window_size['width'], 850)

def wait_until_found(sel, timeout, print_error=True):
    try:
        element_present = EC.visibility_of_element_located((By.CSS_SELECTOR, sel))
        WebDriverWait(browser, timeout).until(element_present)

        return browser.find_element_by_css_selector(sel)
    except exceptions.TimeoutException:
        if print_error:
            print(f"Timeout waiting for element: {sel}")
            discord_notification("Timeout error", sel)
        return None


def change_organisation(org_num):
    # Find and click the profile button
    profile_button = wait_until_found("button#personDropdown", 20)
    if profile_button is None:
        print("Something went wrong while changing the organisation")
        return

    profile_button.click()

    # Find and click the organisation with the right id
    change_org_button = wait_until_found(f"li.tenant-list-item[aria-posinset='{org_num+1}", 10)
    if change_org_button is None:
        print("Something went wrong while changing the organisation")
        return

    # if the user is already in the right organisation, return
    try:
        change_org_button.find_element_by_css_selector("button.active")
    except exceptions.NoSuchElementException:
        pass
    else:
        print("Organisation not changed (Already selected)")
        return

    change_org_button.click()
    time.sleep(5)


load_config()
mode = 1

email = config['email']
password = config['password']

if email == "":
    email = input('Email: ')

if password == "":
    password = getpass('Password: ')

init_browser()

browser.get("https://teams.microsoft.com")

if email != "" and password != "":
    login_email = wait_until_found("input[type='email']", 30)
    if login_email is not None:
        login_email.send_keys(email)

    # find the element again to avoid StaleElementReferenceException
    login_email = wait_until_found("input[type='email']", 5)
    if login_email is not None:
        login_email.send_keys(Keys.ENTER)

    login_pwd = wait_until_found("input[type='password']", 10)
    if login_pwd is not None:
        login_pwd.send_keys(password)

    # find the element again to avoid StaleElementReferenceException
    login_pwd = wait_until_found("input[type='password']", 5)
    if login_pwd is not None:
        login_pwd.send_keys(Keys.ENTER)

    keep_logged_in = wait_until_found("input[id='idBtn_Back']", 5)
    if keep_logged_in is not None:
        keep_logged_in.click()
        
    else:
        print("Login Unsuccessful, recheck entries in config.json")

    use_web_instead = wait_until_found(".use-app-lnk", 5, print_error=False)
    if use_web_instead is not None:
        use_web_instead.click()

# if additional organisations are setup in the config file
if 'organisation_num' in config and config['organisation_num'] > 0:
    change_organisation(config['organisation_num'])

print("Waiting for correct page...", end='')
# try 3 times to check #teams-app-bar is detected or if the errors can be fixed
for i in range(3):
    if wait_until_found("#teams-app-bar", 60) is None:
        # click the Try again button if teams load error
        try_again = wait_until_found("button.oops-button", 10)
        if try_again is not None:
            try_again.click()
        else:
            # if there is no Try again button to click then stop the program
            exit(1)
    else:
        # if the team-app-bar is detected then break the loop and go to the next step
        break

print("\rFound page, do not click anything on the webpage from now on.")
time.sleep(10)


urr = "https://teams.microsoft.com/_#/apps/a2da8768-95d5-419e-9441-3b539865b118/search?q="
browser.get(urr+"qqqqqqqqqqq")


qq = wait_until_found("#searchInputField",10)
qq.click()
qq.send_keys(Keys.ENTER)
time.sleep(1)
browser.get("https://teams.microsoft.com/_#/apps/a2da8768-95d5-419e-9441-3b539865b118/search?q=20EC10064") #change this only if youre raghav lol
flag=1
while(flag==1):
    time.sleep(1)
    text = str(browser.page_source)
    emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", text)
    if(len(emails)>1):
        flag=0

urr = "https://teams.microsoft.com/_#/apps/a2da8768-95d5-419e-9441-3b539865b118/search?q="
aa = config['keys']
ans = []
for a in aa:
    print("extracting for "+a)
    trig=0
    for i in range(1,10):
        lol = a+"000"+str(i)
        browser.get(urr+lol)
        time.sleep(1)
        text = str(browser.page_source)
        emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", text)
        trig+=1
        if(len(emails)==2):
            trig=0
            start = text.find("profilepicturev2?displayname=")
            end = text.find("&amp;size=HR64x64\"")
            ans.append([lol,text[start+29:end],emails[0]])
        if(trig>2):
          break
    for i in range(10,100):
        lol = a+"00"+str(i)
        browser.get(urr+lol)
        time.sleep(1)
        text = str(browser.page_source)
        emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", text)
        trig+=1
        if(len(emails)==2):
            trig=0
            start = text.find("profilepicturev2?displayname=")
            end = text.find("&amp;size=HR64x64\"")
            ans.append([lol,text[start+29:end],emails[0]])
        if(trig>2):
          break
    for i in range(100,200):
        lol = a+"0"+str(i)
        browser.get(urr+lol)
        time.sleep(1)
        text = str(browser.page_source)
        emails = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+", text)
        trig+=1
        if(len(emails)==2):
            trig=0
            start = text.find("profilepicturev2?displayname=")
            end = text.find("&amp;size=HR64x64\"")
            ans.append([lol,text[start+29:end],emails[0]])
        if(trig>2):
          break
ans


# In[ ]:


import pandas as pd
data = pd.DataFrame(ans)


# In[ ]:


data.columns = ['Rno','Name','Email']


# In[ ]:


data.to_csv('data.csv')


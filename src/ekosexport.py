#!/usr/bin/env python
import os
import re

from selenium import webdriver
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.firefox.options import Options
# from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchFrameException
from selenium.common.exceptions import ElementClickInterceptedException

class EkosExport:
    '''Class for accessing and downloading items from Ekos ERP using Selenium
    Webdriver.

    Requires geckodriver for Firefox, ChromeDriver for Chrome
    **CURRENTLY ONLY IMPLEMENTED FOR FIREFOX**

    PARAMS
    --------------
    browser : Selenium Web Browser to be used. Requires browser and the 
    correct webdriver to be installed on the machine

        Firefox --> geckodriver
        Chrome --> ChromeDriver

        **CURRENTLY ONLY IMPLEMENTED FOR FIREFOX. CHROME WILL THROW AN ERROR**

    driver_path : Location of the webdriver required for the selected browser
        
        e.g. /PATH/to/geckodriver for Firefox

    profile_dir : Location where files from the session will be downloaded on
    the local machine

        1 == tmp/
        2 == custom directory provided. Location provided by profile_dir_path
        3 == USERPROFILE

        For more information, refer to Docs
        'https://firefox-source-docs.mozilla.org/testing/geckodriver/Profiles.html'

    profile_dir_path : Location of download directory if setting custom directory

    headless : determines whether or not to run Selenium in headless mode, 
    which is necessary for running on a machine without a monitor e.g. in 
    the cloud
    '''
    def __init__(
        self,
        browser,
        driver_path,
        profile_dir,
        profile_dir_path,
        headless=False
    ):
        self.browser = browser
        self.driver_path = driver_path
        self.profile_dir = profile_dir
        self.profile_dir_path = profile_dir_path
        self.headless = headless

        if self.browser.lower() == 'firefox':
            #set profile
            self.profile = FirefoxProfile()
            #set profile directory preference
            self.profile.set_preference(
                'browser.download.folderList', 
                self.profile_dir
                )
                #if profile_dir set to custom, set path to dir
            if self.profile_dir == 2:
                self.profile.set_preference(
                    'browser.download.dir', 
                    self.profile_dir_path
                    )
            self.profile.set_preference(
                'browser.helperApps.neverAsk.openFile',
                'text/csv,application/vnd.ms-excel'
                )
            self.profile.set_preference(
                'browser.helperApps.neverAsk.saveToDisk',
                'text/csv,application/vnd.ms-excel'
                )

            #set options
            self.options = Options()
            if self.headless == True:
                self.options.add_argument('-headless')
                self.options.set_headless

            self.session = webdriver.Firefox(
                firefox_profile=self.profile,
                executable_path=self.driver_path,
                options=self.options
            )
            # wait for web elements to load
            self.session.implicitly_wait(10)

    def login(self, username, password):
        ''' Logs in to Ekos using credential provided by user and handles
        any alerts that may occur during log in

        PARAMS
        -----------
        session : selenium webdriver session. Initiated when class in invoked
        username : ekos ERP username
        password : ekos ERP password
        '''
        #open webdriver, go to Ekos login page
        print('Logging into Ekos')
        self.session.get('https://login.goekos.com/')
        assert 'Ekos' in self.session.title
        # enter login credentials
        elem = self.session.find_element_by_id('txtUsername')
        elem.send_keys(username)
        elem = self.session.find_element_by_id('txtPassword')
        elem.send_keys(password)
        elem.send_keys(Keys.RETURN)

        print('Login Successful')

        # session.implicitly_wait(10) #wait for page to load

        return

    def download_report(self, report_name):
        '''Clicks report name and downloads report as csv

        PARAMS
        -----------
        session : Selenium webdriver session returned by open_session function
        report_name : report to be downloaded from ekos reports page
        '''

        # session.implicitly_wait(10)

        # get reports page by url
        # print('navigating to reports page')
        # session.get('https://app.goekos.com/03.00/Report_Category') # TODO: user input url?

        # Navigate to reports page
        print('navigating to reports page')
        elem = self.session.find_element_by_xpath(
            # select 4th button in nav-options div
            "//div[@class='nav-options']/button[4]"
        )
        elem.click()
        elem = self.session.find_element_by_xpath(
            # select first link -- Report Category
            "//div[@class='nav-option--group']/a[1]")
        elem.click()

        # Reports page is Ekos Classic iFrame
        # Switch to iFrame
        print('switching to iFrame')
        self.session.switch_to.frame('classicContainer')
        
        # find link by link text
        print('opening report name: {}'.format(report_name))
        # elem = WebDriverWait(session, 5).until(
        #   EC.element_to_be_clickable((By.LINK_TEXT, report_name))
        # )
        elem = self.session.find_element_by_link_text(report_name)
        elem.click()

        # switch to iframe
        print('switching to iFrame')
        self.session.switch_to.frame('formFrame_0')

        # click export button
        print('clicking export button')
        elem = self.session.find_element_by_class_name('buttonGroupInner')
        elem.click()

        # download report as csv
        print('downloading report as csv to {}'.format(self.profile_dir_path))
        elem = self.session.find_element_by_id('csv_export')
        elem.click()

        # close iFrame
        print('switching to default content')
        self.session.switch_to.default_content()
        print('switching to iFrame')
        self.session.switch_to.frame('classicContainer')
        print('closing iFrame')
        elem=self.session.find_element_by_class_name('formClose')
        elem.click()

        return

    def quit(self):
        '''Quits session opened by open_session

        PARAMS
        ---------
        session : Selenium webdriver session returned by open_session function
        '''
        self.session.quit()
        return

    def rename_file(
        self, 
        new_filename,
        regex='Export_\d{14}_\.csv', #default ekos file name format
        PATH=None # None defaults to self.profile_dir_path
    ):
        '''Searches for the downloaded csv file based on a regular expression
        and replaces that filename with the filename provided

        PARAMS
        ---------
        new_filename : new filename for the downloaded file

        regex : regular expression to search for

        PATH : path to directory to search using regex
        '''
        if PATH == None:
            PATH = self.profile_dir_path

        regex = re.compile(regex)
        count = 0 # count to track len of os.listdir(PATH)
        for f in os.listdir(PATH):
            count += 1
            if regex.match(f) != None:
                print('File Found!')
                filename = regex.match(f).string
                os.rename(PATH+filename, PATH+new_filename)
                break
            elif regex.match(f) == None and count == len(os.listdir(PATH)):
                print('File not found')
            else:
                print('File does not match regex. Checking next file')
        return




if __name__ == '__main__':
    geckodriver='/PATH/to/geckodriver'

    username = ''
    password = ''
    report = ''
    dl_dir = ''

    ekos = EkosExport(
        'Firefox',
        geckodriver,
        2,
        dl_dir, # TODO: assert ends in /
        headless=False
    )

    ekos.login(username,password)
    ekos.download_report(report)
    ekos.rename_file('{}.csv'.format(report))
    ekos.quit()



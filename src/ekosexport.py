#!/usr/bin/env python
import logging
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

# Logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
# Handler
log_path = '' # path to log file
fh = logging.FileHandler('{}deliveries.log'.format(log_path))
fh.setLevel(logging.INFO)
# Formatter
formatter = logging.Formatter(
    '%(asctime)s : %(name)s : %(levelname)s : %(message)s'
)
fh.setFormatter(formatter)
logger.addHandler(fh)

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
            # implicit wait
            self.session.implicitly_wait(30)
            #explicit wait
            self.wait = WebDriverWait(self.session, 10)

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
        logger.info('Logging into Ekos')
        self.session.get('https://login.goekos.com/')
        assert 'Ekos' in self.session.title
        # enter login credentials
        elem = self.session.find_element_by_id('txtUsername')
        elem.send_keys(username)
        elem = self.session.find_element_by_id('txtPassword')
        elem.send_keys(password)
        elem.send_keys(Keys.RETURN)

        self.session.implicitly_wait(10)

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
        logger.info('Navigating to Reports Page')
        elem = self.wait.until(
            EC.element_to_be_clickable(
                # select 4th button in nav-options div
                (
                    By.XPATH,
                    "//div[@class='nav-options']//div[text()='Reporting']"
                )
            )
        )
        # elem.click()
        self.session.execute_script(
            "document.getElementsByClassName('nav-option nav-option--main')[4].click()"
        ) # use javascript to negate issues with selenium clicking on element

        elem = self.wait.until(
            EC.element_to_be_clickable(
                # select first link -- Report Category
                (
                    By.XPATH, 
                    "//div[@class='nav-option--group']//div[text()='All Reports']"
                )
            )
        )
        # elem.click()
        self.session.execute_script(
            "document.getElementsByClassName('nav-option nav-option--main')[0].click()"
        ) # use javascript to negate issues with selenium clicking on element

        # Reports page is Ekos Classic iFrame
        # Switch to iFrame
        logger.info('Switching to iFrame')
        self.session.switch_to.frame('classicContainer')
        
        # find link by link text
        logger.info('Opening Report name: {}'.format(report_name))
        # elem = WebDriverWait(session, 5).until(
        #   EC.element_to_be_clickable((By.LINK_TEXT, report_name))
        # )
        elem = self.wait.until(
            EC.element_to_be_clickable(
                (By.LINK_TEXT, report_name)
            )
        )
        elem.click()

        # switch to iframe
        logger.info('Switching to iFrame')
        self.session.switch_to.frame('formFrame_0')

        # click export button
        logger.info('Clicking export button')
        elem = self.wait.until(
            EC.element_to_be_clickable(
                (By.CLASS_NAME, 'buttonGroupInner')
            )
        )
        elem.click()

        # download report as csv
        logger.info('Downloading report as csv to {}'.format(self.profile_dir_path))
        elem = self.wait.until(
            EC.element_to_be_clickable(
                (By.ID, 'csv_export')
            )
        )
        elem.click()

        # close iframe
        logger.info('Closing iFrame')
        elem = self.wait.until(
            EC.element_to_be_clickable(
                (By.CLASS_NAME, 'CloseButton')
            )
        )
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
                logger.info('File Found!')
                filename = regex.match(f).string
                os.rename(PATH+filename, PATH+new_filename)
                break
            elif regex.match(f) == None and count == len(os.listdir(PATH)):
                logger.warning('File not found')
            else:
                logger.info('File does not match regex. Checking next file')
        return




if __name__ == '__main__':
    import yaml
    #Config file
    conf_file = './deliveries_config_SAMPLE.yaml' # path to config file
    stream = open(conf_file, 'r')
    config = yaml.safe_load(stream)

    geckodriver = config['driver_path']

    username = config['ekos_user']
    password = config['ekos_pw']
    report = 'Distro - This Week'
    dl_dir = config['profile_dir_path']

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



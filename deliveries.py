#!/usr/bin/env python
import yaml
import logging

from src import ekosexport
from src import googleapi

#Config file
conf_file = './deliveries_config_SAMPLE.yaml' # path to config file
stream = open(conf_file, 'r')
config = yaml.safe_load(stream)

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

# Variables
# EkosExport class
browser = 'Firefox'
driver_path = config['driver_path']
profile_dir = 2 # set custom directory
profile_dir_path = config['profile_dir_path']
headless = False

# Ekos
username = config['ekos_user']
password = config['ekos_pw']
report_name = 'Distro - This Week'

# GoogleAPI
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = config['spreadsheet_id']

# credentials
cred_path = config['cred_path']
token_path = config['token_path']

# Sheet info
DATA_RANGE_NAME = 'data!A:T'
INFO_RANGE_NAME = 'info!B1'
data = '{}{}.csv'.format(config['profile_dir_path'], report_name)

if __name__ == '__main__':
    try:
        logger.info('Instantiating ekos object')
        ekos = ekosexport.EkosExport(
            browser = browser,
            driver_path = driver_path,
            profile_dir = profile_dir,
            profile_dir_path = profile_dir_path,
            headless = headless
        )
        logger.info('Instantiating google sheets object')
        gs = googleapi.SheetsAPI(
            scopes = SCOPES,
            spreadsheet_id = SPREADSHEET_ID
        )

        logger.info('Beginning report download process')
        ekos.login(username, password)
        ekos.download_report(report_name)
        ekos.rename_file('{}.csv'.format(report_name))
        ekos.quit()

        credentials = gs.get_credentials(cred_path, token_path)
        service = gs.get_service(credentials)
        gs.import_data(
            service = service,
            data = data,
            sheet_range = DATA_RANGE_NAME
        )
        gs.last_updated(service = service, sheet_range = INFO_RANGE_NAME)
    except Exception as e:
        ekos.quit()
        logger.exception(e)


from robocorp import windows
from RPA.Desktop import Desktop
from RPA.Desktop.keywords.mouse import Action
from time import sleep
import re

from logger_config import log_failure, log_success

st = 3
yoffset = 24
top = 86

desktop = Desktop()
# app = windows.find_window('regex:Client Invoice List')

# desktop.wait_for_element('ocr:"Client Invoice List"')
# app.find('automationid:"ListRefreshBtn"').click()
# sleep(st)
project_id = 19219

while True:    
    print('review.py > Enetering while loop')
    log_success('review.py > Enetering while loop')
    # app.find('automationid:"PreviewBtn"').click()
    # sleep(st)
    preview_app = windows.find_window('regex:Preview - ')
    history_grid = preview_app.find('automationid:"InvoiceReviewGridCtrl"')
    history_grid.send_keys('{DOWN}{UP}')
    caretor = desktop.wait_for_element('alias:"caretor"')
    print(222, caretor)

    break

print('review.py > while loop finished')
log_success('review.py > Enetering while loop finished')
from robocorp import windows
from RPA.Desktop import Desktop
from RPA.Desktop.keywords.mouse import Action
from time import sleep
import re
from logger_config import log_failure, log_success  # Import the logging functions

st = 3

desktop = Desktop()

app = windows.find_window('regex:Project Command Center')
app.find('automationid:"ListRefreshBtn"').click()
sleep(st * 2)
log_success('Clicked the Refresh button on the List.')

old_description = ""
project_id = 19219
level = 0

while True:
    sleep(st)
    infotree = app.find('automationid:"InfoTree"')
    infotree.send_keys('{DOWN}')
    sleep(st * 3)      

    try:
        description_elem = app.find('automationid:"InfoInvoiceTab" > automationid:"InfoInvoiceGroupDescriptionEdit" > control:"EditControl"')        
        description = description_elem.get_value()        
        if old_description == description:
            log_success('No new description detected, exiting the loop.')
            break
        old_description = description
        
        project_number = re.search(r'LOD #(\d+)', description)
        if project_number:
            task_id = project_number.group(1)
            log_success(f'Extracted task ID: {task_id}')
        else:
            log_failure('No project number found in the description.')
            continue
        
        print(task_id)
        
        scope = app.find('automationid:"InfoInvoiceTab" > automationid:"InfoScopeEdit"')
        region = scope.rectangle
        point = f'point:{region[2]-7},{region[3]-7}'
        desktop.click(point)
        log_success('Clicked on the InfoScopeEdit region.')

        memo_elem = app.find('automationid:"MemoEdit"')
        memotext = memo_elem.get_text()
        if not memotext:
            log_failure('Memo text is empty, skipping to next iteration.')
            continue

        memolines = memotext.splitlines()
        project_name = memolines[0].split(': ')[1]
        project_number = memolines[2].split(': ')[1]
        correct_number = f'{project_id}-01-R{task_id}'
        
        if project_number != correct_number:
            memotext = memotext.replace(project_number, correct_number)
            sleep(st)            
            memo_elem.set_value(memotext)            
            sleep(st)
            memo_elem.send_keys('{ENTER}{BACK}')
            sleep(st)

            app.find('automationid:"OKButton"').click()
            sleep(st)
            app.find('automationid:"SaveBtn"').click()
            log_success(f'Updated memo text and saved changes for project ID: {project_id}.')
        else:
            log_success('Project number matches the correct number, cancelling changes.')
            sleep(st)
            app.find('automationid:"MemoFieldForm" > automationid:"CancelBtn"').click()

    except Exception as err:
        log_failure(f'An error occurred: {str(err)}')
        continue

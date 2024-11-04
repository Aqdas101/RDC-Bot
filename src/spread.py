from robocorp.tasks import task
from RPA.Cloud.Google import Google

from logger_config import log_success, log_failure

spreadsheet_id = "1sMZhEp7GnyYaCOSAb5KaezgWyOQDXWVOZUc2-fpmMmg"
sheet_name = "Sheet1"
library = Google()
service_account = "clip-new-7e3aaac5032c.json"
library.init_sheets(service_account=service_account)
log_success('spread.py > Connecting with google sheet') 
print('spread.py > Connecting with google sheet')

def insert_data(data):
    log_success('spread.py > Start: inser_data function') 
    print('spread.py > Start: inser_data function')   

    # library.clear_sheet_values(spreadsheet_id, f"{sheet_name}!A:Z")    
    result = library.insert_sheet_values(spreadsheet_id, f"{sheet_name}!A:Z", data)
    print("******************************", result)
    log_success(f'spread.py > "******************************" {result}')

    log_success('spread.py > End: inser_data function') 
    print('spread.py > End: inser_data function')

def get_data():
    log_success('spread.py > start: get_data function') 
    print('spread.py > start: get_date function')

    result = library.get_sheet_values(spreadsheet_id, f"{sheet_name}!A:Z")

    log_success('spread.py > End: get_data function') 
    print('spread.py > End: get_data function')
    return result['values']
    
if __name__ == "__main__":    
    values = [[11], [22],[33],[44],[55],[66],[77]]
    insert_data(values)
    # print(get_data())
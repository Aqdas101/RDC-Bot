from robocorp import windows
from RPA.Desktop import Desktop
from time import sleep
from spread import insert_data, get_data
from datetime import datetime
import json

from logger_config import log_success, log_failure

st = 2
depth = 20

desktop = Desktop()

app = windows.find_window('regex:Project Command Center')
print('retainer.py > Find window successfuly : regex:Project Command Center')
log_success('retainer.py > Find window successfuly : regex:Project Command Center')

sheet_data = get_data()


with open("employee-email.json", "r") as read_file:
    email_data = json.load(read_file)

print('retainer.py > Opening "employee-email.json" file')
log_success('retainer.py > Opening "employee-email.json" file')

# app.find('automationid:"ListRefreshBtn"').click()
# sleep(st * 2)

result = [['Project Name'], ['Project ID'], ['Billing Group Name'], ['Retainer Amount'], ['Request Date'], ['Project Manager'], ['Principal in Charge']]
old_info_text = ''
old_list_text = ''

prev_left = 14
prev_right = 500
listgrid = app.find('automationid:"ListGridControl"')
prev_list_top = listgrid.top 
infogrid = app.find('automationid:"InfoTree"')
prev_info_top = infogrid.top

def click_tab():
    print('retainer.py > Start: Click Tab Function')
    log_success('retainer.py > Start: Click Tab Function')

    tab = desktop.find_elements('alias:"dates_costs_selected"')    
    if len(tab):
        return True
    else:
        tab = desktop.find_elements('alias:"dates_costs_unselected"')    
        if len(tab):
            tab = tab[0]
            desktop.click(tab)            
            new_region = desktop.define_region(tab.left, tab.top - 100, tab.right, tab.bottom)
            desktop.move_mouse(new_region)
            sleep(1)
            return True
        
    print('retainer.py > End: Click Tab Function\n\n')
    log_success('retainer.py > End: Click Tab Function\n\n')

while True:  
    try:
        print('retainer.py > Entering While loop')
        log_success('retainer.py > Entering While loop')

        sleep(st)
        # history_grid = app.find('automationid:"ListGridControl"')    
        caretors = desktop.find_elements('alias:"caretor"')
        print('retainer.py > Finding element alias:"caretor"')
        log_success('retainer.py > Finding element alias:"caretor"')

        if len(caretors) < 2:
            # print("not found caretor")
            continue  
        caretor = caretors[1]
        left = caretor.left
        top = caretor.top
        right = caretor.right
        bottom = caretor.bottom
        
        info_region = desktop.define_region(prev_left, prev_info_top, prev_right, bottom)    
        info_text = desktop.read_text(info_region)
        
        if info_text == old_info_text:
            list_region = desktop.define_region(prev_left, caretors[0].top, prev_right, caretors[0].bottom)    
            list_text = desktop.read_text(list_region)
            if list_text == old_list_text:
                break
            else:
                old_list_text = list_text
                listgrid = app.find('automationid:"ListGridControl"')
                listgrid.send_keys('{DOWN}')
                sleep(st * 2)
                continue
        else:
            old_info_text = info_text
        
        is_dates_costs = click_tab()
        
        if is_dates_costs:
            request_elem = desktop.find_elements('alias:"retainer_requested"')
            if len(request_elem):
                checked = True
            else:
                checked = False            
            
            receive_elem = app.find('control:"EditControl" and name:"Retainer Received"', depth)
            received = receive_elem.get_value()    

            note_elem = app.find('control:"EditControl" and name:"Notes"', depth)
            note = note_elem.get_value()        

            if checked and not received:
                list_region = desktop.define_region(prev_left, caretors[0].top, prev_right, caretors[0].bottom)    
                list_text = desktop.read_text(list_region)
                tab = desktop.find_elements('alias:"general"')    
                if len(tab):
                    tab = tab[0]
                    desktop.click(tab)
                    group_text = app.find('control:"EditControl" and name:"Description"', depth).get_value()
                    group_id = app.find('control:"EditControl" and name:"ID"', depth).get_value()
                    pm = app.find('control:"EditControl" and name:"Project manager"', depth).get_value()
                    pic = app.find('control:"EditControl" and name:"Principal in charge"', depth).get_value()
                    datestr = datetime.today().strftime('%m/%d/%Y')
                    # if sheet_data:
                    #     for item in sheet_data:
                    #         try:
                    #             if item[0] == list_text and item[1] == group_id and item[2] == group_text and item[4]:
                    #                 datestr = item[4]
                    #                 break
                    #         except Exception as err:
                    #             pass
                    
                    result[0].append(list_text)
                    result[1].append(group_id)            
                    result[2].append(group_text)
                    result[3].append(note)
                    result[4].append(datestr)
                    result[5].append(pm)
                    result[6].append(pic)
                    insert_data([[list_text], [group_id], [group_text], [note], [datestr], [pm], [pic]])
                    print(list_text, "==", group_id, "==", group_text, "==", note, "==", datestr, "==", pm, "==", pic)
        # else:        
        #     print('error')  

        infogrid = app.find('automationid:"InfoTree"')
        infogrid.send_keys('{DOWN}')
    except Exception as e:
        error_message = str(e)
        log_failure(f'retainer.py > Failed to initiate Edge driver: reason {error_message.splitlines()[0]}')


print('retainer.py > While loop Finished')
log_success('retainer.py > While loop Finished\n\n')
# insert_data(result)

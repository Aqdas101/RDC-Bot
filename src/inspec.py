from RPA.Desktop.Windows import Windows
from RPA.Desktop import Desktop
from RPA.Desktop.keywords.mouse import Action
from time import sleep

from logger_config import log_success, log_failure

st = 2
x = 18
y = 85
yoffset = 24

try:
    win = Windows()
    desktop = Desktop()
    log_success('inspec.py --- Waiting for opening window > ocr:"Preview - 19219" \n\n')
    desktop.wait_for_element('ocr:"Preview - 19219"')
    log_success('inspec.py --- Detect window > ocr:"Preview - 19219" ')

    region = desktop.define_region(27,110,1127,320)
    text = desktop.read_text(region)
    log_success('inspec.py --- Read text from define region')
    print(text)
    log_success(f'inspec.py --- text extract from specified region : {text}')
    sleep(st)
    log_success(f'inspec.py --- sleep for {st} seconds')
except Exception as e:
    error_message = str(e)
    log_failure(f'inspec.py > Failed extracting text from window, reason: {error_message.splitlines()[0]} ')


try:
    desktop.click(region, action=Action.right_click)
    log_success('inspec.py --- Click on the specified region\n\n')
except Exception as e:
    log_failure('inspec.py --- Failed to click on specified region\n\n')


                
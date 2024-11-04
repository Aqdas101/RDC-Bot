from robocorp import browser
from robocorp.tasks import task
from RPA.Excel.Files import Files as Excel
from time import sleep
from robocorp import windows
from RPA.Desktop import Desktop

from logger_config import log_failure, log_success

USER = "aim@rdc-s111.com"
PASS = "6!UM^/t5"
st = 2

@task
def solve_challenge():
    """
    Main task which solves the RPA challenge!

    Downloads the source data Excel file and uses Playwright to fill the entries inside
    rpachallenge.com.
    """
    browser.configure(
        browser_engine="msedge",
        screenshot="only-on-failure",
        headless=False,
        slowmo=1000
    )
    try:
        page = browser.goto("https://ajera.prarchitects.com/ajera")
        log_success('Navigated to the login page.')

        page.fill(".ax-login-username > .ax-login-input-field", USER)
        log_success('Filled in the username.')
        
        page.fill(".ax-login-password > .ax-login-input-field", PASS)
        log_success('Filled in the password.')

        page.get_by_text("Login").click()
        log_success('Clicked the login button.')
        sleep(st)

        page.click("img.ax-main-header-interop-button")
        log_success('Clicked the main header interop button.')
        sleep(st)

        page.click("li:text('Client Invoices')")
        log_success('Clicked on Client Invoices.')
        sleep(st)

        page.on("dialog", lambda dialog: dialog.accept())
        log_success('Accepted dialog if present.')

        page.get_by_role("button", name="Open").click()
        log_success('Clicked the Open button.')
        sleep(1000)

        app = windows.find_window('regex:.*Client Invoice List').inspect()
        sleep(st)
        app.find('name:"Maximize"').click()
        log_success('Maximized the Client Invoice List window.')
        sleep(st)

        app.find('automationid:"ViewBtn"').click()
        log_success('Clicked on the View button.')
        sleep(st)

        app.find('name:"2024" and control:"ListItemControl"').click()
        log_success('Selected the year 2024.')
        sleep(st)

        app.find('name:"Final" and control:"CheckBoxControl"').click()
        log_success('Checked the Final checkbox.')
        sleep(st)

        app.find('control:"EditControl" and name:"Project"').set_value("Regency Centers On Call LOD's")
        log_success('Set the Project name.')
        sleep(st)

        app.find('automationid:"OKBtn"').click()
        log_success('Clicked the OK button.')
        sleep(st)

        app.find('automationid:"ListGrid" > automationid:"3475390" and path:"13|2"').select()
        log_success('Selected the desired item from the ListGrid.')

        desktop = Desktop()
        desktop.wait_for_element('ocr:"Client Invoice List"')
        desktop.click('ocr:"2121 Crystal"')
        log_success('Clicked on the specified Client Invoice.')

        print("////////////////////////////////////")
        print(result)

    except Exception as e:
        log_failure(f'An error occurred: {str(e)}')
    
    finally:
        # A place for teardown and cleanups. (Playwright handles browser closing)
        print("Automation finished!")

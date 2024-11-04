from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
import json

from logger_config import log_success, log_failure


USER="aim@rdc-s111.com"
PASS="6!UM^/t5"
st = 4

def get_emails():
    log_success('employee.py > Start get_emails function')

    # To save extract table data
    file_name = "../employee-email.json"

    # Path to your Edge WebDriver executable
    webdriver_path = "../msedgedriver.exe"

    # Path to the existing Edge profile
    # profile_path = r"C:/Users/AiM02/AppData/Local/Microsoft/Edge/User Data"

    # Initialize Edge options
    edge_options = Options()

    # Set the path to the user data directory (profile path)
    # edge_options.add_argument(f"user-data-dir={profile_path}")

    # You can specify a particular profile if needed (e.g., Profile 1)
    # edge_options.add_argument("profile-directory=Profile 1")

    # Initialize the WebDriver with the options
    try:
        service = Service(webdriver_path)
        driver = webdriver.Edge(service=service, options=edge_options)
        log_success('employee.py--- Setting up Driver')
        driver.get("https://login.ajera.com")
        log_success("Employee.py --- Navigated to https://login.ajera.com")
        sleep(st)
    except Exception as e:
        error_message = str(e)
        log_failure(f'Employee.py > Failed to initiate Edge driver: reason {error_message.splitlines()[0]}')


    # Open a website for testing
    
    try:
        log_success("Employee.py --- Finding username and password field")
        try:
            username_field = driver.find_element(By.CSS_SELECTOR, ".ax-login-username > .ax-login-input-field")  # Replace with the actual ID or selector
            password_field = driver.find_element(By.CSS_SELECTOR, ".ax-login-password > .ax-login-input-field")  # Replace with the actual ID or selector
        except:
            log_failure("Employee.py --- Failed to find username and password field")

        if username_field and password_field:
            # Enter credentials (replace with actual username and password)
            # username_field.send_keys(USER)
            # password_field.send_keys(PASS)

            # Submit the login form (you can also use the submit button)
            try:
                password_field.send_keys(Keys.RETURN)
                log_success("Employee.py --- Hit enter to submit")
            except:
                log_failure('Employee.py --- Failed to submit form')

            # Wait for the login process to complete and the next page to load
            sleep(st)
            log_success(f"Employee.py --- sleeping for {st} seconds")
            
            # Find the table by its ID
            table = driver.find_element(By.CSS_SELECTOR, "#ax-axgrid-element-1 table tbody")  # Replace with the actual table ID
            log_success("Employee.py --- Extract table CSS Selector")

            # Extract all the rows from the table (assuming it's a <table> element)
            rows = table.find_elements(By.TAG_NAME, "tr")
            log_success("Employee.py --- Extract all rows elements of table")

            # Loop through each row and extract the cell data
            result = []
            log_success("Employee.py --- Start -- Processessing rows")
            for row in rows:
                # Extract all the cells (td) in the row
                
                cells = row.find_elements(By.TAG_NAME, "td")
                name = cells[0].text
                email = cells[1].text
                
                result.append({
                    "name": name,
                    "email": email
                })
            log_success(f"Employee.py --- End process rows")
            
            with open(file_name, "w") as write_file:
                json.dump(result, write_file, indent=4)

            log_success(f"Employee.py --- Saved data with filename: {file_name}")
            
            return result
    
    except:
        log_failure("Employee.py --- Failed to extract table data from webiste")
        username_field = None
        password_field = None
        pass

    finally:
        if driver:
            driver.quit()
            log_success("Employee.py --- Driver has been quit successfully \n\n")

if __name__ == "__main__":
    print(get_emails())


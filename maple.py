# Make sure to install chromedriver first before running this script.
# https://googlechromelabs.github.io/chrome-for-testing/
import re
import subprocess
import sys

def checkModules():
    try:
        import selenium
    except ImportError:
        print('Selenium is not installed. Installing...')
        subprocess.check_call([sys.executable, "-m", "pip", "install", "selenium"])
    try:
        import barnum
    except ImportError:
        print('Barnum is not installed. Installing...')
        subprocess.check_call([sys.executable, "-m", "pip", "install", "barnum"])
    try:
        import names
    except ImportError:
        print('Names is not installed. Installing...')
        subprocess.check_call([sys.executable, "-m", "pip", "install", "names"])

checkModules()

# Re-import the modules to ensure they are available
import selenium
import barnum
import names

if 'selenium' in sys.modules and 'barnum' in sys.modules and 'names' in sys.modules:
    print('\n\033[1;31mAll modules have been installed.\nPlease restart your VS Code IF it is newly installed.\033[0m\n')

# Dependencies needed: Selenium & barnum
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from barnum import gen_data
import names
import random
import string

def makeCode():
    # Set path to chromedriver executable - EDIT HERE
    service = Service("PATH TO CHROME DRIVER")
    pondev = webdriver.Chrome(service=service)
    
    # Creating random email
    letters = string.ascii_lowercase
    mailRand = ''.join(random.choice(letters) for i in range(6)) + '@mailsac.com'
    urlreg = 'https://www.maplesoft.com/products/maple/free-trial/'
    getlink = 'https://mailsac.com/inbox/'

    try:
        # Go to the registration page
        pondev.get(urlreg)

        # Enter the email address
        try:
            print("Entering email address...")
            pondev.find_element(By.ID, 'EmailAddress').send_keys(mailRand)
            print("\033[1;33mEmail address entered successfully.\nEmail: " + mailRand + "\n\033[0m")
            
            # Wait for the button to be clickable and ensure it is visible
            print("Waiting for 'Get your Free Trial' button...")
            button = WebDriverWait(pondev, 20).until(EC.visibility_of_element_located((By.ID, 'btnSubmitEmail')))
            pondev.execute_script("arguments[0].scrollIntoView(true);", button)  # Scroll into view if needed
            print("Clicking the 'Get your Free Trial' button...")
            button.click()
        except TimeoutException:
            print("The 'Get your Free Trial' button was not clickable in time.")
            raise
        except ElementNotInteractableException as e:
            print(f"Error interacting with the 'Get your Free Trial' button: {e}")
            raise

        # Wait for the next page to load completely
        try:
            print("Waiting for the FirstName field to be visible...")
            WebDriverWait(pondev, 20).until(EC.presence_of_element_located((By.ID, 'FirstName')))
            print("Next page loaded.")
        except TimeoutException:
            print("The next page did not load in time.")
            raise

        # Scroll the FirstName field into view to ensure it's visible
        try:
            print("Scrolling 'FirstName' field into view...")
            first_name_field = pondev.find_element(By.ID, 'FirstName')
            pondev.execute_script("arguments[0].scrollIntoView(true);", first_name_field)
            WebDriverWait(pondev, 10).until(EC.element_to_be_clickable((By.ID, 'FirstName')))
            print("FirstName field is visible and ready to interact.")
        except NoSuchElementException as e:
            print(f"FirstName field not found: {e}")
            raise
        except ElementNotInteractableException as e:
            print(f"Error interacting with the 'FirstName' field: {e}")
            raise

        # Fill in the form
        try:
            print("Filling out the form...")
            first_name_field.send_keys(names.get_first_name())
            pondev.find_element(By.ID, 'LastName').send_keys(names.get_last_name())
            pondev.find_element(By.ID, 'Company').send_keys(gen_data.create_company_name())
            pondev.find_element(By.ID, 'JobTitle').send_keys(gen_data.create_job_title())
            print("Form filled successfully.")
        except NoSuchElementException as e:
            print(f"Form element not found: {e}")
            raise
        except ElementNotInteractableException as e:
            print(f"Error interacting with a form element: {e}")
            raise

        # Select country from the dropdown
        try:
            print("Selecting country from dropdown...")
            country_drop_down = Select(pondev.find_element(By.ID, 'CountryDropDownList'))
            country_drop_down.select_by_visible_text("United States")
        except NoSuchElementException:
            print("Country dropdown was not found.")
            raise
        except ElementNotInteractableException as e:
            print(f"Error interacting with the country dropdown: {e}")
            raise

        # Wait for region dropdown to load
        try:
            print("Waiting for region dropdown to load...")
            WebDriverWait(pondev, 10).until(EC.presence_of_element_located((By.XPATH, "//option[@value='CA  ']")))
            print("Region dropdown loaded.")
        except TimeoutException:
            print("Region dropdown did not load in time.")
            raise

        # Select region
        try:
            print("Selecting region...")
            pondev.find_element(By.XPATH, "//option[@value='CA  ']").click()
        except NoSuchElementException:
            print("Region option was not found.")
            raise
        except ElementNotInteractableException as e:
            print(f"Error interacting with the region dropdown: {e}")
            raise

        # Select "Student" option from dropdown
        try:
            print("Selecting 'Student' option from dropdown...")
            segment_dropdown = Select(pondev.find_element(By.ID, 'ddlSegment'))
            segment_dropdown.select_by_visible_text("Student")
        except NoSuchElementException:
            print("'Student' option was not found in the dropdown.")
            raise
        except ElementNotInteractableException as e:
            print(f"Error interacting with the 'Student' dropdown: {e}")
            raise


        # Agree to GDPR checkbox
        try:
            print("Agreeing to GDPR...")
            pondev.find_element(By.ID, 'chkAgreeToGDPR').click()
        except NoSuchElementException:
            print("GDPR checkbox was not found.")
            raise
        except ElementNotInteractableException as e:
            print(f"Error interacting with the GDPR checkbox: {e}")
            raise

        # Submit the form
        try:
            print("Submitting the form...")
            pondev.find_element(By.ID, 'SubmitButton').click()
        except NoSuchElementException:
            print("Submit button was not found.")
            raise
        except ElementNotInteractableException as e:
            print(f"Error interacting with the 'Submit' button: {e}")
            raise

        # Wait for the next step (e.g., URL change or form success message)
        try:
            print("Waiting for URL to change after form submission...")
            WebDriverWait(pondev, 20).until(EC.url_changes(urlreg))
            print("Form submission successful. URL changed.")
        except TimeoutException:
            print("The URL did not change after form submission.")
            raise

        # Now, navigate to the email inbox to fetch the confirmation email
        pondev.get(getlink + mailRand)

        # Wait for the email to arrive (check for the presence of the email in the inbox)
        try:
            print("Waiting for email to arrive in inbox...")
            WebDriverWait(pondev, 20).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[3]/div[1]/div/div[2]/div/table/tbody/tr[2]/td[3]")))
            print("\n\033[1;32mEMAIL RECEIVED!\033[0m\n")
        except TimeoutException:
            print("The confirmation email did not arrive in time.")
            raise

        # Open the latest email
        try:
            print("Opening the latest email...")
            email_entry = pondev.find_element(By.XPATH, "/html/body/div[1]/div[3]/div[1]/div/div[2]/div/table/tbody/tr[2]/td[3]")
            email_entry.click()
        except NoSuchElementException:
            print("Email entry was not found in the inbox.")
            raise
        except ElementNotInteractableException as e:
            print(f"Error interacting with the email in inbox: {e}")
            raise

        # Retrieve and open the activation link, also print the content of the email
        try:
            print("Retrieving the email content...")
            
            # Get the full email content as text or HTML
            email_content = pondev.find_element(By.XPATH, "/html/body/div/div[3]/div[1]/div/div[2]/div/table/tbody/tr[2]/td[2]/div[2]").get_attribute("innerHTML")
            
            # Print the email content to inspect it
            print(f"Email content: {email_content}")
            
            # Use regex to search for the activation link
            activation_link = re.search(r'https://www\.maplesoft\.com/InstantEvalConfirmation/[^\s"<]+', email_content)
            
            if activation_link:
                activation_url = activation_link.group(0)
                print(f"\n\033[1;32mActivation Link: {activation_url}\033[0m\n")
                
                # Navigate to the activation URL
                pondev.get(activation_url)
            else:
                print("No activation link found in the email content.")
            
        except NoSuchElementException:
            print("Email content was not found.")
            raise
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            raise

        # Wait for the link to open and load properly
        try:
            print("Waiting for the activation page to load...")

            # Increase timeout and wait for the page to load completely
            WebDriverWait(pondev, 30).until(
                EC.visibility_of_element_located((By.XPATH, "//h3[contains(text(), 'Download')]"))
            )
            
            print("\n\033[1;32mDETAILS FOUND!\033[0m")

            # Use a more relative XPATH to locate the download link
            first_install = pondev.find_element(By.XPATH, "/html/body/div[5]/div/div[2]/p[3]/span/a").get_attribute("href")
            exp = pondev.find_element(By.ID, 'evaluationExpiry').text
            a_code = pondev.find_element(By.ID, 'evaluationPurchaseCode').text

            print(f"DOWNLOAD LINK: {first_install}\nACTIVATION CODE: {a_code}\nEXPIRY: {exp}")

        except TimeoutException:
            print("The activation page did not load in time.")
            pondev.save_screenshot('activation_page_timeout.png')
            print("Screenshot saved as 'activation_page_timeout.png'")
            raise

        except NoSuchElementException as e:
            print(f"Download link or activation details not found: {e}")
            pondev.save_screenshot('element_not_found.png')
            print("Screenshot saved as 'element_not_found.png'")
            raise

    except TimeoutException:
        print("An element was not found or did not load in time.")
    except ElementNotInteractableException as e:
        print(f"Element not interactable: {e}")
    except NoSuchElementException as e:
        print(f"Element not found: {e}")
    finally:
        pondev.quit()

makeCode()

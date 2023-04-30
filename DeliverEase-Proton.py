from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import time
import argparse
import pyfiglet

def send_proton_email(proto_user, email_to, email_subject, email_message, counter):
    url_proto = "https://account.proton.me/login"
    proto_login = proto_user
    proto_pass = ""

    driver.get(url_proto)
    time.sleep(7)

    print("\n[+] Logging into Protonmail", end='\r')

    driver.find_element(By.ID, "username").send_keys(proto_login)
    driver.find_element(By.ID, "password").send_keys(proto_pass)
    driver.switch_to.active_element.send_keys('\t' + '\t' + '\t' + '\t' + '\n')
    #Alternate method to click Login Button: driver.find_element(By.CLASS_NAME, "button-solid-norm").click()
    
    time.sleep(20)

    print("[+] Composing emails for {0}.".format(email_to))

    i = 1
    tots_sent = counter

    while i <= counter:

        try:

            # Click Compose Email
            driver.find_element(By.CSS_SELECTOR, 'button[data-testid="sidebar:compose"]').click()

            time.sleep(3)

            # Fill 'To' & 'Subject' Field
            driver.switch_to.active_element.send_keys(email_to + '\n' + '\t' + email_subject + '\t' + email_message)
            time.sleep(3)

            # Fill Email Body
            driver.switch_to.active_element.send_keys('\t' + '\t' + '\t' + '\t' + '\t' + '\t' + '\t' + '\t' + '\t' + '\t' + '\t' + '\t' + '\t' + '\t' + '\t')
            time.sleep(3)

            # Send Email
            driver.find_element(By.CSS_SELECTOR, "button.button.button-group-item.button-solid-norm.composer-send-button[data-testid='composer:send-button']").click()
        
            print("[+] Counter: {0}".format(i), end='\r')
        
        except Exception as e:
            print("[!] Error on counter {0}: {0}".format(i, e))
            tots_sent -=1

        i +=1
        time.sleep(10)

    print("[+] {0} emails completed".format(tots_sent))

    driver.find_element(By.CSS_SELECTOR, 'button[data-testid="heading:userdropdown"]').click()
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR, 'button[data-testid="userdropdown:button:logout"]').click()
    time.sleep(5)
    print("[+] Logged out successfully", end='\r')
    


# Main
prebanner = pyfiglet.figlet_format("AutoProton")
banner = prebanner + "\n-- @FireStone65 -- \n\n"
print(banner)

parser = argparse.ArgumentParser(description='[+] Automate sending Proton emails with Selenium')
parser.add_argument('-t', type=str, required=True, help='Target Email ID')
parser.add_argument('-u', type=str, required=True, help='Proton Email ID')
parser.add_argument('-x', type=int, required=True, help='No. of Emails to Send')
args = parser.parse_args()

# Initialize Chrome driver
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


# Email Information
email_to = args.t  
email_subject = "Warmup: Sent Using Selenium"
email_message = "Hello, \n\nThis email has been completely automated using Selenium and Python\n\nBest wishes,\nFireStone65"
send_volume = args.x
proto_login = args.u

send_proton_email(proto_login, email_to, email_subject, email_message, send_volume) 

print("\n[+] Automation complete", end='\r')
inp = input('\n ---- Hit any key to quit')

driver.quit()

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import sys
import time
import argparse
import pyfiglet


def send_owa_email(outlook_user, outlook_pass, email_list, email_subject, email_message, counter, flag):
    url_owa = "https://outlook.office.com/mail"
    owa_login = outlook_user
    owa_pass = outlook_pass

    print("[+] Visiting OWA Portal", end='\r')
    driver.get(url_owa)
    time.sleep(7)

    print("\n[+] Logging into Outlook", end='\r')
    
    driver.find_element(By.ID, "i0116").send_keys(owa_login)
    driver.find_element(By.ID, "idSIButton9").click()
    time.sleep(3)

    driver.find_element(By.ID, "i0118").send_keys(owa_pass)
    driver.find_element(By.ID, "idSIButton9").click()
    time.sleep(3)

    driver.find_element(By.ID, "idBtn_Back").click()
    time.sleep(7)

    print("[+] Authenticated successfully", end='\r')
    
    if flag == 0:
        print("[+] Composing emails for {0}".format(email_list[0]))

    i = 1
    tots_sent = counter

    while i <= counter:

        try:

            driver.find_element(By.CLASS_NAME, "label-186").click()
            time.sleep(3)

            if flag == 1:
                email_to = email_list[i-1]
                print("[+] Composing emails for {0}".format(email_to))

            else:
                email_to = email_list[0]
		
            driver.switch_to.active_element.send_keys(email_to + '\n' + '\n' + '\t' + '\t' + email_subject + '\t' + email_message)
            time.sleep(3)

            driver.find_element(By.XPATH, "//button[@title='Send (Ctrl+Enter)' and contains(@class, 'ms-Button--primary')]").click()
            time.sleep(3)

            print("[+] Counter: {0}".format(i), end='\r')

        except Exception as e:

             print("[!] Error on Counter {0}: {0}".format(i, e))
             tots_sent -=1

        i += 1
        time.sleep(8)

    print("[+] {0} emails completed".format(tots_sent))
    print("[+} Logging out..", end='\r')    

    driver.find_element(By.ID, "O365_MainLink_Me").click()
    time.sleep(3)

    driver.find_element(By.ID, "mectrl_body_signOut").click()
    time.sleep(3)

    print("[+] Logged out successfully", end='\r')
    time.sleep(1)

# Main
prebanner = pyfiglet.figlet_format("AutoOutlook")
banner = prebanner + "\n-- @FireStone65 -- \n\n"
print(banner)

parser = argparse.ArgumentParser(description='[+] Automate sending Outlook emails with Selenium')
parser.add_argument('-t', type=str, required=False, help='Single Target Email ID')
parser.add_argument('-u', type=str, required=True, help='Sender Outlook Email ID')
parser.add_argument('-p', type=str, required=True, help='Sender Outlook Email Password')
parser.add_argument('-x', type=int, required=False, help='No. of Emails to Send (applicable only for single targets)')
parser.add_argument('-w', type=str, required=False, help='Multiple Targets Email Wordlist')

args = parser.parse_args()

# Initialize Chrome driver
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Distinguish between single vs. multiple recipients
flag = 0

#Email Information
email_list = list()

if args.w != None:
    emails_txt = args.w

    with open(emails_txt, 'r') as fp:
        
        flag = 1
        email_list = [line.rstrip('\n') for line in fp.readlines()]
        send_volume = len(email_list)


elif args.t != None:
    email_list.append(args.t)
    
    if args.x != None:
        send_volume = args.x
    else:
        print("[+] Send volume not provided. Defaulting to 1 email")
        send_volume = 1

else:
    print("[!] Required fields: Target Email ID / Target Email Wordlist")
    sys.exit()

email_subject = "Warmup: Sent Using Selenium"
email_message = "Hello, \n\nThis email has been completely automated to reach out using Selenium and Python.\n\nBest wishes,\nFirestone65"

outlook_login = args.u
outlook_pass = args.p

send_owa_email(outlook_login, outlook_pass, email_list, email_subject, email_message, send_volume, flag) 

print("\n[+] Automation complete", end='\r')
inp = input('\n ---- Hit any key to quit')

driver.quit()

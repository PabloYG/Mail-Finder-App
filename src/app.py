import spreadsheet
import spreadsheetOUTPUT
from selenium import webdriver
import time

# Extracts contact data from Google Sheet
def sheet_decoder():

    row_count = spreadsheet.sheet2.row_count
    
    for i in range(2, row_count):
        row_values = spreadsheet.sheet2.row_values(i)
        try:
            name = row_values.__getitem__(0)
            surname = row_values.__getitem__(1)
            company = row_values.__getitem__(2)
            email_sufix = row_values.__getitem__(3)
            gender = row_values.__getitem__(4)
        except IndexError:
            exit()
        mail_combination(name, surname, company, email_sufix, gender)


# Creates all combinations of contact data text for potential email addresses 
def mail_combination(name, surname, company, email_sufix, gender):

    # Creates array of potential addresses
    mail_list = []

    
    # Guess 1
    mail_guess = name + email_sufix
    mail_list.append(mail_guess)

    # Guess 2
    mail_guess = surname + email_sufix
    mail_list.append(mail_guess)

    # Guess 3
    mail_guess = name + surname + email_sufix
    mail_list.append(mail_guess)

    # Guess 4
    mail_guess = name + "." + surname + email_sufix
    mail_list.append(mail_guess)

    # Guess 5
    mail_guess = name[:1] + surname + email_sufix
    mail_list.append(mail_guess)

    # Guess 6
    mail_guess = name[:1] + "." + surname + email_sufix
    mail_list.append(mail_guess)

    # Guess 7
    mail_guess = name + surname[:1] + email_sufix
    mail_list.append(mail_guess)

    # Guess 8
    mail_guess = name + "." + surname[:1] + email_sufix
    mail_list.append(mail_guess)

    # Guess 9
    mail_guess = name[:1] + surname[:1] + email_sufix
    mail_list.append(mail_guess)

    # Guess 10
    mail_guess = surname + name + email_sufix
    mail_list.append(mail_guess)

    # Guess 11
    mail_guess = surname + "." + name + email_sufix
    mail_list.append(mail_guess)

    # Guess 12
    mail_guess = surname + name[:1] + email_sufix
    mail_list.append(mail_guess)

    # Guess 13
    mail_guess = surname + "." + name[:1] + email_sufix
    mail_list.append(mail_guess)

    # Guess 14
    mail_guess = surname[:1] + name + email_sufix
    mail_list.append(mail_guess)

    # Guess 15
    mail_guess = surname[:1] + "." + name + email_sufix
    mail_list.append(mail_guess)

    # Guess 16
    mail_guess = surname[:1] + name[:1] + email_sufix
    mail_list.append(mail_guess)

    # Guess 17
    mail_guess = surname[:1] + "." + name[:1] + email_sufix
    mail_list.append(mail_guess)

    # Guess 18
    mail_guess = name + ".." + surname + email_sufix
    mail_list.append(mail_guess)

   
    # Invokes web scraper to query MailTester
    mail_tester_post(mail_list, name, surname, company, gender)

# Initializes web driver and query loop to MailTester
def mail_tester_post(mail_list, name, surname, company, gender):

    # Opens Chrome Browser
    driver = webdriver.Chrome()

    # Opens Firefox Browser (deprecated)
    # driver = webdriver.Firefox()
    
    # Inputs data in order
    entry_loop(mail_list, name, surname, company, gender, driver)
    
    # Closes driver after each session
    driver.close()

# Loops over all potential addresses to find real ones
def entry_loop(mail_list, name, surname, company, gender, driver):
   
    for i in range(0, len(mail_list)):

        driver.get("http://mailtester.com/")
        input_element = driver.find_element_by_name("email")
        email = mail_list[i]
        input_element.send_keys(email)
        input_element.submit()
        submitted = True
        t0 = time.time()

        while submitted:

            if driver.page_source.__contains__("E-mail address is valid"):
                print(email)
                sheet_encoder(email, name, surname, company, gender)
                submitted = False
            elif driver.page_source.__contains__("E-mail address does not exist on this server"):
                submitted = False
            elif driver.page_source.__contains__("Server doesn't allow e-mail address verification"):
                return
            elif driver.page_source.__contains__("Invalid mail domain."):
                return

            loop_time = time.time() - t0
            if loop_time > 60:
                return

# Saves real addresses in Google Sheet
def sheet_encoder(email, name, surname, company, gender):

    row_values = spreadsheetOUTPUT.sheet1.col_values(1)
    row_count = len(row_values)

    fname = name[:1].upper() + name[1:]
    lname = surname[:1].upper() + surname[1:]

    row = [fname, lname, company, email, gender]
    spreadsheetOUTPUT.sheet1.insert_row(row, row_count + 1)


sheet_decoder()

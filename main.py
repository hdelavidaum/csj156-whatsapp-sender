import os
import time
import pyautogui
import pandas as pd

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

load_dotenv()

message_hello = os.getenv('MESSAGE_HELLO')
message1 = os.getenv('MESSAGE_1')
message2 = os.getenv('MESSAGE_2')

browser = webdriver.Chrome()
browser.get("https://web.whatsapp.com/")
while len(browser.find_elements(By.ID, "side")) < 1:
    time.sleep(10)

contatos_df = pd.read_excel("data/contacts.xlsx")
for i, name in enumerate(contatos_df['nome']):
    print(f"Starting if with element #{i}")
    contact_id = i
    contact_phone = contatos_df.loc[i, 'numero']
    contact_name = contatos_df.loc[i, 'nome']
    contact_file = contatos_df.loc[i, 'ficha_glp']
    print(f"contact_file --> {contact_file}")

    # search for contact since all of them is listed
    browser.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p').send_keys(str(contact_phone))
    time.sleep(5)

    # click on the first match given should have only one
    browser.find_element(By.XPATH, '//*[@id="pane-side"]/div[1]/div/div/div/div/div').click()
    time.sleep(5)
    
    # write the messages
    print('writing messages')
    browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(message_hello.format(contact_name))
    browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(Keys.ENTER)
    time.sleep(10)
    
    browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p')
    for char in message1:
        browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(char)
        # time.sleep(0.1)
    browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(Keys.ENTER)
    time.sleep(7)
    
    browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p')
    for char in message2:
        browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(char)
        # time.sleep(0.1)
    browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(Keys.ENTER)
    time.sleep(7)

    print(os.path.abspath(f"data\\{contact_file}"))
    # find button to add attachment and send it
    browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div').click()
    time.sleep(2)

    browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[1]/li/div/span').click()
    time.sleep(3)

    pyautogui.write(os.path.abspath(f"data\\{contact_file}"))
    pyautogui.press('enter')
    time.sleep(5)

    browser.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').click()

    # waiting to end it all
    counter = 10
    while counter > 0:
        print(counter)
        time.sleep(1)
        counter -= 1

print("SUCCESS!")
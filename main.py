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
message3 = os.getenv('MESSAGE_3')
message4 = os.getenv('MESSAGE_4')
message5 = os.getenv('MESSAGE_5')
whatsapp_contact_url = "https://web.whatsapp.com/send?phone={}"

browser = webdriver.Chrome()
browser.get("https://web.whatsapp.com/")
while len(browser.find_elements(By.ID, "side")) < 1:
    time.sleep(10)

xpath_write_message = '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p'

# contatos_df = pd.read_excel("data/contacts.xlsx")
contatos_df = pd.read_excel("data/contacts-test.xlsx")
for i, name in enumerate(contatos_df['nome']):
    print(f"Starting if with element #{i}")
    contact_id = i
    contact_user_num = contatos_df.loc[i, 'cadastro']
    contact_phone = contatos_df.loc[i, 'numero']
    contact_name = contatos_df.loc[i, 'nome']
    contact_file = contatos_df.loc[i, 'extrato']

    print(f"Going to {whatsapp_contact_url.format(contact_phone)}")
    browser.get(whatsapp_contact_url.format(contact_phone))
    counter = 45
    while counter > 0:
        print(counter)
        time.sleep(1)
        counter -= 1
    
    # write the messages
    print('writing messages')
    browser.find_element(By.XPATH, xpath_write_message).send_keys(message_hello.format(contact_name))
    browser.find_element(By.XPATH, xpath_write_message).send_keys(Keys.ENTER)
    time.sleep(2)
    
    browser.find_element(By.XPATH, xpath_write_message).send_keys(message1)
    browser.find_element(By.XPATH, xpath_write_message).send_keys(Keys.ENTER)
    time.sleep(2)
    
    browser.find_element(By.XPATH, xpath_write_message).send_keys(message2)
    browser.find_element(By.XPATH, xpath_write_message).send_keys(Keys.ENTER)
    time.sleep(2)

    browser.find_element(By.XPATH, xpath_write_message).send_keys(message3)
    browser.find_element(By.XPATH, xpath_write_message).send_keys(Keys.ENTER)
    time.sleep(2)

    browser.find_element(By.XPATH, xpath_write_message).send_keys(message4)
    browser.find_element(By.XPATH, xpath_write_message).send_keys(Keys.ENTER)
    time.sleep(2)

    # find plus sign button to add attachment
    browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div').click()
    time.sleep(2)

    # find Documento button to add attachment
    # browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[1]/li/div/span').click()
    # time.sleep(3)

    # find Fotos e Vídeos button to add attachment
    browser.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[2]/li/div/span').click()
    time.sleep(3)

    # Insert file path and attach it
    pyautogui.write(os.path.abspath(f"data\\tesouraria\\{contact_file}.png"))
    pyautogui.press('enter')
    time.sleep(5)

    # Click to send the file
    browser.find_element(By.XPATH, '//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').click()
    time.sleep(4)

    browser.find_element(By.XPATH, xpath_write_message).send_keys(message5)
    browser.find_element(By.XPATH, xpath_write_message).send_keys(Keys.ENTER)
    time.sleep(2)

    # waiting to end it all
    counter = 10
    while counter > 0:
        print(f"Ending in {counter}")
        time.sleep(1)
        counter -= 1

print("SUCCESS!")
import time
from bs4 import BeautifulSoup
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
import json


with open(file="config.json", mode="r") as f:
    data = json.load(f)

headers = data["headers"][0]
website = "https://cs.utdallas.edu/people/faculty/"
google_form = data["google_form"]

# Scraping the results

response = requests.get(url=website,
                        headers=headers).text
soup = BeautifulSoup(response, 'html.parser')
contact_card = []
for item in soup.tbody.find_all("tr"):
    try: 
        prof_name = item.td.text.strip().split('\n')[0].split(',')[0]
        email = item.find_all('td')[1].text.strip().split('\n')[0].split(',')[0]
        contact = {
        "Professor's Last Name": f"{prof_name}",
        "Email": f"{email}"
        }
        contact_card.append(contact)
    except AttributeError: pass
    except IndexError: pass
    

# Filling the google form

driver = webdriver.Chrome()
for contact in contact_card:
    driver.get(google_form)

    time.sleep(2)

    address = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input')
    address.click()
    address.send_keys(contact["Professor's Last Name"])
    
    time.sleep(2)

    bedroom = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input')
    bedroom.click()
    bedroom.send_keys(contact["Email"])

    time.sleep(2)

    driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span').click()


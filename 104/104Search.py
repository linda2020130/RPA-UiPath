from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import numpy as np
import pandas as pd
import time
import requests


def initiateDriver():
    webdriver_path = 'C:\\IEDriverServer.exe'
    driver = webdriver.Ie(webdriver_path)
    driver.maximize_window()
    return driver

def login(driver, user, pwd):

    # Open IE browser and login to e-portal
    driver.get("http://***")
    id = driver.find_element_by_name("textfield32")
    password = driver.find_element_by_name("textfield33")
    id.send_keys(user)
    password.send_keys(pwd)
    password.send_keys(Keys.RETURN)

    # Dismiss alert if any popup alert happens
    try:
        WebDriverWait(driver, 2).until(EC.alert_is_present(), 'Timed out waiting for popup to appear')
        driver.switch_to.alert.dismiss()
        print("alert dismissed!")
    except TimeoutException:
        print("no dismissed")


def webScrape(driver, start, end):
    id = [ j + 1 for j in range(start, end)]
    columns = ['員工編碼', '姓名', 'First Name', 'Last Name', 
               '辦公室電話', 'EMAIL', '部門代碼', '部門名稱', '職稱']
    df = pd.DataFrame(np.zeros((end - start, 9)), columns=columns, index=id)
    for i in range(start, end):
        url = "http://***?ID=" + str(i + 1)
        req = requests.get(url)

        # Check if url works
        if req.status_code == requests.codes.ok: #pylint: disable=no-member
            try:
                driver.get(url)
                soup = BeautifulSoup(driver.page_source, "html.parser")

            # Invalid ID
            except:
                WebDriverWait(driver, 2).until(EC.alert_is_present(), 'Timed out waiting for popup to appear')
                driver.switch_to.alert.dismiss()
                dataRow = ["InvalidID"] * 9
                df.iloc[i - start, :] = dataRow
            else:
                data = soup.find_all("nobr")
                if len(data) == 18:
                    dataRow = [data[2 * j + 1].text for j in range(9)]
                    df.iloc[i - start, :] = dataRow

                # No data
                else:
                    dataRow = [None] * 9
                    df.iloc[i - start, :] = dataRow
        else:
            print("url invalid for id " + str(i + 1))

    df.to_excel("Employee List.xlsx")


if __name__ == '__main__':
    tStart = time.time()
    driver = initiateDriver()
    login(driver, "userid", "password")
    webScrape(driver, 0, 4000)
    driver.quit()
    tEnd = time.time()
    print("Total running time: " + str(tEnd - tStart))
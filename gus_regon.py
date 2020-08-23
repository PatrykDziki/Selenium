import traceback
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import pandas as pd
import time
from pandas import ExcelWriter
from selenium.webdriver.support import expected_conditions as EC

index = 0
errors = 0
batch_data = pd.read_excel('nip_list.xlsx', dtype={'NIP': object})
driver = webdriver.Chrome()
driver.get("https://wyszukiwarkaregon.stat.gov.pl/appBIR/index.aspx")
log = {'NIP': [],
        'Nazwa': [],
        'PodstawowaFormaPrawna': [],
        'SzczególnaFormaPrawna': [],
        'NazwaFormyWłasności': [],
        'OrganRejestrowy': [],
        'RodzajRejestruLubEwidencji': [],
        'NumerRejestruLubEwidencji': [],
        'ERROR': []}

def save_logs_if_FIZ():
        log['Nazwa'].append(driver.find_element_by_id('fiz_nazwa').text)
        log['PodstawowaFormaPrawna'].append(driver.find_element_by_id('fiz_nazwaPodstawowejFormyPrawnej').text)
        log['SzczególnaFormaPrawna'].append(driver.find_element_by_id('fiz_nazwaSzczegolnejFormyPrawnej').text)
        log['NazwaFormyWłasności'].append(driver.find_element_by_id('fiz_nazwaFormyWlasnosci').text)
        log['OrganRejestrowy'].append('nie dotyczy')
        log['RodzajRejestruLubEwidencji'].append('nie dotyczy')
        log['NumerRejestruLubEwidencji'].append('nie dotyczy')
        log['ERROR'].append('brak błędu')
        time.sleep(0.2)

def save_logs_if_PRAW():
        log['PodstawowaFormaPrawna'].append(driver.find_element_by_id('praw_nazwaPodstawowejFormyPrawnej').text)
        log['SzczególnaFormaPrawna'].append(driver.find_element_by_id('praw_nazwaSzczegolnejFormyPrawnej').text)
        log['NazwaFormyWłasności'].append(driver.find_element_by_id('praw_nazwaFormyWlasnosci').text)
        log['Nazwa'].append(driver.find_element_by_id('praw_nazwa').text)
        log['OrganRejestrowy'].append(driver.find_element_by_id('praw_nazwaOrganuRejestrowego').text)
        log['RodzajRejestruLubEwidencji'].append(driver.find_element_by_id('praw_nazwaRodzajuRejestru').text)
        log['NumerRejestruLubEwidencji'].append(driver.find_element_by_id('praw_numerWrejestrzelubEwidencji').text)
        log['ERROR'].append('brak błędu')
        time.sleep(0.2)

def save_logs_if_NOTFOUND():
        log['Nazwa'].append('Nie znaleziono podmiotów.')
        log['PodstawowaFormaPrawna'].append('Nie znaleziono podmiotów.')
        log['SzczególnaFormaPrawna'].append('Nie znaleziono podmiotów.')
        log['NazwaFormyWłasności'].append('Nie znaleziono podmiotów.')
        log['OrganRejestrowy'].append('nie dotyczy')
        log['RodzajRejestruLubEwidencji'].append('nie dotyczy')
        log['NumerRejestruLubEwidencji'].append('nie dotyczy')
        log['ERROR'].append('Nie znaleziono podmiotów.')
        time.sleep(0.2)

def save_logs_if_ERROR():
        log['Nazwa'].append('Error')
        log['PodstawowaFormaPrawna'].append('Error')
        log['SzczególnaFormaPrawna'].append('Error')
        log['NazwaFormyWłasności'].append('Error')
        log['OrganRejestrowy'].append('Error')
        log['RodzajRejestruLubEwidencji'].append('Error')
        log['NumerRejestruLubEwidencji'].append('Error')
        log['ERROR'].append('Error')
        time.sleep(0.2)


while index < len(batch_data['NIP']):

    df = pd.DataFrame(log)
    writer = ExcelWriter('gus_regon-raport-wyjsciowy.xlsx')
    df.to_excel(writer, 'Sheet1', index=False)

    if errors > 2:
        try:
            print('reloading the page...')
            driver = webdriver.Chrome()
            driver.get("https://wyszukiwarkaregon.stat.gov.pl/appBIR/index.aspx")
            errors = 0
        except:
            print(traceback.format_exc())
            time.sleep(60)
            continue

    try:
        time.sleep(1)
        log['NIP'].append(batch_data['NIP'][index])
        WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, 'txtNip')))
        nip_textbox = driver.find_element_by_id('txtNip')
        nip_textbox.clear()
        nip_textbox.send_keys(batch_data['NIP'][index])
        nip_textbox.send_keys(Keys.RETURN)
        time.sleep(1)

        if driver.find_element_by_id('divInfoKomunikat').text == 'Nie znaleziono podmiotów.':
            save_logs_if_NOTFOUND()
            writer.save()
            print(f'{index})', batch_data['NIP'][index], 'Nie znaleziono podmiotów.')
            index += 1

        else:
            links = driver.find_elements_by_partial_link_text('')
            for link in links:
                if 'javascript' in link.get_attribute("href"):
                    link.click()
                    break

            time.sleep(0.2)

            if driver.find_element_by_id('praw_nazwaPodstawowejFormyPrawnej').text == '':
                save_logs_if_FIZ()
                writer.save()
                print(f'{index})', batch_data['NIP'][index], driver.find_element_by_id('fiz_nazwa').text)
                errors = 0
                index += 1

            else:
                save_logs_if_PRAW()
                writer.save()
                print(f'{index})', batch_data['NIP'][index], driver.find_element_by_id('praw_nazwa').text)
                writer.save()
                errors = 0
                index += 1

    except:
        save_logs_if_ERROR()
        print('something goes wrong =(')
        writer.save()
        errors += 1
        index += 1

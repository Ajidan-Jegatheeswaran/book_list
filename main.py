import xlsxwriter
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time

driverService = Service('webdriver97/chromedriver.exe')
driver = webdriver.Chrome(service=driverService)

#create Excel
print('Excel wird geschrieben...')
workbook = xlsxwriter.Workbook('booklist.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Autor')
worksheet.write('B1', 'Name')
worksheet.write('C1', 'Seitenzahl')
worksheet.write('D1', 'Gattung')

print('Excel wurde geschlossen...')
check = True

def getSeitenzahl(suche, check=check):
    try:
        link = 'https://www.orellfuessli.ch/suche?filterPATHROOT=&sq=' + suche.replace(' ', '+')

        driver.get(link)
        element = driver.find_element(by=By.XPATH, value='/html/body/main/section/suchergebnis-liste/ul/li[1]/a')
        link_to_book = element.get_attribute('href')
        driver.get(link_to_book)
        elements = driver.find_elements(By.CLASS_NAME, 'dataLabel')

        counter = 1
        for i in elements:

            if 'Seitenzahl' in i.text:

                break

            counter += 1

        element = driver.find_element(By.XPATH, '//*[@id="pmProductView ads-grid"]/main/div/section[3]/div/div[2]/div[1]/table/tbody/tr[' + str(counter) + ']/td')

        print('Element: ' + str(element.text))
        return element.text
    except:
        return ''

file = open('data/booklist_roh.txt')
list_of_book = file.read().splitlines()
print(list_of_book)
for item in list_of_book:
    if item == '':
        list_of_book.remove(item)

print(list_of_book)



autor = ''
name = ''
gattung = ''
seitenzahl = ''

insCounter = 2
counter = 1
for item in list_of_book:
    if counter == 1:
        autor=item
        counter += 1
    elif counter == 2:
        name=item
        counter += 1
        seitenzahl = getSeitenzahl(name + '+' + autor + '+' + gattung)
        time.sleep(10)
        if seitenzahl == '':
            seitenzahl = getSeitenzahl(name + autor)
            time.sleep(10)
            if seitenzahl == '':
                seitenzahl = getSeitenzahl(name)
                time.sleep(10)
                if seitenzahl == '':
                    seitenzahl = getSeitenzahl(autor)
                    time.sleep(10)
                    if seitenzahl == '':
                        seitenzahl = getSeitenzahl(gattung)
                        time.sleep(10)

    elif counter == 3:
        gattung = item
        worksheet.write('A'+str(insCounter), autor)
        worksheet.write('B' + str(insCounter), name)
        worksheet.write('C' + str(insCounter), seitenzahl)
        worksheet.write('D' + str(insCounter), gattung)
        insCounter += 1
        counter = 1
        print('Autor: '+autor, 'Name: ' + name, 'Seitenzahl: ' + seitenzahl, 'Gattung: ' + gattung)

workbook.close()


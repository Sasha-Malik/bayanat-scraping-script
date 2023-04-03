from selenium import webdriver

from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait

from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.chrome.options import Options



import os

import time

import csv



pageCount = 1 #page to start from

nextPage = 16 #next button

pageHelper = 1



# Path to the folder where downloaded data sets will be saved

save_folder = '/Users/ssm9960/desktop/bayanat-data'

if not os.path.exists(save_folder):

    os.makedirs(save_folder)



url = "https://admin.bayanat.ae/home/datasets?langKey=en"

driver = webdriver.Chrome()

driver.get(url)



wait = WebDriverWait(driver, 10)



# Find the first page link and click it to load the page

wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".kt-pagination__link--active"))).click()



links = []

names = []



with open('/Users/ssm9960/desktop/bayanat-data/metadata', 'w', encoding='UTF8') as f:

    writer = csv.writer(f)

    header = ['name','description','tags']

    writer.writerow(header)







while True:

    # Scrape data from the current page

    div_element = driver.find_element(By.CSS_SELECTOR,"#datasetListing > div")

    div_children = div_element.find_elements(By.XPATH,"./*")



    for child in div_children:



        csv_data = []



        desired_child = child.find_element(By.XPATH,".//*[contains(@class, 'kt-widget4__username')][1]")

        print(desired_child.get_attribute("innerHTML").strip())

        names.append(desired_child.get_attribute("innerHTML"))

        curName = desired_child.get_attribute("innerHTML").strip()

        csv_data.append(curName)

        link = desired_child.get_attribute("href")

        links.append(link)

        links.append("/")



        try:

            description_div = child.find_element(By.XPATH,".//*[contains(@class, 'kt-widget4__text')][1]")

            #print(description_div.get_attribute("innerHTML").strip())

            description = description_div.get_attribute("innerHTML").strip()

        except:

            description = "NA"

        csv_data.append(description)





        try:

            tag_div = child.find_element(By.XPATH,".//*[contains(@class, 'badges-container mt-2')][1]")

            tags = tag_div.find_elements(By.XPATH,"./*")

            tagList = []



            for t in tags:

                #print(t.get_attribute("innerHTML"))

                tagList.append(t.get_attribute("innerHTML"))

                csv_data.append(t.get_attribute("innerHTML"))

        except:

             csv_data.append("NA")



        with open('/Users/ssm9960/desktop/bayanat-data/metadata', 'a', encoding='UTF8') as f:

            writer = csv.writer(f)

            writer.writerow(csv_data)



    print(pageCount)

    if pageCount == 261:

        break

    pageCount += 1





    try:

        #Find the next page link and click it to load the next page

        if pageCount == nextPage:

            pageHelper = 1

            selector_path = "#page-selection > ul > li:nth-child(" + str(pageHelper + 3) + ")"

            pageHelper += 1

            nextPage += 14

        else:

            selector_path = "#page-selection > ul > li:nth-child(" + str(pageHelper + 3) + ")"

            pageHelper += 1



        wait = WebDriverWait(driver, 20)

        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,  selector_path)))

        element = driver.find_element(By.CSS_SELECTOR, selector_path)

        element.click()

        time.sleep(5)



    except:

        # No more pages to load

        break



curPath = '/Users/ssm9960/desktop/bayanat-data/'

nameCount = 0

for l in links:

    if l == "/":

        nameCount+=1

        continue

    driver.get(l)

    div_element = driver.find_element(By.CSS_SELECTOR,"#ResourceListing")

    div_children = div_element.find_elements(By.XPATH,"./*")

    for child in div_children:

        try:

            wait = WebDriverWait(driver, 10)

            element = wait.until(EC.element_to_be_clickable((By.XPATH, ".//*[contains(@class, 'btn btn-label-primary btn-bold btn-sm btn-icon-h kt-margin-l-10 download-btn-rsrc')][1]")))



            desired_child = child.find_element(By.XPATH,".//*[contains(@class, 'btn btn-label-primary btn-bold btn-sm btn-icon-h kt-margin-l-10 download-btn-rsrc')][1]")

            download_link = desired_child.get_attribute("href")

            print(names[nameCount])

            print(download_link)

            PATH = curPath + names[nameCount]



            if not os.path.exists(PATH):

                os.makedirs(PATH)



            options = Options()

            prefs = {"download.default_directory" : PATH}

            options.add_experimental_option("prefs",prefs)

            driver = webdriver.Chrome(options=options)

            driver.get(download_link)

            time.sleep(2)



        except:

            print("Element not found within timeout period")

driver.quit()


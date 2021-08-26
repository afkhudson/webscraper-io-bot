import csv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui
import time
import os
import pygetwindow
import glob
import pandas as pd
import os.path, time
import openpyxl
from datetime import datetime
import xlwings as xw

path_chromedriver = 'C:/Users/sosa/OneDrive/Coding/chromedriver.exe'

options = Options()
options.headless = False
options.add_argument("--user-data-dir=C:/Users/sosa/AppData/Local/Google/Chrome/User Data")
options.add_argument('--profile-directory=Profile 1')
options.add_argument("--auto-open-devtools-for-tabs")
options.add_argument("--enable-javascript")

def runAdorama1440():
    os.system("taskkill /im chrome.exe /f")
    browser = webdriver.Chrome(executable_path=path_chromedriver, options=options)
    browser.maximize_window()
    browser.get('https://www.google.com/?gws_rd=ssgl')
    browser.set_page_load_timeout(10)

    #--------- RUN ADORAMA SCRAPER
    print(pyautogui.position())
    #pyautogui.click(1893,15)
    pyautogui.click(768, 1006) # click webscraper
    time.sleep(1)
    pyautogui.click(82, 1154) # click sitemap adorama
    time.sleep(1)
    pyautogui.click(185, 1031) # click sitemap functions
    time.sleep(1)
    pyautogui.click(140, 1139) #click scrape
    time.sleep(1)
    pyautogui.click(252, 1123) #click run
    time.sleep(4)

    browser.maximize_window() #back to browser
    pyautogui.click(185, 1031) # click sitemap functions
    time.sleep(1)
    pyautogui.click(199, 1214) #click export as csv
    time.sleep(1)
    pyautogui.click(422, 1095) # export csv / click download now
    time.sleep(2)

def runAdorama1080():
    os.system("taskkill /im chrome.exe /f")
    browser = webdriver.Chrome(executable_path=path_chromedriver, options=options)
    browser.maximize_window()
    browser.get('https://www.google.com/?gws_rd=ssgl')
    browser.set_page_load_timeout(10)

    #--------- RUN ADORAMA SCRAPER
    print(pyautogui.position())
    #pyautogui.click(1893,15)
    pyautogui.click(784, 645) # click webscraper
    time.sleep(1)
    pyautogui.click(167, 759) # click sitemap adorama
    time.sleep(1)
    pyautogui.click(194, 674) # click sitemap functions
    time.sleep(1)
    pyautogui.click(175, 780) #click webscraper
    time.sleep(1)
    pyautogui.click(190, 757) #click run
    time.sleep(4)

    browser.maximize_window() #back to browser
    pyautogui.click(194, 674) # click sitemap functions
    time.sleep(1)
    pyautogui.click(194, 852) #click export as csv
    time.sleep(1)
    pyautogui.click(418, 737) # export csv
    time.sleep(2)

def runbby1080():
    os.system("taskkill /im chrome.exe /f")
    browser = webdriver.Chrome(executable_path=path_chromedriver, options=options)
    browser.maximize_window()
    browser.get('https://www.google.com/?gws_rd=ssl')
    browser.set_page_load_timeout(10)

    #--------- RUN ADORAMA SCRAPER
    print(pyautogui.position())
    #pyautogui.click(1893,15)
    pyautogui.click(784, 645) # click webscraper
    time.sleep(1)
    pyautogui.click(80, 880) # click sitemap bby
    time.sleep(1)
    pyautogui.click(194, 674) # click sitemap functions
    time.sleep(1)
    pyautogui.click(175, 780) #click webscraper
    time.sleep(1)
    pyautogui.click(190, 757) #click run
    time.sleep(50)

    browser.maximize_window() #back to browser
    pyautogui.click(194, 674) # click sitemap functions
    time.sleep(1)
    pyautogui.click(194, 852) #click export as csv
    time.sleep(1)
    pyautogui.click(418, 737) # export csv
    time.sleep(2)

def runbby1440():
    os.system("taskkill /im chrome.exe /f")
    browser = webdriver.Chrome(executable_path=path_chromedriver, options=options)
    browser.maximize_window()
    browser.get('https://www.google.com/?gws_rd=ssgl')
    browser.set_page_load_timeout(10)

    #--------- RUN ADORAMA SCRAPER
    print(pyautogui.position())
    #pyautogui.click(1893,15)
    pyautogui.click(768, 1006) # click webscraper
    time.sleep(1)
    pyautogui.click(81, 1255) # click sitemap bby
    time.sleep(1)
    pyautogui.click(185, 1031) # click sitemap functions
    time.sleep(1)
    pyautogui.click(140, 1139) #click scrape
    time.sleep(1)
    pyautogui.click(252, 1123) #click run
    time.sleep(60)

    browser.maximize_window() #back to browser
    pyautogui.click(185, 1031) # click sitemap functions
    time.sleep(1)
    pyautogui.click(199, 1214) #click export as csv
    time.sleep(1)
    pyautogui.click(422, 1095) # export csv / click download now
    time.sleep(2)

def runBH1440():
    os.system("taskkill /im chrome.exe /f")
    browser = webdriver.Chrome(executable_path=path_chromedriver, options=options)
    browser.maximize_window()
    browser.get('https://www.google.com/?gws_rd=ssgl')
    browser.set_page_load_timeout(10)

    #--------- RUN ADORAMA SCRAPER
    print(pyautogui.position())
    #pyautogui.click(1893,15)
    pyautogui.click(768, 1006) # click webscraper
    time.sleep(1)
    pyautogui.click(127, 1282) # click sitemap bhphoto
    time.sleep(1)
    pyautogui.click(185, 1031) # click sitemap functions
    time.sleep(1)
    pyautogui.click(140, 1139) #click scrape
    time.sleep(1)
    pyautogui.click(252, 1123) #click run
    time.sleep(8)

    browser.maximize_window() #back to browser
    pyautogui.click(185, 1031) # click sitemap functions
    time.sleep(1)
    pyautogui.click(199, 1214) #click export as csv
    time.sleep(1)
    pyautogui.click(422, 1095) # export csv / click download now
    time.sleep(2)

def runMSFT1440():
    os.system("taskkill /im chrome.exe /f")
    browser = webdriver.Chrome(executable_path=path_chromedriver, options=options)
    browser.maximize_window()
    browser.get('https://www.google.com/?gws_rd=ssgl')
    browser.set_page_load_timeout(10)

    #--------- RUN ADORAMA SCRAPER
    print(pyautogui.position())
    #pyautogui.click(1893,15)
    pyautogui.click(768, 1006) # click webscraper
    time.sleep(1)
    pyautogui.click(97, 1303) # click sitemap msft
    time.sleep(1)
    pyautogui.click(185, 1031) # click sitemap functions
    time.sleep(1)
    pyautogui.click(140, 1139) #click scrape
    time.sleep(1)
    pyautogui.click(252, 1123) #click run
    time.sleep(5)

    browser.maximize_window() #back to browser
    pyautogui.click(185, 1031) # click sitemap functions
    time.sleep(1)
    pyautogui.click(199, 1214) #click export as csv
    time.sleep(1)
    pyautogui.click(422, 1095) # export csv / click download now
    time.sleep(3)

def update_channel_file():
    #open channel file
    channel_file = r"C:/Users/sosa/OneDrive - Razer (Asia-Pacific) Pte. Ltd/Razer/Worksheets/Channel Health Dashboard/Channel Health Dashboard.xlsm"
    wb_app = xw.App(visible=False) #show.hide sheet
    wb_app.display_alerts=False
    wb = xw.Book(channel_file)

    msftdata = wb.sheets['Microsoft Data']
    bhdata = wb.sheets['BH Data']
    adoramadata = wb.sheets['Adorama Data']
    bbydata = wb.sheets['Best Buy Data']

    def updateMSFT():
        msft_dir = "C:/Users/sosa/Downloads/microsoft-systems*"
        msft_latest = max(glob.glob(msft_dir), key=os.path.getctime) #gets latest file from dir
        print(f'msft file created: {time.ctime(os.path.getmtime(msft_latest))}')

        #turning created date into pandas date
        created_time = time.ctime(os.path.getctime(msft_latest)) #Sun Jul 25 19:07:30 2021
        cut_time = created_time[4:7] + " " + created_time[8:10].strip() + " " + created_time[-4:] #getting elements to transform
        datetime_object1 = datetime.strptime(cut_time, '%b %d %Y')
        shortdate = datetime_object1

        df = pd.read_csv(open(msft_latest))

        df[df.columns[5:7]] = df[df.columns[5:7]].replace('[/$,]', '', regex=True).astype(float) #removing dollar sign from pricing
        df.insert(0, "formula", "", True) #adding empty formula column
        df.insert(0, "Time", shortdate, True) #adding clean date

        print(f'lines of data for msft: {len(df)}')
        lastrow = msftdata.range('A' + str(msftdata.cells.last_cell.row)).end('up').address #finding last row
        #finding last cell and adding +1
        emptycell = lastrow.replace("$", "")
        num = str(int(emptycell[1:])+1)
        fixed = emptycell[0] + num
        print(f'pasted to line: {fixed}')

        msftdata.range(fixed).options(index=False,header=False).value = df
        msftdata.range('A:M').api.WrapText = False #remove wrap text
        print('-----------DONE')

    def updateBH():
        bh_dir = "C:/Users/sosa/Downloads/bhphoto-systems*"
        bh_latest = max(glob.glob(bh_dir), key=os.path.getctime) #gets latest file from dir
        print(f'bh file created: {time.ctime(os.path.getmtime(bh_latest))}')

        #turning created date into pandas date
        created_time = time.ctime(os.path.getctime(bh_latest)) #Sun Jul 25 19:07:30 2021
        cut_time = created_time[4:7] + " " + created_time[8:10].strip() + " " + created_time[-4:] #getting elements to transform
        datetime_object1 = datetime.strptime(cut_time, '%b %d %Y')
        shortdate = datetime_object1

        df = pd.read_csv(open(bh_latest))

        df[df.columns[5:6]] = df[df.columns[5:6]].replace('[/$,]', '', regex=True).astype(float) #removing dollar sign from pricing
        df.insert(0, "formula1", "", True) #adding empty formula column
        df.insert(0, "formula2", "", True) #adding empty formula column
        df.insert(0, "Time", shortdate, True) #adding clean date

        print(f'lines of data for bh: {len(df)}')
        lastrow = bhdata.range('A' + str(bhdata.cells.last_cell.row)).end('up').address #finding last row
        #finding last cell and adding +1
        emptycell = lastrow.replace("$", "")
        num = str(int(emptycell[1:])+1)
        fixed = emptycell[0] + num
        print(f'pasted to line: {fixed}')

        bhdata.range(fixed).options(index=False,header=False).value = df
        bhdata.range('A:M').api.WrapText = False #remove wrap text
        print('-----------DONE')


    def updateadorama():
        adorama_dir = "C:/Users/sosa/Downloads/adorama-systems*"
        adorama_latest = max(glob.glob(adorama_dir), key=os.path.getctime) #gets latest file from dir
        print(f'adorama file created: {time.ctime(os.path.getmtime(adorama_latest))}')

        #turning created date into pandas date
        created_time = time.ctime(os.path.getctime(adorama_latest)) #Sun Jul 25 19:07:30 2021
        cut_time = created_time[4:7] + " " + created_time[8:10].strip() + " " + created_time[-4:] #getting elements to transform
        datetime_object1 = datetime.strptime(cut_time, '%b %d %Y')
        shortdate = datetime_object1

        df = pd.read_csv(open(adorama_latest))

        df[df.columns[5:6]] = df[df.columns[5:6]].replace('[/$,]', '', regex=True).astype(float) #removing dollar sign from pricing
        df.insert(0, "formula1", "", True) #adding empty formula column
        df.insert(0, "Time", shortdate, True) #adding clean date

        print(f'lines of data for adorama: {len(df)}')
        lastrow = adoramadata.range('A' + str(adoramadata.cells.last_cell.row)).end('up').address #finding last row
        #finding last cell and adding +1
        emptycell = lastrow.replace("$", "")
        num = str(int(emptycell[1:])+1)
        fixed = emptycell[0] + num
        print(f'pasted to line: {fixed}')

        adoramadata.range(fixed).options(index=False,header=False).value = df
        adoramadata.range('A:M').api.WrapText = False #remove wrap text
        print('-----------DONE')


    def updatebby():
        bby_dir = "C:/Users/sosa/Downloads/bby-scraper*"
        bby_latest = max(glob.glob(bby_dir), key=os.path.getctime) #gets latest file from dir
        print(f'bby file created: {time.ctime(os.path.getmtime(bby_latest))}')

        #turning created date into pandas date
        created_time = time.ctime(os.path.getctime(bby_latest)) #Sun Jul 25 19:07:30 2021
        cut_time = created_time[4:7] + " " + created_time[8:10].strip() + " " + created_time[-4:] #getting elements to transform
        datetime_object1 = datetime.strptime(cut_time, '%b %d %Y')
        shortdate = datetime_object1

        df = pd.read_csv(open(bby_latest))

        df[df.columns[3:4]] = df[df.columns[3:4]].replace('[/$,]', '', regex=True).astype(float) #removing dollar sign from pricing
        df.insert(0, "formula1", "", True) #adding empty formula column
        df.insert(0, "formula2", "", True) #adding empty formula column
        df.insert(0, "Time", shortdate, True) #adding clean date

        print(f'lines of data for bby: {len(df)}')
        lastrow = bbydata.range('A' + str(bbydata.cells.last_cell.row)).end('up').address #finding last row
        #finding last cell and adding +1
        emptycell = lastrow.replace("$", "")
        num = str(int(emptycell[1:])+1)
        fixed = emptycell[0] + num
        print(f'pasted to line: {fixed}')

        bbydata.range(fixed).options(index=False,header=False).value = df
        bbydata.range('A:M').api.WrapText = False #remove wrap text
        print('-----------DONE')


    updateMSFT()
    updateBH()
    updateadorama()
    updatebby()

    wb.save()
    wb.close()
    wb_app.quit()


def runall1440():
    input('Press any key to start scraper')
    runMSFT1440()
    runBH1440()
    runbby1440()
    runAdorama1440()



#run after this line ---------------
#runall1440()
update_channel_file()

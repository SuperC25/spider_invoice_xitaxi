from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import xlrd
import os
from PIL import Image

def cutscreen(pp):
    im = Image.open(pp)
    #chrome浏览器
    #box=(316,471,1584,816)
    #chrome --headless
    box=(452,376,1468,655)
    #PhantomJS浏览器
    #box=(173,380,1200,672)
    region = im.crop(box)
    region.save(pp)

#dir = path.dirname(__file__)
dir = os.getcwd()
#打开发票excel
ExcelFile = xlrd.open_workbook(dir+"\\"+"source.xlsx")

#sheet表0
sheet=ExcelFile.sheet_by_index(0)
for x in range(sheet.nrows):
    rows = sheet.row_values(x)
    file_num=x+1
    fpdm=rows[0]
    fphm=rows[1]
    #browser = webdriver.Chrome()
    #browser = webdriver.PhantomJS()
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    #chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--start-maximized')
    #chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options)
    url = "http://12366.xj-n-tax.gov.cn/BsfwtWeb/pages/cx/sscx/fpcx.html"
    browser.get(url)
    browser.set_window_size(1920, 1080)
    #brower.maximize_window()
    time.sleep(1)
    browser.find_element_by_xpath("/html/body/div[@class='center-1200']/div[@class='cx-box-wrap']/div[@class='width-1000']/table[@class='width-1000']/tbody/tr[1]/td/label[3]/input").click()
    browser.find_element_by_xpath("/html/body/div[@class='center-1200']/div[@class='cx-box-wrap']/div[@class='width-1000']/table[@class='width-1000']/tbody/tr[2]/td[1]/span[@id='tjfpdm']/span[@class='mini-textbox-border']/input[@id='tjfpdm$text']").send_keys(fpdm)
    browser.find_element_by_xpath("/html/body/div[@class='center-1200']/div[@class='cx-box-wrap']/div[@class='width-1000']/table[@class='width-1000']/tbody/tr[2]/td[2]/span[@id='tjfphm']/span[@class='mini-textbox-border']/input[@id='tjfphm$text']").send_keys(fphm)
    browser.find_element_by_xpath("/html/body/div[@class='center-1200']/div[@class='cx-box-wrap']/div[@class='width-1000']/table[@class='width-1000']/tbody/tr[3]/td/input[@class='btn-base'][2]").click()
    #time.sleep(2)
    try:
        element = WebDriverWait(browser,20).until(
            EC.visibility_of_element_located((By.ID, "cx_rq8"))
        )
    except:
        print('error in '+str(file_num))
    finally:
        screen=dir+"\\pictures\\"+str(file_num)+".png"
        browser.save_screenshot(screen)
        cutscreen(screen)
        #browser.close()
        browser.quit()
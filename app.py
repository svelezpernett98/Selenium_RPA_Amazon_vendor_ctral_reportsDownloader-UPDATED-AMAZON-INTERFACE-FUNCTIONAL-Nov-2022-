import glob
from pydoc import importfile
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time, os
import pandas as pd
from datetime import datetime
import win32com.client as win32

currentMonth = datetime.now().month
dir = os.getcwd()
PATH = str(dir)+"/chromedriver.exe"
picture_options = Options()
picture_options.add_argument("--start-maximized")
picture_options.add_experimental_option(
    "prefs", {"download.default_directory":"C:\\Users\\fmoncayo\\Documents\\FrancoMoncayoUribe\\RPA\\BotAmazon\\excel"}
)

init = webdriver.Chrome(PATH, chrome_options=picture_options)  
init.get('https://vendorcentral.amazon.com')
user = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.ID, 'ap_email')))
user.send_keys('username')
pwd = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.ID, 'ap_password')))
pwd.send_keys("password")

enter_btn = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.ID, 'signInSubmit'))).click()

reports_analytics = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH,'/html/body/div/div[1]/div/div/div[2]/div[5]/span'))).click()

analytics_option = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div[1]/div/div/div[2]/div[5]/div/a[2]'))).click()

sales_diagnostic = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div/div/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div/ul/li[1]/div/div/div/span[1]/a'))).click()

time_frame_ddl = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div/div[3]/div[1]/div[2]/div[1]/div[1]/div/div[1]/kat-dropdown'))).click()

monthly_option = init.execute_script("""return document.querySelector("#time-period").shadowRoot.querySelector("div.kat-select-container > div.select-options > div > div > slot > kat-option:nth-child(3)")""").click()

year_ddl = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div/div[3]/div[1]/div[2]/div[1]/div[1]/div/div[2]/kat-dropdown'))).click()

year_2021 = init.execute_script("""return document.querySelector("#monthly-year").shadowRoot.querySelector("div.kat-select-container > div.select-options > div > div > slot > kat-option:nth-child(2)")""").click()

distributor_view_ddl = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div/div[3]/div[1]/div[2]/div[1]/div[1]/div/div[4]/kat-dropdown'))).click()

sourcing_option = init.execute_script("""return document.querySelector("#distributorView").shadowRoot.querySelector("div.kat-select-container > div.select-options > div > div > slot > kat-option:nth-child(2)")""").click()

def file_download_monthly(x,y):

    for i in range(x,0):

        string_i = str(i)
        string_i_final = string_i.replace("-", "")
        
        month_ddl = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div/div[3]/div[1]/div[2]/div[1]/div[1]/div/div[3]/kat-dropdown'))).click()

        month_option = init.execute_script("""return document.querySelectorAll(".ltr-1and29")[2].shadowRoot.querySelector("div.kat-select-container > div.select-options > div > div > slot > kat-option:nth-child("""+str(string_i_final)+""")")""").click()

        apply_btn = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div/div[3]/div[1]/div[2]/div[1]/div[1]/div/div[6]/kat-button'))).click()
        time.sleep(5)

        download_btn = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div/div[3]/div[1]/div[2]/div[1]/div[2]/div/div/div[1]/kat-button'))).click()
        time.sleep(2)
        view_downloads_link = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div/div[3]/div[1]/div[2]/div[1]/div[2]/div/div/div[2]/a'))).click()
        time.sleep(30)

        download_link = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div[1]/div[2]/kat-table/kat-table-body/kat-table-row[1]/kat-table-cell[2]/a'))).click()
        file_name = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div[1]/div[2]/kat-table/kat-table-body/kat-table-row[1]/kat-table-cell[1]/div[1]'))).text
        old_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/"+str(file_name)+".csv"
        time.sleep(10)

        if "1-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_Ene_"+y+".csv"

        if "2-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_Feb_"+y+".csv"

        if "3-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_Mar_"+y+".csv"

        if "4-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_Abr_"+y+".csv"

        if "5-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_May_"+y+".csv"

        if "6-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_Jun_"+y+".csv"

        if "7-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_Jul_"+y+".csv"

        if "8-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_Ago_"+y+".csv"

        if "9-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_Sep_"+y+".csv"

        if "10-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_Oct_"+y+".csv"
        
        if "11-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_Nov_"+y+".csv"

        if "12-1-"+str(y) in file_name:
            new_name = "C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/Sales_Sourcing_Dic_"+y+".csv"

        os.rename(old_name, new_name)
        
        close_side_tab = init.execute_script("""return document.querySelector("#root > div > div.ltr-1odx2my > div.ltr-15zcp20 > kat-icon.ltr-1upyiy5")""").click()

    fileList = os.listdir('C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel')
    excelWriter = pd.ExcelWriter('C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/Reporte/Reportes_Amazon_'+y+"_"+str(currentMonth)+'.xlsx',engine='xlsxwriter')
    files = [file.split('.',1)[0] for file in fileList]

    for file in files:

        excel_list = pd.read_csv("C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/"+file+".csv")
        excel_list.to_excel(excelWriter,sheet_name=file,index=True)

    excelWriter.save()

file_download_monthly(-12,"2021")

year_ddl = WebDriverWait(init, 300).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/div/div/div[3]/div[1]/div[2]/div[1]/div[1]/div/div[2]/kat-dropdown'))).click()
year_2022 = init.execute_script("""return document.querySelector("#monthly-year").shadowRoot.querySelector("div.kat-select-container > div.select-options > div > div > slot > kat-option:nth-child(1)")""").click()

mes_2022 = 0
if currentMonth == 10:
    mes_2022 = -9
elif currentMonth == 11:
    mes_2022 = -10
elif currentMonth == 12:
    mes_2022 = -11
elif currentMonth == 1:
    mes_2022 = -12

file_download_monthly(mes_2022,"2022")

fileList = glob.glob('C:/Users/fmoncayo/Documents/FrancoMoncayoUribe/RPA/BotAmazon/excel/*')

for file in fileList:
    os.remove(file)

outlook= win32.Dispatch('outlook.application')

mail= outlook.CreateItem(0)

mail.To= 'email@email.com'

mail.Subject= 'Archivos Sales_Sourcing descargados satisfactoriamente'

mail.Body= 'Mensaje generado por TI'

mail.HTMLBody= ('Archivos Sales_Sourcing descargados satisfactoriamente')

mail.Send()

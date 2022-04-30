#builtin
from numpy import nan
import pandas
import time
import tkinter
from tkinter.ttk import *
from concurrent.futures import ThreadPoolExecutor
from urllib.request import urlopen

#downloaded
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

#root settings
root = tkinter.Tk()
root.title('RERA Local Assistant')
root.geometry("300x200")

#variables
URL = 'https://maharera.mahaonline.gov.in/'
FILENAME = 'format.xlsx'
TFSI_DEFAULT = tkinter.StringVar(root, value = 0)

#drop box list
district_dict = {
    'Select District' : 0,
    'Mumbai City' : 17,
    'Mumbai Suburban' : 18,
    'Palghar' : 24,
    'Pune' : 26,
    'Thane' : 33
}


#functions
def fetch(link):
    html = urlopen(link).read()
    soup = BeautifulSoup(html, "lxml")

    for script in soup(["script", "style"]):
        script.decompose()

    strips = list(soup.stripped_strings)
    if strips[4] == 'Other Than Individual':
        strips = strips[4:300]

        add_details = strips.index('Address Details')
        org_name = strips[3]
        block_no = strips[strips.index('Block Number', add_details) + 1]
        building_name = strips[strips.index('Building Name', add_details) + 1]
        street_name = strips[strips.index('Street Name', add_details) + 1]
        locality = strips[strips.index('Locality', add_details) + 1]
        district = strips[strips.index('District', add_details) + 1]

        member_names = ''
        try:
            name_ = strips.index('Name :', add_details)
        except ValueError:
            return None
        if 'Member Information' in strips:
            try:
                member_info = strips.index('Member Information', add_details)
                small_strips = strips[member_info:name_]
                member_names = ''

                if 'Photo' in small_strips:
                    for m in small_strips[4::3]:
                        member_names += m + ',\n'
                
            except ValueError:
                member_names = ''
        elif 'Other Organization Type Member Information' in strips:
            try:
                ootmi = strips.index('Other Organization Type Member Information', add_details)
                small_strips = strips[ootmi:name_]
                member_names = ''

                for m in small_strips[3::2]:
                    member_names += m + ',\n'

            except ValueError:
                member_names = ''
        project = strips.index('Project', name_)
        fsi_details = strips.index('FSI Details', project)
        project_name = strips[strips.index('Project Name', project) + 1]
        if 'Extended Date of Completion' in strips[project:fsi_details]:
            pdoc = str(strips[strips.index('Extended Date of Completion', project) + 1])
        elif 'Revised Proposed Date of Completion' in strips[project:fsi_details]:
            pdoc = str(strips[strips.index('Revised Proposed Date of Completion', project) + 1])
        elif 'Proposed Date of Completion' in strips[project:fsi_details]:
            pdoc = str(strips[strips.index('Proposed Date of Completion', project) + 1])
        project_district = strips[strips.index('District', project) + 1]

        tfsi = float(strips[strips.index('TotalFSI', fsi_details) + 1])

        return [org_name, block_no, building_name, street_name, locality, district, member_names, project_name, pdoc, project_district, tfsi]
    else:
        return None

def on_run():
    run_button.config(state='disabled')
    progress.config(value=10)
    root.update_idletasks()

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get(URL)
    district_value = district_dropbox.get()
    index_number = district_dict[district_value]
    organisation = organisation_entry.get()
    tfsi = float(tfsi_entry.get())

    time.sleep(2)
    
    driver.find_element(By.XPATH, '/html/body/header/div[5]/div/div/div/ul/li[5]/a').send_keys(Keys.ARROW_DOWN * 2, Keys.RETURN)
    driver.switch_to.alert.accept()
    
    time.sleep(2)
    
    driver.switch_to.window(driver.window_handles[1])
    driver.find_element(By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[1]/div/div[2]/input[1]').click()
    driver.find_element(By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[8]/div/div/div[1]/input').click()
    
    time.sleep(1)
    
    driver.find_element(By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[4]/div/div[2]/select').send_keys(Keys.RETURN * 3, Keys.ARROW_DOWN * 3, Keys.RETURN * 3)
    
    progress.config(value=15)
    root.update_idletasks()
    time.sleep(1)

    driver.find_element(By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[5]/div[1]/div[2]/select').send_keys(Keys.RETURN * 3, Keys.ARROW_DOWN * index_number, Keys.RETURN * 3)
    
    progress.config(value=20)
    root.update_idletasks()
    time.sleep(1)

    driver.find_element(By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[2]/div/div[2]/div[2]/input[1]').send_keys(organisation)
    
    progress.config(value=25)
    root.update_idletasks()
    time.sleep(1)

    driver.find_element(By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[8]/div/div/div[3]/input[1]').click()

    l = []
    i = 1
    rera_data = pandas.DataFrame(columns= ['Organisation Name', 'Block Number', 'Building Name', 'Street Name', 'Locality', 'District', 'Member Names', 'Project Name', 'Proposed DOC', 'District', 'Total FSI'])

    progress.config(value=30)
    root.update_idletasks()

    start_time = time.time()
    while True:
        if i == 50:
            l.append(driver.find_element(By.XPATH, f'/html/body/section/div/div/form/div[3]/div/div[2]/div[1]/div/table/tbody/tr[{i}]/td[5]/b/a').get_attribute('href'))
                
            with ThreadPoolExecutor(100) as p:
                results = p.map(fetch, l)
            for r in list(results):
                rera_data.loc[len(rera_data.index)] = r
            driver.find_element(By.XPATH, '/html/body/section/div/div/form/div[3]/div/div[2]/div[2]/ul/li[3]/button[3]').click()
            i = 1
            l = []

        try:
            l.append(driver.find_element(By.XPATH, f'/html/body/section/div/div/form/div[3]/div/div[2]/div[1]/div/table/tbody/tr[{i}]/td[5]/b/a').get_attribute('href'))
            print(i)
            i += 1
        except NoSuchElementException:
            with ThreadPoolExecutor(100) as p:
                results = p.map(fetch, l)
            for r in list(results):
                rera_data.loc[len(rera_data.index)] = r
            break
    driver.quit()

    progress.config(value=80)
    root.update_idletasks()

    rera_data.replace(r"^\s*$", nan, regex = True)
    rera_data.dropna(axis = 'index', how = 'all', thresh = None, subset = None, inplace = True)
    rera_data = rera_data[rera_data['Total FSI'] > tfsi]
    rera_data.reset_index(inplace=True, drop = True)
    rera_data.index += 1

    progress.config(value=85)
    root.update_idletasks()

    wb = openpyxl.load_workbook(FILENAME)
    ws = wb.active
    ws.cell(row = 2, column = 2, value = district_value)
    rows = dataframe_to_rows(rera_data)
    if tfsi == 0.0: tfsi = ''
    datafile = f'RERA DATA {district_value.upper().strip()} {organisation.upper().strip()} {str(tfsi).strip()}'

    progress.config(value=90)
    root.update_idletasks()

    for r_idx, row in enumerate(rows, 5):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row = r_idx, column = c_idx, value = value)
    wb.save(f'Excel Files\\{datafile.strip()}.xlsx')
    print(time.time() - start_time)

    progress.config(value=0)
    run_button.config(state = 'enabled')
    root.update_idletasks()

#labels
district_label = Label(root, text = 'Choose District in Maharashtra(if any): ')
organisation_label = Label(root, text = 'Enter Organisation Name (if any): ')
tfsi_label = Label(root, text = 'Enter Total FSI Greater Than(if any): ')

#widgets
keys = list(district_dict.keys())
district_dropbox = Combobox(root, values = keys, state = 'readonly')
district_dropbox.current(0)
run_button = Button(root, text = 'Run', command = on_run, state = "enabled")
organisation_entry = Entry(root)
tfsi_entry = Entry(root, textvariable = TFSI_DEFAULT)
progress = Progressbar(root, orient='horizontal', length=100, mode= 'determinate')

#placement
district_label.pack()
district_dropbox.pack()
organisation_label.pack()
organisation_entry.pack()
tfsi_label.pack()
tfsi_entry.pack()
progress.pack()
run_button.pack()

#application loop
root.mainloop()
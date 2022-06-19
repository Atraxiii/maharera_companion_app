# builtin
import pandas
import time
import tkinter
from tkinter.ttk import *
from concurrent.futures import ThreadPoolExecutor
import urllib.request
import urllib.error

# downloaded
import backoff
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl

# root settings
root = tkinter.Tk()
root.title('RERA Local Assistant')
root.geometry("300x200")

# variables
URL = 'https://maharera.mahaonline.gov.in/'
FILENAME = 'format.xlsx'
TFSI_DEFAULT = tkinter.StringVar(root, value=0)

# drop box list
district_dict = {
	'Select District': 0,
	'Mumbai City': 17,
	'Mumbai Suburban': 18,
	'Palghar': 24,
	'Pune': 26,
	'Thane': 33
}


# functions
@backoff.on_exception(backoff.expo,urllib.error.URLError,max_value=20)
def fetch(link):
	html = urllib.request.urlopen(link).read()
	soup = BeautifulSoup(html, "lxml")
	strips = []

	results = soup.find_all('div', {'class': 'col-md-3 col-sm-3'})
	for tag in results:
		strips.append(tag.text.strip())

	if strips[1] != 'Other Than Individual':
		return []

	df = pandas.read_html(link)
	member = ''
	for table in df:
		if 'Member Name' in table.columns:
			members = table['Member Name'].to_list()
			for i in members:
				if i == members[-1]:
					member += i
				member += (i + ',\n')
			break
		elif 'Name' in table.columns:
			members = table['Name'].to_string()
			for i in members:
				if i == members[-1]:
					member += i
				member += (i + ',\n')
			break

	data = [
		strips[4],
		strips[strips.index('Block Number') + 1],
		strips[strips.index('Building Name') + 1],
		strips[strips.index('Street Name') + 1],
		strips[strips.index('Locality') + 1],
		strips[strips.index('District') + 1],
		member,
		strips[strips.index('Project Name') + 1],
		strips[strips.index('Litigations related to the project ?') - 1],
		strips[strips.index('District', 29) + 1],
		strips[strips.index('TotalFSI') + 1],
	]
	print('SUCCESS', end='|', flush=True)
	return data


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

	driver.find_element(
		By.XPATH, '/html/body/header/div[5]/div/div/div/ul/li[5]/a').send_keys(Keys.ARROW_DOWN * 2, Keys.RETURN)
	driver.switch_to.alert.accept()

	time.sleep(2)

	driver.switch_to.window(driver.window_handles[1])
	driver.find_element(
		By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[1]/div/div[2]/input[1]').click()
	driver.find_element(
		By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[8]/div/div/div[1]/input').click()

	time.sleep(1)

	driver.find_element(By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[4]/div/div[2]/select').send_keys(
		Keys.RETURN * 3, Keys.ARROW_DOWN * 3, Keys.RETURN * 3)

	progress.config(value=15)
	root.update_idletasks()
	time.sleep(1)

	driver.find_element(By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[5]/div[1]/div[2]/select').send_keys(
		Keys.RETURN * 3, Keys.ARROW_DOWN * index_number, Keys.RETURN * 3)

	progress.config(value=20)
	root.update_idletasks()
	time.sleep(1)

	driver.find_element(
		By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[2]/div/div[2]/div[2]/input[1]').send_keys(organisation)

	progress.config(value=25)
	root.update_idletasks()
	time.sleep(1)

	driver.find_element(
		By.XPATH, '/html/body/section/div/div/form/div[2]/div[2]/div[8]/div/div/div[3]/input[1]').click()

	l = []
	i = 1
	rera_data = pandas.DataFrame(columns=['Organisation Name', 'Block Number', 'Building Name', 'Street Name', 'Locality', 'District', 'Members', 'Project Name', 'Proposed DOC', 'District', 'Total FSI'])

	progress.config(value=30)
	root.update_idletasks()

	start_time = time.time()
	while True:
		if i == 50:
			l.append(driver.find_element(
				By.XPATH, f'/html/body/section/div/div/form/div[3]/div/div[2]/div[1]/div/table/tbody/tr[{i}]/td[5]/b/a').get_attribute('href'))

			with ThreadPoolExecutor(75) as p:
				results = p.map(fetch, l)
			for r in list(results):
				if r == []:
					continue
				rera_data.loc[len(rera_data)] = r
			driver.find_element(
				By.XPATH, '/html/body/section/div/div/form/div[3]/div/div[2]/div[2]/ul/li[3]/button[3]').click()
			i = 1
			l = []

		try:
			l.append(driver.find_element(
				By.XPATH, f'/html/body/section/div/div/form/div[3]/div/div[2]/div[1]/div/table/tbody/tr[{i}]/td[5]/b/a').get_attribute('href'))
			i += 1
		except NoSuchElementException:
			with ThreadPoolExecutor(75) as p:
				results = p.map(fetch, l)
			for r in list(results):
				if r == []:
					continue
				rera_data.loc[len(rera_data)] = r
			break

	workbook = openpyxl.load_workbook("format.xlsx")
	writer = pandas.ExcelWriter('RERA_DATA.xlsx', engine='openpyxl')
	writer.book = workbook
	writer.sheets = dict((ws.title, ws) for ws in workbook.worksheets)
	rera_data.to_excel(writer, startrow=5, startcol=1, index=False, header=False)
	
	writer.save()
	writer.close()

	progress.config(value=0)
	run_button.config(state='enabled')
	root.update_idletasks()
	driver.quit()


# labels
district_label = Label(root, text='Choose District in Maharashtra(if any): ')
organisation_label = Label(root, text='Enter Organisation Name (if any): ')
tfsi_label = Label(root, text='Enter Total FSI Greater Than(if any): ')

# widgets
keys = list(district_dict.keys())
district_dropbox = Combobox(root, values=keys, state='readonly')
district_dropbox.current(0)
run_button = Button(root, text='Run', command=on_run, state="enabled")
organisation_entry = Entry(root)
tfsi_entry = Entry(root, textvariable=TFSI_DEFAULT)
progress = Progressbar(root, orient='horizontal', length=100, mode='determinate')

# placement
district_label.pack()
district_dropbox.pack()
organisation_label.pack()
organisation_entry.pack()
tfsi_label.pack()
tfsi_entry.pack()
progress.pack()
run_button.pack()

# application loop
root.mainloop()

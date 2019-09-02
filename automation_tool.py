#imports
import requests
import xml.etree.ElementTree
import sys
import xlrd
import xlwt
import getpass

#prtg import ###
import time
from selenium import webdriver
from bs4 import BeautifulSoup

import time
from selenium import webdriver

#from bs4 import BeautifulSoup
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select  
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from multiprocessing import Queue
#from .packages import chardet
###
excel_filepath=input("Please enter the excel path:- ")


#write
workbook2=xlwt.Workbook("encoding=utf-8")
worksheet=workbook2.add_sheet('Sheet1',cell_overwrite_ok=True)

req_sensor=''
host_base_url=''
device_id=''
incident=''
fetchdata=0

print("")
prtgusername=input("Please Enter PRTG Username :-")
prtgpassword=getpass.getpass("Please Enter PRTG password :-")

print("")
#servicenow login credential
snuser=input("Please Enter ServiceNow Username :-")
snpwd=getpass.getpass("Please Enter ServiceNow Username :-")

#driver
driver_init_status=0
ignore_verification=1

#servicenow
servicenow_login = 0 
sensor_status = 0
resolve_status=0
selectval=''
ser_username="ctl00_ContentPlaceHolder1_UsernameTextBox"
ser_password="ctl00_ContentPlaceHolder1_PasswordTextBox"
ser_signin='//*[@id="ctl00_ContentPlaceHolder1_SubmitButton"]'
servicenow_login=0 #to check if login is done or not



def url_builder(host_base_url,device_id):
	hypertext='https://'
	add_api='/api/table.xml?'
	content='&content=sensortree'

	url=hypertext+host_base_url+add_api+device_id+content
	print("The build url is: ",url)

	return url
#to be build with many and optional arg
def indentations(message):
		print("")
		print("")
		#time.sleep(0.5)
		print(message)


#call excel_array,nrows=read_excel(filepath)
def read_excel(excel_filepath):
	workbook = xlrd.open_workbook(excel_filepath)
	sheet1 = workbook.sheet_by_name('Sheet1')
	nrows = sheet1.nrows
	i=0
	excel_array=[]
	#for row in range(nrows-1): #edited on 19031701
	for row in range(0,nrows): #edited on 19031701

		excel_array.append([])
		excel_array[i]=sheet1.row_values(i)
		i=i+1
	indentations("excel read successful...")
	indentations("printing excelarray data...") #test to be deleted
	print(excel_array) #test to be deleted
	#time.sleep(3) #test to be deleted  #commented #20031642

	return excel_array,nrows

#call result=get_prtghost_data(url,prtgusername,prtgpassword)
def get_prtghost_data(url,prtgusername,prtgpassword):
	try:
		result=requests.get(url,params={'username':prtgusername,'password':prtgpassword},verify=False,stream=True)
		indentations("Prtg host fetch data successful...")
		fetchdata=1
	except requests.exceptions.RequestException:
		print("Error 404...")
		fetchdata=0
		result=0
	return result,fetchdata #response object

#call raw_prtg_data_write(result)
def raw_prtg_data_write(result):
	file=open('d:\\new folder\\result.txt','wb') #wb was to open a file in write modein binary
	file.write(result.content)
	indentations("Response object written successfully...")


#call allsensordata=prtg_xml_parse(result)
def prtg_xml_parse(result):

	tree=xml.etree.ElementTree.fromstring(result.content)  #changing from file to direct response object
	#tree=xml.etree.ElementTree.parse('D:\\New folder\\result.txt')
	# root=tree.getroot() works for text but for response .contemt no need so error
	root=xml.etree.ElementTree.fromstring(result.content) #changing from file to direct response object
	print("The root element is :",root)

	# each for loop is for the hierarchy
	allsensordata=[]
	i=0
	print("entering for last loop")
	for sensortree in root.findall('sensortree'):
		print("in prtg tag")
		for nodes in sensortree.findall('nodes'):
			print("in node tag")
			for deviceid in nodes.findall('device'):
				print("in device id")
				#allsensordata.append([])   #for 2d array therefore 2ndlast loop
				for sensorid in deviceid.findall('sensor'):
					name=sensorid.find('name').text
					status=sensorid.find('status').text
					print(name,status)
					allsensordata.append([])   #for 2d array therefore 2ndlast loop
					allsensordata[i].append(name)
					allsensordata[i].append(status)
					i=i+1

	print("Printing in prtg_xml_parse..." + "\n"+"\n",allsensordata)
	#time.sleep(3) # to be deleted #commented #20031642
	indentations("Data Parsed Successfully...")
	return allsensordata

#call sensordata,found_sensor_status = req_sensor_status(allsensordata,req_sensor)
def req_sensor_func(allsensordata,req_sensor):
	print("\n"+ "\n"+  " test in req sensor func...")
	#print("\n"+ "\n"+  " printing allsensordata...") #commented on 19031540
	#print("\n"+ "\n",allsensordata) #commented on 19031540
	print("\n"+ "\n"+ "printing req sensor..." + "\n"+ "\n",req_sensor)

	i=0
	sensordata=[]
	for row in allsensordata:

		if (allsensordata[i][0].strip().lower() == req_sensor.strip().lower() ):
			sensordata=allsensordata[i]
			found_sensor_status="Successful in fetching Required Sensor. "
			autostatus="Successful in fetching the sensor"
		i=i+1	


	if( len( sensordata ) == 0 ):
			sensordata=['no data', 'no data']
			found_sensor_status="Faild to find the Required Sensor."
			autostatus="Failed to Fetch the required Sensor"
			print("\n"+"\n")
	print(sensordata)
	return sensordata,found_sensor_status,autostatus

# sensor_status=sensor_status(sensordata)
def sensor_status_func(sensordata):
	if(sensordata[1].strip().lower()=='up'):
		sensor_status=1
		print("\n"+"\n"+"Sensor status is up and value of sensor_status is:",sensor_status)
		autostatus="The Sensor is UP"
	else: 
		sensor_status=0
		print("\n"+"\n"+"Sensor status is down and value of sensor_status is:",sensor_status)
		autostatus="The Sensor is Down"


	#print("\n"+"\n"+"The status of sensor is:",sensor_status)  #commented on 19031540
	return sensor_status,autostatus

# call driver=driver_object_init(ignore_verification)
def driver_object_init(ignore_verification):
	if ( ignore_verification==1 ):
		options=webdriver.ChromeOptions()
		options.add_argument('--ignore-certificate-errors')
		driver=webdriver.Chrome(chrome_options=options)
	else:
		driver=webdriver.Chrome()

	driver_init_status=1
	indentations("Driver Initialized successfully...")
	return driver,driver_init_status


#call driver,servicenow_login=servicenow_login_func(driver,ser_username,ser_password,ser_signin)
def servicenow_login_func(driver,ser_username,ser_password,ser_signin,snuser,snpwd): # wait code to be found and written
	
	indentations("Connecting to ServiceNow Instance...")
	driver.get("https://servicenow_link")
	wait(driver, 25).until(EC.presence_of_element_located((By.ID, ser_username))) #!!!! searching the incident # newline 16452003

	indentations("Entering Credentials...")
	#time.sleep(3) 
	element=driver.find_element_by_id(ser_username)
	element.send_keys(snuser)
	element2=driver.find_element_by_id(ser_password)
	element2.send_keys(snpwd)
	driver.find_element_by_xpath(ser_signin).click()
	
	servicenow_login = 1
	print("\n"+"\n"+"Login Successful  value of servicenow_login is...",servicenow_login)
	return driver,servicenow_login

'''repeated  servicenow_incidentent_status_check
#call driver,selectval = servicenow_incident_status_check(driver)
def servicenow_incident_status_check(driver):
	wait(driver, 10).until(EC.presence_of_element_located((By.ID, "sysparm_search"))) #!!!! searching the incident
	
	indentations("Searching for the Incident...")
	element3=driver.find_element_by_id("sysparm_search")
	element3.clear()
	element3.send_keys(Incident + Keys.ENTER)
	
	wait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("gsft_main"))
	wait(driver, 10).until(EC.presence_of_element_located((By.ID, "incident.state")))

	indentations("Checking the status of the ticket")
	driver,selectval = service_now_element_type_check_and_status(driver):
	
	return driver,selectval
'''

#call driver,selectval,autostatus,resolve_status = service_now_element_type_check_and_status(driver):
def servicenow_incident_status_check(driver):

	try:
		wait(driver, 25).until(EC.presence_of_element_located((By.ID, "sysparm_search"))) #!!!! searching the incident
		time.sleep(2)

		indentations("Searching for the Incident...")
		element3=driver.find_element_by_id("sysparm_search")
		element3.clear()
		element3.send_keys(Incident + Keys.ENTER)
		
		wait(driver, 15).until(EC.frame_to_be_available_and_switch_to_it("gsft_main"))
		wait(driver, 15).until(EC.presence_of_element_located((By.ID, "incident.state")))
		time.sleep(5) #test as below implicit wait is not working #@@this is working but find a workaround

		#driver.implicitly_wait(10)  #newline added #19031938 as there was incident.state no such elementfound error
		#time.sleep(5) # for testing to be deleted as element_type is always showing text
		indentations("Checking the status of the ticket")
		
		""" need to work on it
		if driver.find_element_by_id("sys_readonly.incident.state"): #newline addedd becoz if the incident is closed it's still shows as select type
			print("this condition is runing can code baseon this")
		"""


		element_type=driver.find_element_by_id("incident.state").get_attribute("type")
		print("\n"+"\n"+"Printing the type of incident.state element",element_type) #test to be deleted #19031710
		time.sleep(4) #test to be deleted #19031710
		selectval=""   #copied from normal.py# 1531344 new line as when no server is found timeout exception is raised so putting if else for that and when it is not found selectval which is returmed on next line will not wriitten nodata error but blank
		
		if( element_type != "text" and element_type != "input"):
			select = Select(driver.find_element_by_id('incident.state')) #it's input but kept variable name as select to compare
			selectval=select.first_selected_option.text
			print("\n"+"\n"+"in incident check status in if that is select,val of selectval is...",selectval)
		else:
			select=driver.find_element_by_id('incident.state')
			selectval=select.get_attribute('value')
			print("\n"+"\n"+"in incident check status in else that is text input,val of selectval is...",selectval)

		driver.switch_to.default_content()  #19031449 #coz can get frame issue

		time.sleep(5) #for testing del it

		if(selectval.strip().lower()== 'resolved' or selectval.strip().lower()== 'closed' or selectval.strip().lower()== '7'):   #!!!already resolved
			indentations("Ticket is already resolved.So, Moving on to Next Incident")
			autostatus="already in resolved state"
			resolve_status=1 #1 means already resolved
			print("\n"+"\n"+"in incident check status in if that is resolve_status,val of selectval is...",resolve_status)

		else:
			resolve_status=0
			autostatus="The ticket is not in resolved state"
			print("\n"+"\n"+"in incident check status in else that is resolve_status,val of selectval is...",resolve_status)
		
	except NoSuchElementException:
		autostatus="The connection was slow hence unable to find element"
		resolve_status=0
		selectval=""

	return driver,selectval,autostatus,resolve_status

#--if code(number 34)
#if(resolve_status==0) 
#line 2 call  driver,resolve_status,Incident = servicenow_incident_resolve(driver,selectval,Incident)
def servicenow_incident_resolve(driver,selectval,Incident):
	#if(selectval.strip().lower()== 'open' or selectval.strip().lower()== 'work in progress'):
	#if only open and work in progress tickets are to be resolved
	print("")
	print("")
	print("Starting the process of resolving the ticket...",Incident)
	wait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it("gsft_main"))

	wait(driver, 10).until(EC.presence_of_element_located((By.ID, "incident.state")))

	select = Select(driver.find_element_by_id('incident.state')) #copied from normal.py#error select not defined so put this line#1432037
	select.select_by_visible_text('Resolved')
	#driver.find_element_by_xpath('//*[@id="resolve_incident"]').click() #commented 19031525
	#---tab code
	#try:
	time.sleep(5)  # as error that the element is not visible
	wait(driver, 10).until(EC.presence_of_element_located((By.ID, "incident.close_code"))) #@@can be frame issue as clicking resolve button #@@ also there can be issue if tabs are used for resolve info
	#except:
	 	#driver.find_element_by_xpath('//*[@id="tabs2_section"]/span[3]/span').click()
	 	#xpath to be found 
	 	#wait(driver, 10).until(EC.presence_of_element_located((By.ID, "incident.close_code"))) #@@can be frame issue as clicking resolve button #@@ also there can be issue if tabs are used for resolve info
	#--tab code end--

	select2 = Select(driver.find_element_by_id('incident.close_code'))
	select2.select_by_visible_text('Solved (Permanently)')
	select3 = driver.find_element_by_id("incident.close_notes")#selectnoneed
	select3.send_keys("This is a False Alert, Hence resolving the ticket " )
	time.sleep(2)
	driver.find_element_by_xpath('//*[@id="resolve_incident"]').click()
	autostatus='Ticket is been resolved'
	resolve_status=1
	time.sleep(5)
	driver.switch_to.default_content()  #19031449
	return driver,resolve_status,Incident,autostatus
#print(".................",sensor_status,".................")
#--else #(34) no need of else


#call req_sensor,host_base_url,device_id,Incident = excel_val_init(i)
def excel_val_init(i):
	print("\n"+"\n"+"test in excel_val_init::: value of i : is ",i) #to be deeted #19031637
	req_sensor=excel_array[i][0]
	host_base_url=excel_array[i][1]
	device_id=excel_array[i][2]
	Incident=excel_array[i][3]

	return req_sensor,host_base_url,device_id,Incident

#call req_sensor,host_base_url,device_id,Incident = excel_val_init(i)
def excel_val_empty():
	req_sensor=''
	host_base_url=''
	device_id=''
	Incident=''

	return req_sensor,host_base_url,device_id,Incident

def newexcel_write(excel_array,autostatus):
	
	j=0
	for x in range(0,4): #4 iteration
	
	#noneed#if(j <3): #will run till 3 that is val 0 sensor val 1 hosturl val2 incident no
		#var[j]=excel_array[i-1][j] #edited on 19031600
		var[j]=excel_array[i][j]

		j=j+1
	var[4]=  autostatus #if we write this in the loop it will give array out of bound error for val
	
	j=0
	for x in range(0,5):
		worksheet.write(i-1,j,var[j]) #coz i starts from 1 and is incremented earlier start of for loop
		j=j+1 #incrementing only j coz i will be incremented from the outer loop

	#how can we give name to excel if it is already given
	workbook2.save("D:\\New folder\\Work\\New folder\\created.xls")
excel_array,nrows=read_excel(excel_filepath)
var=['','','','','']

i=1 #we don't want column name
 
for x in range(0,nrows-1): #as no column name so one row less  
#for x in excel_array:

	
	req_sensor,host_base_url,device_id,Incident = excel_val_init(i) #edited on 19031641
	#req_sensor,host_base_url,device_id,Incident = excel_val_init(i-1)

	#test
	print("checking the excel 1st row"+"\n"+"\n")
	print(req_sensor,host_base_url,device_id,Incident)
	time.sleep(5)

	url=url_builder(host_base_url,device_id)

	result,fetchdata=get_prtghost_data(url,prtgusername,prtgpassword)

	if (fetchdata==1 ):

		raw_prtg_data_write(result)
		allsensordata=prtg_xml_parse(result)
		sensordata,found_sensor_status,autostatus = req_sensor_func(allsensordata,req_sensor)
		sensor_status,autostatus=sensor_status_func(sensordata)

		print("\n"+"\n"+"test driver_init_status...",driver_init_status)
		if(driver_init_status==0 ):
			driver,driver_init_status=driver_object_init(ignore_verification)

		print("\n"+"\n"+"test service now login...",servicenow_login)
		if (servicenow_login == 0 ):
			driver,servicenow_login=servicenow_login_func(driver,ser_username,ser_password,ser_signin,snuser,snpwd)

		print("\n"+"\n"+"test sensor status...",sensor_status)
		if (sensor_status == 1 ):
		    indentations("hence now entering the servicenow incident status check functions...")
		    driver,selectval,autostatus,resolve_status = servicenow_incident_status_check(driver)

		print("\n"+"\n"+"test sensor status and resolve status...",servicenow_login,resolve_status)
		if ( sensor_status==1 and resolve_status==0 ):
			indentations("The sensor is up and the incident is open so resolving the incident...entering servicenow_incident_resolve...")
			driver,resolve_status,Incident,autostatus = servicenow_incident_resolve(driver,selectval,Incident)
		
		newexcel_write(excel_array,autostatus)

		i=i+1



		print(".................",sensor_status,found_sensor_status,".................")
	
	else:
		indentations("error 404...moving to next incident")
	#if(i%21==0):
		#print("closing ssshhhhhhhh..... ")
		#driver.implicitly_wait(120)  #newline added #19031938 as there was incident.state no such elementfound error
		#driver.close()

#driver.close()

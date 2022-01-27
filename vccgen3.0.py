import requests
import random
import lxml
import json
import time
import xlwt
from xlwt import Workbook
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from lxml import html

def trunc_gauss(mu, sigma, bottom, top):
    a = random.gauss(mu,sigma)
    while (bottom <= a <= top) == False:
        a = random.gauss(mu,sigma)
    return a

class Generate():
	def __init__(self):
		options = webdriver.ChromeOptions() 
		options.add_experimental_option("excludeSwitches", ["enable-logging"])
		self.w = webdriver.Chrome(options=options)
		self.count = 0

	def login(self,user,pword):
		web = self.w.get('https://getgo.tradeshift.com/signin/login')

		print('Logging in...')

		time.sleep(10)

		print('Typing Username...')
		UserName = user
		user = self.w.find_element_by_xpath('//*[@id="go-app"]/div/div/div/div/div/div[2]/form/fieldset[1]/label/input')
		for char in UserName:
			user.send_keys(char)
			time.sleep(trunc_gauss(.1,.01,.07,.3))

		
		time.sleep(5)

		print('Typing Password...')
		Pass = pword
		pw = self.w.find_element_by_xpath('//*[@id="go-app"]/div/div/div/div/div/div[2]/form/fieldset[2]/label/input')
		for char in Pass:
			pw.send_keys(char)
			time.sleep(trunc_gauss(.1,.01,.07,.3))


		time.sleep(2)
		
		LoginButton = self.w.find_element_by_xpath('//*[@id="go-app"]/div/div/div/div/div/div[2]/form/fieldset[4]/menu/li[1]/button')
		LoginButton.click()

	def createRequestMultiple(self, start):
		cnt = int(start)
		self.count = cnt
		CRButton = self.w.find_element_by_xpath('//*[@id="js-sidebar-new"]/div[2]/menu[1]/li[7]/a')
		CRButton.click()
		print('Creating Request...')

		time.sleep(trunc_gauss(17,1,15,20))

		print('Typing Purpose...')
		Purpose = self.count
		purp = self.w.find_element_by_xpath('//*[@id="hash4"]/event-purchases-request-form/div/purchase-request-form-v2/div/form/fieldset[1]/label/input')
		purp.send_keys(Purpose)


		time.sleep(trunc_gauss(3,1,1,5))

		print('Typing Amount...')
		Amount = ["1","0","0","0","0"]
		amt = self.w.find_element_by_xpath('//*[@id="hash4"]/event-purchases-request-form/div/purchase-request-form-v2/div/form/ts-input/fieldset/label/input')
		for x in Amount:
			time.sleep(trunc_gauss(.1,.01,.07,.3))
			amt.send_keys(x)

		time.sleep(trunc_gauss(4,1,3,6))

		SubmitButton = self.w.find_element_by_xpath('/html/body/div/div/div/div[1]/div/conversation/div/object-event[2]/div/div/div/event-purchases-request-form/div/purchase-request-form-v2/div/form/fieldset[5]/menu/li/button')
		SubmitButton.click()

		print('Request Created')

		time.sleep(trunc_gauss(17,1,15,20))

		print('Accessing Approvals...')
		
		Approvals = self.w.find_element_by_xpath('/html/body/go-drawer/div[2]/menu[1]/li[3]/a')
		Approvals.click()

		time.sleep(trunc_gauss(6,1,5,7))

		CreateCard = self.w.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div/div/button[1]')
		CreateCard.click()

		print('Creating Card...')

		time.sleep(trunc_gauss(6,1,5,7))

		FinalizeCard = self.w.find_element_by_xpath('/html/body/aside[7]/div[1]/div/div/form/fieldset[3]/menu/li[1]/button')
		FinalizeCard.click()

		self.count +=1
		return self.count

	def createRequestSubscription(self):
		self.count
		CRButton = self.w.find_element_by_xpath('//*[@id="js-sidebar-new"]/div[2]/menu[1]/li[7]/a')
		CRButton.click()
		print('Creating Request...')

		time.sleep(trunc_gauss(17,1,15,20))

		Purpose = self.count
		purp = self.w.find_element_by_xpath('//*[@id="hash4"]/event-purchases-request-form/div/purchase-request-form-v2/div/form/fieldset[1]/label/input')
		purp.send_keys(Purpose)

		time.sleep(trunc_gauss(3,1,1,5))

		Amount = ["1","0","0","0","0"]
		amt = self.w.find_element_by_xpath('//*[@id="hash4"]/event-purchases-request-form/div/purchase-request-form-v2/div/form/ts-input/fieldset/label/input')
		for x in Amount:
			time.sleep(trunc_gauss(.1,.01,.07,.3))
			amt.send_keys(x)

		time.sleep(trunc_gauss(3,1,1,5))

		CardTypeButton = self.w.find_element_by_xpath('//*[@id="hash4"]/event-purchases-request-form/div/purchase-request-form-v2/div/form/fieldset[2]/label/input')
		CardTypeButton.click()
		time.sleep(trunc_gauss(1,.1,.8,3))

		SubButton = self.w.find_element_by_xpath('/html/body/aside[3]/div/menu/li[3]')
		SubButton.click()
		time.sleep(trunc_gauss(1,.1,.8,3))

		FreqButton = self.w.find_element_by_xpath('//*[@id="hash4"]/event-purchases-request-form/div/purchase-request-form-v2/div/form/fieldset[3]/label/input')
		FreqButton.click()
		time.sleep(trunc_gauss(1,.1,.8,3))

		FreqType = self.w.find_element_by_xpath('/html/body/aside[3]/div/menu/li[3]')
		FreqType.click()
		time.sleep(trunc_gauss(7,1,5,10))

		SubmitButton = self.w.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div/conversation/div/object-event[2]/div/div/div/event-purchases-request-form/div/purchase-request-form-v2/div/form/fieldset[6]/menu/li/button')
		SubmitButton.click()

		print('Request Created')

		time.sleep(trunc_gauss(17,1,15,20))

		print('Accessing Approvals...')
		
		Approvals = self.w.find_element_by_xpath('/html/body/go-drawer/div[2]/menu[1]/li[3]/a')
		Approvals.click()

		time.sleep(trunc_gauss(6,1,5,7))

		CreateCard = self.w.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[1]/div[2]/div/div/button[1]')
		CreateCard.click()

		print('Creating Card...')

		time.sleep(trunc_gauss(6,1,5,7))

		FinalizeCard = self.w.find_element_by_xpath('/html/body/aside[7]/div[1]/div/div/form/fieldset[3]/menu/li[1]/button')
		FinalizeCard.click()

		self.count +=1
		return self.count

	def loadWallet(self):

			print('Loading Wallet...')
			wallet = self.w.find_element_by_xpath('/html/body/go-drawer/div[2]/menu[1]/li[2]/a')
			wallet.click()

			time.sleep(10)

	def scrapeCards(self,name):
		print('Scraping Cards...')
		wb = Workbook()

		sheet1 = wb.add_sheet('Sheet 1')
		sheet1.write(0,0,'CARD NAME')
		sheet1.write(0,1,'CARD NUMBER')
		sheet1.write(0,2,'CVV')

		for x in range (1,125):
			cardTitle = self.w.find_element_by_xpath('/html/body/div/div/div/div[1]/wallet/div/div['+str(x)+']/wallet-list-item/div/div[2]/div[1]').text
			print('Card Name: ' + cardTitle)
			sheet1.write(x,0,cardTitle)
			cardNumber = self.w.find_element_by_xpath('/html/body/div/div/div/div[1]/wallet/div/div['+str(x)+']/wallet-list-item/div/div[2]/virtual-card/div/div/div[4]').text
			cardNumber = cardNumber.replace("\r","")
			cardNumber = cardNumber.replace("\n","")
			sheet1.write(x,1,cardNumber)
			cardCVV = self.w.find_element_by_xpath('/html/body/div/div/div/div[1]/wallet/div/div['+str(x)+']/wallet-list-item/div/div[2]/virtual-card/div/div/div[5]/div[2]/div[2]').text
			sheet1.write(x,2,cardCVV)
			print('Card Number: ' + cardNumber)
			print('CVV: ' + cardCVV)
			time.sleep(1)

		for x in range (127,148):
			cardTitle = self.w.find_element_by_xpath('/html/body/div/div/div/div[1]/wallet/div/div['+str(x)+']/wallet-list-item/div/div[2]/div[1]').text
			print('Card Name: ' + cardTitle)
			sheet1.write(x,0,cardTitle)
			cardNumber = self.w.find_element_by_xpath('/html/body/div/div/div/div[1]/wallet/div/div['+str(x)+']/wallet-list-item/div/div[2]/virtual-card/div/div/div[4]').text
			cardNumber = cardNumber.replace("\r","")
			cardNumber = cardNumber.replace("\n","")
			sheet1.write(x,1,cardNumber)
			cardCVV = self.w.find_element_by_xpath('/html/body/div/div/div/div[1]/wallet/div/div['+str(x)+']/wallet-list-item/div/div[2]/virtual-card/div/div/div[5]/div[2]/div[2]').text
			sheet1.write(x,2,cardCVV)
			print('Card Number: ' + cardNumber)
			print('CVV: ' + cardCVV)
			time.sleep(1)

		wb.save(name + '.xlsx')

	def quit(self):
		print('Closing window.')
		self.w.quit()

print('Opening Window...')
generate = Generate()
print('Please Enter Username:')
user = input()
print('Please Enter Password:')
pw = input()
generate.login(user,pw)
userInput = ''
while userInput != 'quit':
	print('Enter 1 to generate VCC')
	print('Enter 2 to export current VCC')
	print('Enter quit to quit')
	userInput = input()
	if userInput == '1':
		print('How many times would you like to generate? (Max 50)')
		cnt = input()
		print('Enter Starting Point:')
		numCards = input()

		for x in range(0,int(cnt)):
			time.sleep(trunc_gauss(30,1,24,35))
			generate.createRequestMultiple(numCards)
			time.sleep(trunc_gauss(30,1,24,35))
			generate.createRequestSubscription()
			
		continue
		
	if userInput == '2':
		print('Please enter name for save file')
		saveName = input()
		generate.loadWallet()
		generate.scrapeCards()
		continue

	if userInput == 'quit':
		generate.quit()

	else:
		print('Please use a valid input.')
		time.sleep(1)
		continue
	



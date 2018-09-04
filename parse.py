# -*- encoding: utf8 -*-

from urllib.request import urlopen
from bs4 import BeautifulSoup
from threading import Thread
from openpyxl import Workbook
from os import makedirs, path

grades = list()
finished = list()
movieCode = 0
maxPage = 0

def parseReply(replyList) :
	result = list()

	for reply in replyList :
		text  = reply.find(class_='score_reple').find('p').text.strip()
		if reply.find(class_='score_reple').find('p').find(class_='ico_viewer') is not None:
			text = text[3:]

		score = int(reply.find(class_='star_score').find('em').string)
		result.append((score, text))
	
	return result

def getOnePage(page) :
	print("getting page", page)

	url = 'https://movie.naver.com/movie/bi/mi/pointWriteFormList.nhn?code=' + str(movieCode) + '&type=after&isActualPointWriteExecute=false&isMileageSubscriptionAlready=false&isMileageSubscriptionReject=false&page=' + str(page)

	soup = BeautifulSoup(urlopen(url).read(), 'html.parser')
	replyList = soup.find(class_='score_result').findAll('li')

	return parseReply(replyList)

def getPages(ones) :
	print("start getting page. ones :", ones)
	for tens in range(0, maxPage//10) :
		try :
			grades.extend(getOnePage(tens * 10 + ones))
		except :
			print("catch exception on", (tens * 10 + ones))
			pass
	finished.append(ones)
	print("end getting page. ones :", ones)
	saveToExcel()

def saveToExcel() :				# callback
	print("called save to excel. finished :", len(finished))
	if len(finished) == 10 :
		print("start saving to excel")
		wb_learn = Workbook()
		wb_test  = Workbook()
		ws_learn = wb_learn.active
		ws_test  = wb_test.active

		ws_learn['A1'] = "평점"
		ws_learn['B1'] = "댓글"
		ws_test['A1'] = "평점"
		ws_test['B1'] = "댓글"

		for i in range(len(grades) * 8 // 10) :
			ws_learn['A' + str(i + 2)] = grades[i][0]	# score
			ws_learn['B' + str(i + 2)] = grades[i][1]	# text
		
		testStart = len(grades) * 8 // 10

		for i in range(testStart, len(grades)) :
			ws_test['A' + str(i - testStart + 2)] = grades[i][0]	# score
			ws_test['B' + str(i - testStart + 2)] = grades[i][1]	# text
		
		if path.isdir(str(movieCode)) == False :
			makedirs(str(movieCode))

		wb_learn.save("./" + str(movieCode) + "/sentence_learn.xlsx")
		wb_test.save("./" + str(movieCode) + "/sentence_test.xlsx")
		print("finish saving to excel")

def getPageCount() :
	global maxPage

	print("start getting page count")

	url = 'https://movie.naver.com/movie/bi/mi/pointWriteFormList.nhn?code=' + str(movieCode) + '&type=after&isActualPointWriteExecute=false&isMileageSubscriptionAlready=false&isMileageSubscriptionReject=false&page=999999'
	soup = BeautifulSoup(urlopen(url).read(), 'html.parser')

	maxPage = int(soup.find(id='page')['value'])

	print("end getting page count. max page :", maxPage)

def main() :
	global movieCode
	movieCode = int(input("movie code : "))

	getPageCount()

	for ones in range(1, 11) :
		Thread(target=getPages, args=(ones,)).start()

if __name__ == '__main__' :
	main()
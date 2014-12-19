import xlrd
import urllib, re
import urllib2
import requests
import csv
import sys
import HTMLParser

default_encoding = 'utf-8'
if sys.getdefaultencoding() != default_encoding:
	reload(sys)
	sys.setdefaultencoding(default_encoding)

FILE_LOCATION = "/Users/yaxing/Desktop/WriterContacts.xls"

emaillist = []

def write_data_to_csv(data):

	with open('test.csv', 'w') as fp:
	    a = csv.writer(fp, delimiter=',')
	    a.writerows(data)

def find_email(url):

	htmlFile = requests.get(url)
	html = htmlFile.text

	regexp_email = r'mailto:((&#\d+;)+)'
	pattern = re.compile(regexp_email)
	emailAddresses = re.findall(pattern, html)

	if emailAddresses:
		h = HTMLParser.HTMLParser()
		email = h.unescape(emailAddresses[0][0]) 		
	else: 
		return -1

	print email
	return email

def main():
	workbook = xlrd.open_workbook(FILE_LOCATION, formatting_info = True)
	sheet = workbook.sheet_by_index(1)

	for row_index in range(6001, sheet.nrows):
		user = []
		name = sheet.cell_value(row_index, 0)
		
		print name
		link = sheet.hyperlink_map.get((row_index, 0))

		URL = link.url_or_path
		email = find_email(URL)

		if email:
			user.append(name)
			user.append(email)
			emaillist.append(user)
		else:
			continue

	write_data_to_csv(emaillist)

if __name__ == '__main__':
	main()

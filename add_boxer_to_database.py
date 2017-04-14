import xlrd
import re #regular expresions for python
import glob, os
from pymongo import MongoClient

client = MongoClient()
db = client.savate

def get_registered_page(data_sheet):
	for x in range(0, data_sheet.nsheets):
		current = data_sheet.sheet_by_index(x)
		if re.search("inscrit", current.name.lower()):
			return current

def insert_boxers(data_sheet):
	registered = get_registered_page(data_sheet)
	weight = 0 #col for weights
	dep = 0 #col for department
	name = 0 #col for full name
	gym	= 0 #col for the club
	for x in range(registered.nrows): # iterating on the registered page to get indexes of the name, dep, club and weight values
		for y in range(registered.ncols):
			if registered.cell_type(x,y) == 1:
				if re.search("poid", registered.cell_value(x,y).lower()):
					weight = y
				if re.search("dept", registered.cell_value(x,y).lower()):
					dep = y
				if re.search("nom", registered.cell_value(x,y).lower()):
					name = y
				if re.search("club", registered.cell_value(x,y).lower()):
					gym = y
	for x in range(2, registered.nrows):
		user = {}
		if registered.cell_type(x,weight) == 1:
			user['weight'] = registered.cell_value(x,weight)
			user['dep'] = int(registered.cell_value(x,dep))
			user['name'] = registered.cell_value(x,name)
			user['gym'] = registered.cell_value(x,gym)
			print (user)

os.chdir("data")
for file in glob.glob("*.xls"):
	data_sheet = xlrd.open_workbook(file)
	insert_boxers(data_sheet)
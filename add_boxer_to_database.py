import xlrd
import re #regular expresions for python
import string
import glob, os
from pymongo import MongoClient

client = MongoClient()
db = client.savate
users = db.users
fights = db.fights

regex = re.compile('[%s]' % re.escape(string.punctuation))

def get_registered_page(data_sheet):
	for x in range(0, data_sheet.nsheets):
		current = data_sheet.sheet_by_index(x)
		if re.search("inscrit", current.name.lower()):
			return current
	return data_sheet.sheet_by_index(0)

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
			user['weight'] = registered.cell_value(x,weight).strip()
			user['dep'] = int(registered.cell_value(x,dep))
			user['name'] = re.sub("\s+", " ", registered.cell_value(x,name).strip().lower())
			user['gym'] = regex.sub('', registered.cell_value(x,gym).strip().lower())
			users.update(user, user, upsert=True)

def get_match_types(file):
	# print (file)
	if re.search("lite", file) or re.search("erium", file) or re.search("espoir", file):
		return "combat"
	else:
		return "assaut"


def insert_matches(data_sheet, file):
	for x in range (0, data_sheet.nsheets):
		current = data_sheet.sheet_by_index(x)
		if re.search('[MF][0-9]{2,3}', current.name): #regex used to find if the current styleSheet name is one of a CATEGORY
			for y in range(0, current.nrows):
				if re.search('[0-9]{1}.(TOUR)', current.cell(y, 0).value) or re.search('[0-9]{1}.(TOUR)', current.cell(y - 1, 0).value):
					place = 0
					fight = {}
					for z in range(0, current.ncols):
						if type(current.cell(y, z).value) is not float and re.search('[a-zA-Z]{1,}\s{1}[a-zA-Z]*', current.cell(y, z).value):
							if place == 0:
								red_boxer = current.cell(y, z).value
								fight['red'] = current.cell(y, z).value.strip().lower()
							if place == 1:
								blue_boxer = current.cell(y, z).value
								fight['blue'] = current.cell(y, z).value.strip().lower()
							if place == 2:
								forfait = 0
								ko = 0
								winner = current.cell(y, z).value
								if winner.lower().find('forfait') > -1:
									forfait = 1
								if winner.lower().find('h.c') > -1:
									ko = 1
								if winner.lower().find('combat') > -1:
									ko = 1
								if winner.find(red_boxer) > -1:
									fight['winner'] = 'red'
								if winner.find(blue_boxer) > -1:
									fight['winner'] = 'blue'
								if forfait > 0:
									fight['victory'] = 'forfait'
								elif ko == 1:
									fight['victory'] = 'ko'
								else:
									fight['victory'] = 'decision'
								if 'winner' in fight:
									fight['cat'] = current.name
									fight['type'] = get_match_types(current.cell(0,0).value.lower())
									fight['competition'] = regex.sub('', current.cell_value(0,0).strip().lower())
									fights.update(fight, fight, upsert=True)
									break
							place = place + 1

os.chdir("data")

for file in glob.glob("*.xls*"):
	print("File ", file)
	data_sheet = xlrd.open_workbook(file)
	insert_boxers(data_sheet)
	insert_matches(data_sheet, file.lower())
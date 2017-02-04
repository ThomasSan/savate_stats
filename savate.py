import xlrd
import re
import glob, os


blue = 0
red = 0
red_boxer = ""
blue_boxer = ""

def get_book_stats(data_sheet, red, blue):
	for x in range (0, data_sheet.nsheets):
		current = data_sheet.sheet_by_index(x)
		if re.search('[MF][0-9]{2,3}', current.name): #regex used to find if the current styleSheet name is one of a CATEGORY
			# print current.name
			for y in xrange(0, current.nrows):
				if re.search('[0-9]{1}.(TOUR)', current.cell(y, 0).value) or re.search('[0-9]{1}.(TOUR)', current.cell(y - 1, 0).value):
					place = 0
					for z in range(0, current.ncols):
						if type(current.cell(y, z).value) is not float and re.search('[a-zA-Z]{1,}\s{1}[a-zA-Z]*', current.cell(y, z).value):
							if place == 0:
								red_boxer = current.cell(y, z).value
								# print "red " + red_boxer
							if place == 1:
								blue_boxer = current.cell(y, z).value
								# print "blue " + blue_boxer
							if place == 2:
								forfait = 0
								winner = current.cell(y, z).value
								if winner.lower().find('forfait') > -1:
									forfait = 1
								elif winner.find(red_boxer) > -1:
									red = red + 1
									# print "red"
								elif winner.find(blue_boxer) > -1:
									blue = blue + 1
									# print "blue"
								if forfait > 0:
									print "Forfait\n"
								else:
									print "WINNER " + current.cell(y, z).value + '\n'
								# print "=> BLUE " + str(blue) + "/ =>RED " + str(red) + '\n'
							# print "cell(" + str(y) + "," + str(z) + ") = " + current.cell(y, z).value
							place = place + 1

	result = [red]
	result.append(blue)
	return result

os.chdir("data")
for file in glob.glob("*.xls"):
	data_sheet = xlrd.open_workbook(file)
	res = get_book_stats(data_sheet, red, blue)
	red = res[0]
	blue = res[1]
	print "=> BLUE " + str(blue) + "/ =>RED " + str(red) + '\n'

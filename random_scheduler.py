"""
This script belongs to the Phi Kappa Psi Cal Gamma chapter. 

Use it to generate random door and bar shifts. 

@author Vladislav Karchevsky
@author Ryan Flynn
"""

import random
import xlsxwriter
import os
#import easygui as eg 

# How many shifts?
NUM_SLOTS = 7
DATE = input("What is the date of this event?\n")
#DATE = eg.enterbox("What is the date of the event?", "Date")

def pickRandomBros(broList, numSample):
	# Note that this function affects the original list
	randomSet = random.sample(set(broList), numSample)
	for bro in randomSet:
		broList.remove(bro)
	return randomSet

doorShift1 = []
doorShift2 = []
barShift1 = []
barShift2 = []


# Create brotherhood lists
beta = ["Sam Parks"]
gamma = ["Stefan Isenberger",
		"Max Mathison",
		"Nathan Gomez",
		"Shahrukh Ghazali",
		"Jay Patel",
		"Ali Imani"]
delta = ["Ronak Patel",
		"Jason Johl",
		"Aaron Boussina",
		"Jonathan Hsu",
		"Rehan Hasan"]
epsilon = ["Curtis Siegfried",
		"Mitchell Pok",
		"Paul Levchenko",
		"Conor Stanton",
		"Vlad Karchevsky",
		"Michael Metzler",
		"Manny Sabbagh",
		"Ian Chin",
		"Rehman Minhas"]
zeta = ["Taylor Ferguson",
		"Evan Mason",
		"Eric Gabrielli",
		"Chris Farmer",
		"Zachary Hawtof",
		"Ian Mason",
		"Elliot Surovell",
		"Anurag Reddy",
		"Ryan Flynn",
		"Sam Rausser",
		"Rikesh Patel",
		"Jack Hendershott",
		"Mark Traganza",
		"Han Li",
		"Karan Karia",
		"Evin Wieser",
		"Matthew Buckley",
		"Erik Bartlett",
		"Eric Smith",
		"Spencer Hawley"]
eta = ["Kyle Joyner",
		"Richard Mercer",
		"Andrew Soncrant",
		"Joey Papador",
		"Christian Collins",
		"Anand Dharia",
		"Francisco Torres",
		"Martin Rayburn",
		"Donovan Frazer",
		"Nick Alaverdyan",
		"Mustapha Khokhar",
		"Laith Alqaisi"]
pledges = ["Aman Khan",
		"Andrew Ahmadi",
		"Aneesh Prasad",
		"Ben Kurschner",
		"Christiaan Khurana",
		"Christos Gkolias",
		"Elliot Dunn",
		"Harrison Agrusa",
		"Jack Sweeney",
		"Jason Blore",
		"Jeremy Mack",
		"Joe Labrum",
		"Keeton Ross",
		"Lawrence Dong",
		"Matt Nisenboym",
		"Mitchell Stieg",
		"Nabil Faruoqi",
		"Nathan Aminpour",
		"Reese Levine",
		"Ricky Philipossian",
		"Riley Pok",
		"Rokhan Khan",
		"Sahand Saberi",
		"Thomas Zorilla",
		"Will Morrow"]

# Aggregate brothers gone from social event
permaAbsent = ["Vlad Karchevsky", "Rikesh Patel", "Andrew Soncrant", 
	"Donovan Frazer"]

# -------- Ryan's addition to make excluding absent brothers easier -----------

anyAbsent = input("Will any brothers be absent from this event? ")
#absentMsg = "Will any brothers be absent from this event?"
#absentTtl = "Any Absent?"
#anyAbsent = eg.enterbox(absentMsg, absentTtl)
if anyAbsent.strip().lower() == "yes":
	runAbsentSurvey = True
else:
	runAbsentSurvey = False

absentTonight = []

while runAbsentSurvey:
	name = input("Please enter the name of the absent brother. "+
		"If this is the last absent brother, end with a period. \n")
	#msg = ("Please enter the name of the absent brother. "+
	#	"If this is the last absent brother, end with a period.")
	#ttl = "Absent Brothers"
	#name = eg.enterbox(msg, ttl)
	if name[len(name) - 1] == ".":
		last = True
	else:
		last = False
	if last:
		absentTonight += [name[:-1],]
		runAbsentSurvey = False
	else:
		absentTonight += [name]

# ----------------------- End of Ryan's First Addition ------------------------

absent = permaAbsent + absentTonight

# Create a subset of brothers that can do work during social event
eligibleBros = list(set(eta + zeta + pledges + epsilon))

# -------------------------- Ryan's next addition -----------------------------
# takes into account class in the likelihood of selection
# also writes to excel file

# List of brothers who are good at door
brothers_good_at_door = ["Curtis Siegfried",
						 "Taylor Ferguson",
						 "Evan Mason",
						 "Chris Farmer",
						 "Zachary Hawtof",
						 "Ian Mason",
						 "Elliot Surovell",
						 "Ryan Flynn",
						 "Jack Hendershott",
						 "Han Li",
						 "Matthew Buckley",
						 "Erik Bartlett",
						 "Kyle Joyner",
						 "Richard Mercer",
						 "Andrew Soncrant",
						 "Joey Papador",
						 "Christian Collins",
						 "Anand Dharia",
						 "Francisco Torres",
						 "Nick Alaverdyan",
						 ]

# List of those available by class who are also not in brothers_good_at_door
available_pledges = [pledge for pledge in pledges if pledge \
	not in absentTonight]
available_eta = [eta_mem for eta_mem in eta if not (eta_mem in absentTonight
	or eta_mem in permaAbsent or eta_mem in brothers_good_at_door)]
available_zeta = [zeta_mem for zeta_mem in zeta if not (zeta_mem in 
	absentTonight or zeta_mem in permaAbsent or zeta_mem in 
	brothers_good_at_door)]
available_epsilon = [epsilon_mem for epsilon_mem in epsilon if not (
	epsilon_mem in absentTonight or epsilon_mem in permaAbsent or epsilon_mem in 
	brothers_good_at_door)]
available_delta = [delta_mem for delta_mem in delta if not (
	delta_mem in absentTonight or delta_mem in permaAbsent or delta_mem in 
	brothers_good_at_door)]
available_brothers_good_at_door = [bro for bro in brothers_good_at_door if
	bro not in absentTonight]

# Picks pledges out of available_pledges now so that they will not be given 
# multiple shifts later
doorShift2 = pickRandomBros(available_pledges, NUM_SLOTS)
roofShift = pickRandomBros(available_pledges, 4)

# List of tuples. First item in tuple is the available members of the class.
# Second item is used as a weight so that lower class brothers (and pledges)
# get shifts more often.
lst_availables = [(available_delta, .3), (available_epsilon, .4), 
	(available_zeta, .5), (available_eta, .9), (available_pledges, 1)]

final_lst = []

# Similar to PickRandomBros, but not destructive of broList
def pickRandomBrosFromClass(broList, numSample):
	randomLst = random.sample(broList, numSample)
	return randomLst

# Places a number of bros from each class in the final list.  Higher
# classes get less bros (per population) placed in this list
for lst in lst_availables:
	num_bros = int(len(lst[0]) * lst[1])
	add_lst = pickRandomBrosFromClass(lst[0], num_bros)
	for el in add_lst:
		final_lst += [el,]

# Picks only brothers good at door to work doorShift1 (alongside a pledge)
doorShift1 = pickRandomBros(available_brothers_good_at_door, NUM_SLOTS)

barShift1 = pickRandomBros(final_lst, NUM_SLOTS)
barShift2 = pickRandomBros(final_lst, NUM_SLOTS)

# ---------------- Writes shift assignments to an excel file ------------------

filename = "ShiftSample1.xlsx"
workbook = xlsxwriter.Workbook(filename)
worksheet = workbook.add_worksheet()
worksheet.set_landscape()

title = workbook.add_format()
title.set_font_size(30)
title.set_bold()
title.set_align('center')
names = workbook.add_format()
names.set_font_size(12)
header = workbook.add_format()
header.set_font_size(14)
header.set_bold()
header.set_bottom(1)
header.set_left(1)
time = workbook.add_format()
time.set_font_size(14)
time.set_bold()
time.set_top(1)
time.set_right(1)
names1 = workbook.add_format()
names1.set_font_size(12)
names1.set_border(1)
names2 = workbook.add_format()
names2.set_font_size(12)
names2.set_top(1)
names3 = workbook.add_format()
names3.set_font_size(12)
names3.set_left(1)
names4 = workbook.add_format()
names4.set_font_size(12)
names4.set_left(1)
names4.set_top(1)
empty = workbook.add_format()
empty.set_border(1)
empty.set_bg_color('gray')
empty1 = workbook.add_format()
empty1.set_top(1)
empty1.set_right(1)
empty1.set_left(1)
empty1.set_bg_color('gray')


worksheet.set_column(0, 0, 1)
worksheet.set_column(1, 1, 15)
worksheet.set_column(2, 2, 20)
worksheet.set_column(3, 3, 20)
worksheet.set_column(4, 4, 20)
worksheet.set_column(5, 5, 20)
worksheet.set_column(6, 6, 20)

TITLE = '&C&30&"Calibri,Bold"Phi Psi Door and Bar Shift ' + DATE
worksheet.set_header(TITLE)

worksheet.write(3, 1, "10:00-10:30", time)
worksheet.write(4, 1, "10:30-11:00", time)
worksheet.write(5, 1, "11:00-11:30", time)
worksheet.write(6, 1, "11:30-12:00", time)
worksheet.write(7, 1, "12:00-12:30", time)
worksheet.write(8, 1, "12:30-1:00", time)
worksheet.write(9, 1, "1:00-1:30", time)

worksheet.write(2, 2, "Door 1", header)
row = 3
for bro in doorShift1:
	if row == 9:
		worksheet.write(row, 2, bro, names3)
	else:
		worksheet.write(row, 2, bro, names1)
	row += 1

worksheet.write(2, 3, "Door 2", header)
row = 3
for bro in doorShift2:
	if row == 9:
		worksheet.write(row, 3, bro, names3)
	else:
		worksheet.write(row, 3, bro, names1)
	row += 1

worksheet.write(2, 4, "Roof Shift", header)
row = 5
for i in range(3, 10):
	if i < 5:
		worksheet.write(i, 4, "", empty)
	elif i >= 5 and i < 9:
		worksheet.write(i, 4, roofShift[i-5], names1)
	else:
		worksheet.write(i, 4, "", empty1)

worksheet.write(2, 5, "Bar 1", header)
row = 3
for bro in barShift1:
	if row == 9:
		worksheet.write(row, 5, bro, names3)
	else:
		worksheet.write(row, 5, bro, names1)
	row += 1

worksheet.write(2, 6, "Bar 2", header)
row = 3
for bro in barShift2:
	if row == 9:
		worksheet.write(row, 6, bro, names4)
	else:
		worksheet.write(row, 6, bro, names2)
	row += 1



workbook.close()

os.startfile(filename)

# ------------------------ end of Ryan's contribution -------------------------






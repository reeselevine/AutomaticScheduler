"""
This script belongs to the Phi Kappa Psi Cal Gamma chapter. 

Use it to generate random door and bar shifts. 

@author Vladislav Karchevsky
@author Ryan Flynn
@author Reese Levine
"""

import random
import xlsxwriter
import os
#import easygui as eg 

# How many shifts?
NUM_SLOTS = 7
DATE = raw_input("What is the date of this event?\n")
#DATE = eg.enterbox("What is the date of the event?", "Date")

def pickRandomBros(broList, numSample):
	# Note that this function affects the original list
	randomSet = random.sample(set(broList), numSample)
	for bro in randomSet:
		broList.remove(bro)
	return randomSet

doorShift1 = []
doorShift2 = []
doorShift3 = []
doorShift4 = []
barShift1 = []
barShift2 = []
barShift3 = []
barShift4 = []


# Create brotherhood lists
gamma = ["Nathan Gomez"]
delta = ["Jason Johl",
                "Aaron Boussina"]
epsilon = ["Curtis Siegfried",
		"Mitchell Pok",
		"Paul Levchenko",
		"Conor Stanton",
		"Manny Sabbagh",
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
		"Donovan Frazer",
		"Nick Alaverdyan",
		"Mustapha Khokhar",
		"Laith Alqaisi"]
theta = ["Aman Khan",
		"Andrew Ahmadi",
		"Aneesh Prasad",
		"Ben Kurschner",
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
iota = ["Alex Clark",
                "Kenny Dang",
                "Brent Freed",
                "Jacob Gill",
                "Darius Kay",
                "David Kret",
                "Will Lopez",
                "Dhruv Malik",
                "Ian Moon",
                "Francisco Peralta",
                "Brandt Sheets",
                "Andrew Ting",
                "Evan Wilson"]

pledges = ["Anthony Fortney",
                "Steven Lin",
                "Ford Noble",
                "Ryan Leyba",
                "Robert Mcilhatton",
                "Jonathan",
                "Morris Ravis",
                "Ben Lalezari",
                "Drew Hanson",
                "Josh Bradley-Bevan",
                "Steven Beelar",
                "Gabriel Bogner",
                "Dylan Dreyer",
                "Luke Thomas",
                "Tzartzas Tinos",
                "Nate Parke",
                "Dan Lee",
                "Max Seltzer",
                "Andy Frey",
                "Nathan Kelleher",
                "Arnav Chaturvedi",
                "Sam Giacometti",
                "Sam Bauman"]

# Aggregate brothers gone from social event
permaAbsent = ["Brent Freed", "Jack Hendershott", "Donovan Frazer"]

# -------- Reese's addition for special events with more positions -----------
specialEvent = raw_input("Is this a special event? ")

if specialEvent.strip().lower() == "yes":
         runSpecialScheduler = True
else:
        runSpecialScheduler = False

# -------- Ryan's addition to make excluding absent brothers easier -----------

anyAbsent = raw_input("Will any brothers be absent from this event? ")
#absentMsg = "Will any brothers be absent from this event?"
#absentTtl = "Any Absent?"
#anyAbsent = eg.enterbox(absentMsg, absentTtl)
if anyAbsent.strip().lower() == "yes":
	runAbsentSurvey = True
        print("Please enter the names of each absent brother, one per line. Once" +
              "you have entered all the names, end with a new line containing a" +
              "single period. \n")
else:
	runAbsentSurvey = False

absentTonight = []

while runAbsentSurvey:
	name = raw_input()
	#msg = ("Please enter the name of the absent brother. "+
	#	"If this is the last absent brother, end with a period.")
	#ttl = "Absent Brothers"
	#name = eg.enterbox(msg, ttl)
        if name[0] == ".":
                last = True
	else:
		last = False
	if last:
		runAbsentSurvey = False
	else:
		absentTonight += [name]

print("Creating the schedule now.")

# ----------------------- End of Ryan's First Addition ------------------------

absent = permaAbsent + absentTonight

# Create a subset of brothers that can do work during social event
eligibleBros = list(set(epsilon + zeta + eta + theta + iota + pledges))

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
available_iota = [iota_mem for iota_mem in iota if not (iota_mem in absentTonight
        or iota_mem in permaAbsent or iota_mem in brothers_good_at_door)]
available_theta = [theta_mem for theta_mem in theta if not (theta_mem in absentTonight
        or theta_mem in permaAbsent or theta_mem in brothers_good_at_door)]
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
if runSpecialScheduler:
        doorShift4 = pickRandomBros(available_pledges, NUM_SLOTS)

# List of tuples. First item in tuple is the available members of the class.
# Second item is used as a weight so that lower class brothers (and pledges)
# get shifts more often.
if runSpecialScheduler: 
        lst_availables = [(available_epsilon, .3), (available_zeta, .7),
                          (available_eta, 1), (available_theta, 1),
                          (available_iota, 1), (available_pledges, 1)]

else:
        lst_availables = [(available_epsilon, .3), (available_zeta, .4),
                          (available_eta, .5), (available_theta, .6), 
                          (available_iota, .8),  (available_pledges, 1)]

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
if runSpecialScheduler:
        doorShift3 = pickRandomBros(final_lst, NUM_SLOTS)
        barShift3 = pickRandomBros(final_lst, NUM_SLOTS)
        barShift4 = pickRandomBros(final_lst, NUM_SLOTS)
# ---------------- Writes shift assignments to an excel file ------------------

filename =  "PhiPsiShifts.xlsx"
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
if runSpecialScheduler:
    worksheet.set_column(7, 7, 20)
    worksheet.set_column(8, 8, 20)
    worksheet.set_column(9, 9, 20)
    worksheet.set_column(10, 10, 20)

TITLE = '&C&30&"Calibri,Bold"Phi Psi Door and Bar Shift ' + DATE
worksheet.set_header(TITLE)

worksheet.write(3, 1, "10:00-10:30", time)
worksheet.write(4, 1, "10:30-11:00", time)
worksheet.write(5, 1, "11:00-11:30", time)
worksheet.write(6, 1, "11:30-12:00", time)
worksheet.write(7, 1, "12:00-12:30", time)
worksheet.write(8, 1, "12:30-1:00", time)
worksheet.write(9, 1, "1:00-1:30", time)

if runSpecialScheduler:
    worksheet.write(2, 2, "Back Door 1", header)
else:
    worksheet.write(2, 2, "Door 1", header)
row = 3
for bro in doorShift1:
	if row == 9:
		worksheet.write(row, 2, bro, names3)
	else:
		worksheet.write(row, 2, bro, names1)
	row += 1

if runSpecialScheduler:
    worksheet.write(2, 3, "Back Door 2", header)
else:
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

worksheet.write(2, 5, "Downstairs Bar 1", header)
row = 3
for bro in barShift1:
	if row == 9:
		worksheet.write(row, 5, bro, names3)
	else:
		worksheet.write(row, 5, bro, names1)
	row += 1

worksheet.write(2, 6, "Downstairs Bar 2", header)
row = 3
for bro in barShift2:
	if row == 9:
		worksheet.write(row, 6, bro, names4)
	else:
		worksheet.write(row, 6, bro, names2)
	row += 1

if runSpecialScheduler:
    worksheet.write(2, 7, "Upstairs Bar 1", header)
    row = 3
    for bro in barShift3:
        if row == 9:
                worksheet.write(row, 7, bro, names3)
        else:
                worksheet.write(row, 7, bro, names1)
        row += 1

    worksheet.write(2, 8, "Upstairs Bar 2", header)
    row = 3
    for bro in barShift4:
        if row == 9:
                worksheet.write(row, 8, bro, names3)
        else:
                worksheet.write(row, 8, bro, names1)
        row += 1

    worksheet.write(2, 9, "Courtyard 1", header)
    row = 3
    for bro in doorShift3:
        if row == 9:
                worksheet.write(row, 9, bro, names3)
        else:
                worksheet.write(row, 9, bro, names1)
        row += 1

    worksheet.write(2, 10, "Courtyard 2", header)
    row = 3
    for bro in doorShift4:
        if row == 9:
                worksheet.write(row, 10, bro, names4)
        else:
                worksheet.write(row, 10, bro, names2)
        row += 1

workbook.close()

os.system("open " + filename)

# ------------------------ end of Ryan's contribution -------------------------






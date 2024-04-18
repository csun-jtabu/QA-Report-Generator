# Jaztin Tabunda
# COMP 467 - Prof. Chaja
# 3-24-2024
# Project 2: The Reckoning

# Create a script that a user is able to parse and input data from a QA CSV into a database
#
# - Script will need to use Argparse
# - Database will be MongoDB (but can be any other DB if you prefer) Mongo is widely used
# in M&E for it's ease of flexible scheme, completely versatile non-relational DB
# - Will input your weekly QA reports into Collection 1
# - Will input a mock "DB dump" which will be everyone's reports in one mega file into
# Collection 2 (will also be a modified excel file)
# - Use DB to quickly create own reports

# -----------------------------------------------------------------------------------------------
# IMPORTANT NOTE:
# MY DEFINITION OF DUPLICATES: IF ALL FIELDS/COLUMNS IN ONE ROW/ENTRY MATCH ANOTHER ROW
# -----------------------------------------------------------------------------------------------
# libraries used
import pandas  # only using this to convert excel to csv
import openpyxl  # Pandas uses this library so we need this too.
import argparse
import csv
import pymongo
import re

# -----------------------------------------------------------------------------------------------
# Argparse setup
parser = argparse.ArgumentParser(description='Used to manipulate CSV data')

# This command/argument will hold the file names we want to reference
parser.add_argument('--files', dest='files',
                    help = 'we will store the files here', nargs='*')

# This command/argument will be used to convert Excel file to CSV
parser.add_argument('--toCSV', action='store_true', dest='csv',
                    help = 'Used to convert Excel file to CSV file')

# These commands will do the inserts to MongoDB
parser.add_argument('--add1', action='store_true', dest='add1',
                    help='add to Collection 1')
parser.add_argument('--add2', action='store_true', dest='add2',
                    help='add to Collection 2')

# 3a) List all work done by Your user (myself) - from both collections(No duplicates)
parser.add_argument('--findMyWork', action='store_true', dest='findMyWork',
                    help='This would find all the work done by me')
# 3b) All repeatable bugs - from both collections(No duplicates)
parser.add_argument('--findAllRepeat', action='store_true', dest='repeat',
                    help='Finds all repeatable bugs')
# 3c) All Blocker bugs - from both collections(No duplicates)
parser.add_argument('--findAllBlocker', action='store_true', dest='blocker',
                    help='Finds all blocker bugs')
# 3d) All reports on build 3/19/2024 - from both collections(No duplicates)
parser.add_argument('--findBuildXXXX', action='store_true', dest='findBuild',
                    help='Finds all reports on build xxxx')
# 3e) Report back the very 1st test case (Test #1), the middle test case
# (you determine that),and the final test case of your database - from collection 2
parser.add_argument('--FML', action='store_true', dest='fml',
                    help='Report back the 1st test case, the middle test case, '
                         'and the final test case of your database')

# 4. CSV export of user "Kevin Chaja" (use argparse to trigger code) - from collection 2
parser.add_argument('--findChajasWork', action='store_true', dest='findChajasWork',
                    help='This would find all the work done by Kevin Chaja')

# This is where all the arguments from the parser will be stored
args = parser.parse_args()

# -----------------------------------------------------------------------------------------------
# Mongodb setup
myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["project2db"]
col1 = mydb["Collection 1"]
col2 = mydb["Collection 2"]

# -----------------------------------------------------------------------------------------------
# Pre-Processing files
# Method to convert Excel file to CSV
def excelToCSV():
    # if we specify in the command line that we want to convert excel to csv.
    # Additionally, there must be a file name passed in
    if (args.csv == True) and (args.files != None):
        for file in args.files: # we check each file that was passed in
            excelChecker = re.search(".*\.xlsx$", file) # we check if it's an excel file
            if excelChecker != None: # if it is,
                excelFile = pandas.read_excel(file) # we use Pandas take the file in
                newFileName = str(file).replace('.xlsx', '.csv') # changing extension

                # ensures date format (m/d/y - No trailing zeros) in Build # field.
                # Not in date format = null/empty cell
                excelFile['Build #'] = pandas.to_datetime(excelFile['Build #'], errors='coerce')
                excelFile['Build #'] = excelFile['Build #'].dt.strftime('%#m/%d/%Y')

                excelFile.to_csv(str(newFileName), index=False, header=True) # we then convert it
            else: # if we pass in a different file type
                print('you didn\'t convert anything')
pass

# -----------------------------------------------------------------------------------------------
# Inserting into MongoDB
# This method is going to filter/remove invalid cells
def sanitizeDatabase(col):
    # makes a regular expression object for each column/field of our db
    validTestNumPattern = "^[0-9]+$" # Basically any number
    validBuildNumPattern = "^[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}$" # This Date format only: MM-DD-YYYY
    validCatTestExpectActualPattern = ".+" # Any string besides empty string
    validRepBlockPattern = "^yes$|^no$" # literally either "yes" or "no"
    validOwnerPattern = "^(?!([0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9])$|^$).*$"
    # Any string besides Dates and the empty cell. Uses "negative lookahead" for negation

    # We delete any row/entry that doesn't satisfy the valid regex for each respective field
    # "i" option ignores casing
    col.delete_many({"Test #": {"$not": {"$regex": validTestNumPattern}}})
    col.delete_many({"Build #": {"$not": {"$regex": validBuildNumPattern}}})
    col.delete_many({"Category": {"$not": {"$regex": validCatTestExpectActualPattern}}})
    col.delete_many({"Test Case": {"$not": {"$regex": validCatTestExpectActualPattern}}})
    col.delete_many({"Expected Result": {"$not": {"$regex": validCatTestExpectActualPattern}}})
    col.delete_many({"Actual Result": {"$not": {"$regex": validCatTestExpectActualPattern}}})
    col.delete_many({"Repeatable?": {"$not": {"$regex": validRepBlockPattern, "$options": "i"}}})
    col.delete_many({"Blocker?": {"$not": {"$regex": validRepBlockPattern, "$options": "i"}}})
    col.delete_many({"Test Owner": {"$not": {"$regex": validOwnerPattern}}})
pass

# This is going to add to collection 1
def inputToCollection1():
    global args, myclient, mydb, col1

    if args.add1 == True: # if --add1 is inputted in command line
        for element in args.files: # we check each file we pass in the command line
            with open(element, 'r', encoding='utf-8') as file: # opens a file
                csvReader = csv.reader(file) # pass the file into a csv reader
                for line in csvReader: # each line will be checked
                    myDict = {                 # each line's cells will be inputted
                        'Test #': line[0],     # into a dictionary
                        'Build #': line[1],
                        'Category': line[2],
                        'Test Case': line[3],
                        'Expected Result': line[4],
                        'Actual Result': line[5],
                        'Repeatable?': line[6],
                        'Blocker?': line[7],
                        'Test Owner': line[8]
                    }
                    # This basically checks if the row we are entering is already in the Collection
                    # i.e. duplicates
                    dupeCheck = col1.find_one(myDict)

                    if dupeCheck:  # If there is a duplicate
                        print('This is a duplicate. It\'s in Collection 1 already.')
                    else:  # If it isn't we insert it into the Collection
                        x = col1.insert_one(myDict)
    # we sanitize the database after inserting
    sanitizeDatabase(col1)
pass

# This is going to add to collection 2
def inputToCollection2():
    global args, myclient, mydb, col2

    if args.add2 == True: # if --add2 is inputted in command line
        for element in args.files: # we check each file we pass in the command line
            with open(element, 'r', encoding='utf-8') as file: # opens a file
                csvReader = csv.reader(file) # pass the file into a csv reader
                for line in csvReader: # each line will be checked
                    myDict = {                 # each line's cells will be inputted
                        'Test #':line[0],      # into a dictionary
                        'Build #': line[1],
                        'Category': line[2],
                        'Test Case': line[3],
                        'Expected Result': line[4],
                        'Actual Result': line[5],
                        'Repeatable?': line[6],
                        'Blocker?': line[7],
                        'Test Owner': line[8]
                    }
                    # This basically checks if the row we are entering is already in the Collection
                    # i.e. duplicates
                    dupeCheck = col2.find_one(myDict)

                    if dupeCheck: # If there is a duplicate
                        print('This is a duplicate. It\'s in Collection 2 already.')
                    else: # If it isn't we insert it into the Collection
                        x = col2.insert_one(myDict)
    # we sanitize the database after inserting
    sanitizeDatabase(col2)
pass

# -----------------------------------------------------------------------------------------------
# 3) Database Answers

# Query/Database Call Number 1:
# List all work [bugs] done by Your user [Jaztin Tabunda]
# from both collections(No duplicates)
def findMyWork():
    global args, myclient, mydb, col1, col2

    if args.findMyWork == True: # if --findMyWork is in commandline
        data = [] # this is where we'll store the rows/entries we are retrieving from query
        fieldNames = ['Test #', 'Build #','Category','Test Case','Expected Result',
                      'Actual Result','Repeatable?','Blocker?','Test Owner'] # used when we write to csv
        for element in col1.find({"Test Owner":"Jaztin Tabunda"}, {'_id':0}): # find in collection 1
            if element not in data: # ensuring no duplicates
                data.append(element) # we add the row to the list
        for element in col2.find({"Test Owner":"Jaztin Tabunda"}, {'_id':0}): # find in collection 2
            if element not in data: # ensuring no duplicates
                data.append(element) # we add the row to the list
        with open('DBAnswer1.csv', 'w', newline='', encoding='utf-8') as csvFile: # write to csv file
            csvWriter = csv.DictWriter(csvFile, fieldnames=fieldNames)
            csvWriter.writeheader()
            csvWriter.writerows(data)
pass

# Query/Database Call Number 2:
# All repeatable bugs- from both collections(No duplicates)
def findAllRepeatable():
    global args, myclient, mydb, col1, col2

    if args.repeat == True: # if --findAllRepeat is in commandline
        data = []
        fieldNames = ['Test #', 'Build #','Category','Test Case','Expected Result',
                      'Actual Result','Repeatable?','Blocker?','Test Owner']
        # find repeats in collection 1 no matter the casing
        for element in col1.find({"Repeatable?":{"$regex":"^yes$", "$options":"i"}}, {'_id': 0}):
            if element not in data: # ensuring no duplicates
                data.append(element) # we add the row to the list
        # find repeats in collection 2 no matter the casing
        for element in col2.find({"Repeatable?":{"$regex":"^yes$", "$options":"i"}}, {'_id': 0}):
            if element not in data: # ensuring no duplicates
                data.append(element) # we add the row to the list
        with open('DBAnswer2.csv', 'w', newline='', encoding='utf-8') as csvFile: # write to new csv file
            csvWriter = csv.DictWriter(csvFile, fieldnames=fieldNames)
            csvWriter.writeheader()
            csvWriter.writerows(data)
pass

# Query/Database Call Number 3:
# All Blocker bugs- from both collections(No duplicates)
def findAllBlocker():
    global args, myclient, mydb, col1, col2

    if args.blocker == True: # if --findAllBlocker is in commandline
        data = []
        fieldNames = ['Test #', 'Build #','Category','Test Case','Expected Result',
                      'Actual Result','Repeatable?','Blocker?','Test Owner']
        # find blockers in collection 1 no matter the casing
        for element in col1.find({"Blocker?":{"$regex":"^yes$", "$options":"i"}}, {'_id': 0}):
            if element not in data: # ensuring no duplicates
                data.append(element) # we add the row to the list
        # find blockers in collection 2 no matter the casing
        for element in col2.find({"Blocker?":{"$regex":"^yes$", "$options":"i"}}, {'_id': 0}):
            if element not in data: # ensuring no duplicates
                data.append(element) # we add the row to the list
        with open('DBAnswer3.csv', 'w', newline='', encoding='utf-8') as csvFile: # write to new csv file
            csvWriter = csv.DictWriter(csvFile, fieldnames=fieldNames)
            csvWriter.writeheader()
            csvWriter.writerows(data)
pass

# Query/Database Call Number 4:
# All reports on build 3/19/2024 - from both collections(No duplicates)
def findAll_3_19_2024():
    global args, myclient, mydb, col1, col2

    if args.findBuild == True: # if --findBuildXXXX is in commandline
        data = []
        fieldNames = ['Test #', 'Build #','Category','Test Case','Expected Result',
                      'Actual Result','Repeatable?','Blocker?','Test Owner']
        # find Build #s = 3/19/2024 in collection 1 no matter the casing
        for element in col1.find({"Build #":"3/19/2024"}, {'_id': 0}):
            if element not in data: # ensuring no duplicates
                data.append(element) # we add the row to the list
        # find Build #s = 3/19/2024 in collection 1 no matter the casing
        for element in col2.find({"Build #":"3/19/2024"}, {'_id': 0}):
            if element not in data: # ensuring no duplicates
                data.append(element) # we add the row to the list
        with open('DBAnswer4.csv', 'w', newline='', encoding='utf-8') as csvFile: # write to new csv file
            csvWriter = csv.DictWriter(csvFile, fieldnames=fieldNames)
            csvWriter.writeheader()
            csvWriter.writerows(data)
pass

# Query/Database Call Number 5:
# Report back the very 1st test case (Test #1),
# the middle test case (you determine that),
# and the final test case of your database - from collection 2
def findFML():
    global args, myclient, mydb, col2

    if args.fml == True:
        FML = []
        fieldNames = ["Placement", "_id", 'Test #', 'Build #','Category','Test Case','Expected Result',
                      'Actual Result','Repeatable?','Blocker?','Test Owner']

        # we get a mongdodb cursor object that references the collection 2 and is sorted by id in
        # ascending order
        documents = col2.find().sort([("_id", pymongo.ASCENDING)])
        docList = list(documents) # we convert to a list

        middle = int((len(docList)-1)/2) # index of middle element/entry/document
        last = (len(docList)-1) # index of last element/entry/document
        docList[0]["Placement"] = "First" # we add another key/value to first dictionary
        docList[middle]["Placement"] = "Middle" # we add another key/value to middle dictionary
        docList[last]["Placement"] = "Last" # we add another key/value to last dictionary
        FML.append(docList[0]) # we add the new dictionaries to a list
        FML.append(docList[middle])
        FML.append(docList[last])
        with open('DBAnswer5.csv', 'w', newline='', encoding='utf-8') as csvFile: # we write to a new csv
            csvWriter = csv.DictWriter(csvFile, fieldnames=fieldNames)
            csvWriter.writeheader()
            csvWriter.writerows(FML)
pass

# -----------------------------------------------------------------------------------------------
# 4. CSV export of user "Kevin Chaja" (use argparse to trigger code) - from collection 2
def findChajasWork():
    global args, myclient, mydb, col1, col2
    if args.findChajasWork == True: # if --findChajasWork is in commandline
        data = []
        fieldNames = ['Test #', 'Build #','Category','Test Case','Expected Result',
                      'Actual Result','Repeatable?','Blocker?','Test Owner']
        # find Kevin Chaja bugs in collection 2
        for element in col2.find({"Test Owner":"Kevin Chaja"}, {'_id':0}):
            if element not in data: # ensuring no duplicates
                data.append(element) # we add the row to the list
        with open('KevinChajaWork.csv', 'w', newline='', encoding='utf-8') as csvFile: # write to csv file
            csvWriter = csv.DictWriter(csvFile, fieldnames=fieldNames)
            csvWriter.writeheader()
            csvWriter.writerows(data)
pass

# -----------------------------------------------------------------------------------------------
def main():
    global args
    if args.files == []:
        print('No files selected')
    else:
        print("Files loaded")
    excelToCSV()
    inputToCollection1()
    inputToCollection2()
    findMyWork()
    findAllRepeatable()
    findAllBlocker()
    findAll_3_19_2024()
    findFML()
    findChajasWork()
pass

main()
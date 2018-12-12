import re
import xlrd
import xlwt
import os.path
import openpyxl
from datetime import datetime, timedelta
from time import sleep

def getWeekdayFromDate(date):
    day = datetime.strptime(date, '%Y/%m/%d').weekday()
    if day == 0:
        return "Monday"
    elif day == 1:
        return "Tuesday"
    elif day == 2:
        return "Wednesday"
    elif day == 3:
        return "Thursday"
    elif day == 4:
        return "Friday"
    elif day == 5:
        return "Saturday"
    else:
        return "Sunday"

def promptForMissingFields(data):
    if data[0] == "":
        data[0] = input("Could not find date, provide if known: ")
    if data[1] == "":
        data[1] = input("Could not find call number, provide if known ")
    if data[2] == "":
        data[2] = input("Could not find ticket number, provide if known: ")
    if data[5] == "":
        data[5] = input("Could not find postal code, provide if known: ")
    if data[6] == "":
        data[6] = input("Could not find name, provide if known: ")
    if data[7] == "":
        data[7] = input("Could not find 'ins field', provide if known (ATB/RB): ")
    if data[9] == "":
        data[9] = input("Could not find call description, provide if known (service.. etc): ")
    if data[10] == "":
        data[10] = input("Could not find km, please provide: ")
    if data[12] == "":
        data[12] = input("Could not find pay, please provide: ")
        data[12] = float(data[12])
    return data

# called after writing an excel entry, clears all the data variables
def clearFields(data):
    return [''] * 13

# must have the postcodes file in directory (or else change the first func line)
# produces a dictionary that maps postal codes to km and price
def readPostalCodes():    
    workbook = xlrd.open_workbook('postalCodes.xlsm')
    sheet = workbook.sheet_by_index(0)
    numCol = sheet.ncols
    numRows = sheet.nrows
    postalDictionary = {}

    for col in range(1, numRows):
        currentPostalCode = sheet.cell(col,3).value
        postalDictionary[currentPostalCode] = [sheet.cell(col,5).value, sheet.cell(col,6).value] #[km, price]        
    return postalDictionary

# Inserts a new row to the excel file with the data from recent call item
def addEntryToExcel(fileName, data):
    
    if not os.path.exists(fileName):
        wb = openpyxl.Workbook()
        wb.save(fileName)
    
    wb = openpyxl.load_workbook(fileName)
    sheet = wb['Sheet']

    sheet.append(data)
    wb.save(fileName)

# Used to add color/style to text (does not work in all OS system/terminals)
class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m' 


# initialize all call variables to something in case they are not found (and later have an access attempt on them)
date, callNum, ticketNum, merchantNum, ins, postalCode, name, time, desc, numTerm, address = "","","","","","","","","","",""
km, pay = 0, 0
callCounter = 0

# 
correctResponse = False
while not correctResponse:
    usageResponse = input('Produce invoice data (i) or just print calls to console (c)? ')
    if usageResponse.lower() == 'i':
        fileName = input("Provide name for excel file (without extension): ")
        fileName += ".xlsx" # openpyxl library will only work with xlsx, not xls
        correctResponse = True
    elif usageResponse.lower() == 'c':
        correctResponse = True
    else:
        print('Incorrect response...')
        sleep(1)

        
# get a dictionary that maps postal codes to km and pay, requires the postal codes file in directory (with my name change)
postalCodes = readPostalCodes()

# make list to store all the calls
calls = []

# GET ALL THE CALL DATA
with open('novFiles.txt') as txt:
    # for each line in email data
    for line in txt:
        
        if "Date :" in line: # get date & time
            date = line[9:19]
            emailTime = line[len(line)-9:].strip()
            try:
                day = getWeekdayFromDate(date)
            except:
                day = ""
        
        if 'Cust. Service' in line: # get call num
            callNum = line[26:].strip()
        
        if 'Service To' in line: # get name and merchant num
            name = line[14:len(line)-10].strip()
            merchantNum = line[len(line)-10:].strip()
        
        ticketLine = re.findall(r'[/]\w+\s+\d{8}', line) # get ticket num
        if len(ticketLine):
            ticketNum = ticketLine[0][10:]
            
        if 'Address' in line: # get postal code and using that, try to get km and pay
            postalCode = line[len(line)-7:].strip()
            address = line[13:].strip()
            if postalCode in postalCodes:
                km = int(postalCodes[postalCode][0])
                pay = '$' + str(postalCodes[postalCode][1])
                
        if '----PART.BOUNDARY.1--' in line: # markes end of call, save to data structure
            call = {}
            call['date'] = date
            call['name'] = name
            call['callNum'] = callNum
            call['ticketNum'] = ticketNum
            call['postalCode'] = postalCode
            call['day'] = day
            call['emailTime'] = emailTime
            call['km'] = km
            call['pay'] = pay
            call['merchantNum'] = merchantNum
            call['address'] = address
            calls.append(call) # calls is a list of dictionaries (and each dict is a call)
            date, callNum, ticketNum, postalCode, name = "", "", "", "", ""
            km, pay = 0, 0

calls.sort(key=lambda item:item['date']) # sort calls by date

for call in calls:
    if namefilter in call['name']:
        print('\n\nName: ' + bcolors.BOLD + bcolors.FAIL + call['name'] + bcolors.ENDC)
        print('Date: ' + call['date'][:8] + bcolors.WARNING + call['date'][8:] + bcolors.ENDC)
        print('Address:', call['address'])
        print('Call issued at:', call['emailTime'])
        print('Ticket #:', call['ticketNum'])
        print('Call #:', call['callNum'])
        print('km:', call['km'])
        if usageResponse.lower() == 'c':
            print('Weekday: ' + bcolors.WARNING + call['day'] + bcolors.ENDC)
        
        # make excel data if user requested this option
        if usageResponse.lower() == 'i':
            workedIt = input("(q to quit)\nDid you work this " + bcolors.WARNING + call['day'] + bcolors.ENDC + " call (1/0)? \n\n")
            if workedIt.lower() == 'q':
                break
            elif int(workedIt) == 1:
                time = input("Time on site (minutes)? ")
                time = str(timedelta(minutes=int(time))) # format the time string
                time = time[:len(time)-3]

                numTerm = input("Number of terminals? ")

                # Determine call type
                if call['callNum'] == "N/A": 
                    desc = input("No service call number found, provide call type/descrption: ")
                else:
                    desc = "Service"

                # Format km
                if int(call['km']) <= 10 and int(call['km']) > 0:
                    call['km'] = "0-10"
                elif int(km) <= 60:
                    call['km'] = "11-60"

                # Determine 'ins' field
                if len(call['merchantNum']) == 9:
                    if call['merchantNum'][0] == 'A': # if first character is A
                        ins = 'ATB'
                    else:
                        ins = 'RB'
                else:
                    ins = input("Could not determine 'ins' from merchant #, please provide (RB/ATB):")

                data = [call['date'], call['callNum'], call['ticketNum'], "", "", call['postalCode'] 
                        , call['name'], ins, time, desc, call['km'], numTerm, call['pay']]
                promptForMissingFields(data)
                data[12] = data[12]+'0' # dirty fix to add extra zero to pay field
                addEntryToExcel(fileName, data)

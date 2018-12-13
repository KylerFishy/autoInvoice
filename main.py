import re
import xlrd
import xlwt
import os.path
import openpyxl
from datetime import datetime, timedelta
from Invoice import *

command = getUserCommands()

# Create postalCodes to pay/km map, make list to store all the calls, and initialize a call dictionary
postalCodes = readPostalCodes()
calls = []
call = initializeCallObj()

# Collect all call data
with open('novFiles.txt') as txt:
    for line in txt: # for each line in email data..
                
        if lookForDate(line):
            call['date'], call['day'] = lookForDate(line)
            
        if lookForEmailTime(line):
            call['emailTime'] = lookForEmailTime(line)
            
        if lookForCallNum(line):
            call['callNum'] = lookForCallNum(line)
            
        if lookForNameAndMerchantNum(line):
            call['name'], call['merchantNum'] = lookForNameAndMerchantNum(line)

        if lookForTicket(line):
            call['ticketNum'] = lookForTicket(line)
                
        if lookForAddress(line, postalCodes):
            call['postalCode'], call['address'], call['km'], call['pay'] = lookForAddress(line, postalCodes)
                
        if '----PART.BOUNDARY.1--' in line: # markes end of call, save to data structure
            calls.append(call) # calls is a list of dictionaries (and each dict is a call)
            call = initializeCallObj()

calls.sort(key=lambda item:item['date']) # sort calls by date

for call in calls:
    printCallSummary(call, command)

    if command.lower() == 'i':
        continueToNextCall = excelEntryPrompt(call, fileName)
        if not continueToNextCall:
            break

from Invoice import *

# Create dict mapping postalCodes to km and pay, make list to store all the calls, 
# initialize a call dictionary, and get input from user (print calls or make invoice)
postalCodes = readPostalCodes()
calls = []
call = initializeCallObj()
command, fileName = getUserCommands()

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
            call = initializeCallObj() # reset call dictionary obj

calls.sort(key=lambda item:item['date']) # sort calls by date

for call in calls:
    printCallSummary(call, command) # command just to determine if weekday should be printed here

    if command.lower() == 'i': # if user chose to create (i)nvoice data
        continueToNextCall = excelEntryPrompt(call, fileName) # returns false if user quits
        if not continueToNextCall:
            print('Quitting...')
            break

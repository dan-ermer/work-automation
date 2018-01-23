#! python 3
#cetActivity.py

#libraries
import csv, codecs, datetime
types_of_encoding = ["utf8", "cp1252"]


#PARAMETERS
filePath = 'C:\\Users\\dermer\\PythonScripts\\Caesars\\'
processDate = "01-18"

datelabel = datetime.datetime.today().strftime("%Y%m%d")
print(datelabel)

cfp = filePath + processDate + "\\CET_WHG_20180118.txt"
ofp = filePath + processDate + "\\CET_STAYS_" + datelabel + "_100101.txt"
sfp = filePath + "ArrivalsDepartures.csv"
rfp = filePath + "CET_AwardNightRedemptionReport_012318.csv"

#Dictionaries/lists
sabreMemDict = {}
sabreConfDict = {}
redemptionDictIn = {}
cetArrIn = []


#Open Sabre Reservation Reference
for encoding_type in types_of_encoding:
    with codecs.open(sfp, encoding = encoding_type, errors ='replace') as sabreIn:
        #sabreIn = open(sfp, 'r')
        sabreReader = csv.reader(sabreIn)
        for srow in sabreReader:
            if len(srow) > 74:
                if srow[90]:
                    sabreMemDict[srow[90]] = srow[74]
                    sabreConfDict[srow[90]] = srow[25]
                    #print(str(sabreReader.line_num) + ' ' + str(srow[90]) + ' ' + str(srow[25]))
    sabreIn.close()

#To do: Open Olson Redemption Reference
redemptionIn = open(rfp, 'r')
redemptionReader = csv.reader(redemptionIn)
for rrow in redemptionReader: #16 & 17
    if len(rrow) == 16:
        redemptionDictIn[rrow[15]] = rrow[13]
        #print(str(redemptionReader.line_num) + ' ' + str(rrow[15]) + ' ' + str(rrow[13]))
    elif len(rrow) == 17:
        redemptionDictIn[rrow[16]] = rrow[14]
        #print(str(redemptionReader.line_num) + ' ' + str(rrow[16]) + ' ' + str(rrow[14]))
redemptionIn.close()

#Operating stats variables
countTotal = 0
countBarNoMem = 0
countBarMem = 0
countFRE = 0
countFST = 0

#Open Activty file from CET
cetIn = open(cfp, 'r')
cetReader = csv.reader(cetIn, delimiter='|')
headerRow = next(cetReader)
cetArrIn.append(headerRow)
for crow in cetReader:
    memNum = sabreMemDict.get(crow[12], '')
    confNum = sabreConfDict.get(crow[12], '')
    redemptionNum = redemptionDictIn.get(confNum, '')
    crow[17] = redemptionNum

    countTotal += 1
    
    #reformatting dates
    newArrDateArr = crow[8].split('-')
    newArrDate = str(newArrDateArr[1]) + '/' + str(newArrDateArr[2]) + '/' + str(newArrDateArr[0])
    crow[8] = newArrDate
    newDepDateArr = crow[9].split('-')
    newDepDate = str(newDepDateArr[1]) + '/' + str(newDepDateArr[2]) + '/' + str(newDepDateArr[0])
    crow[9] = newDepDate

    if memNum:
        crow[0] = memNum
        cetArrIn.append(crow)

        # compile stats
        if crow[15] == "WYBAR":
            countBarMem += 1
        elif crow[15] == "WYFRE":
            countFRE += 1
        elif crow[15] == "WYFST":
            countFST += 1
        #print(str(cetReader.line_num) + ' | ' + str(crow[0]) + ' | ' + str(crow[12]))
    else:
        countBarNoMem += 1
        
cetIn.close()

#To do: Open Olson Redemption Reference

#To do: Award Number Match

#write back out the updated file
outputFile = open(ofp, 'w', newline='')
outputWriter = csv.writer(outputFile, delimiter='|')
for row in cetArrIn:
    outputWriter.writerow(row)
outputFile.close()

#to do - add monitoring - stays without member numbers, redemptsions w/o member#

#print out summary stats
print("Total Stays: " + str(countTotal).rjust(24))
print("BAR Stays w/o Member Number: " + str(countBarNoMem).rjust(8))     
print("BAR Stays w Member Number: " + str(countBarMem).rjust(10))
print("GO FREE Stays w Member Number: " + str(countFRE).rjust(6))
print("GO FAST Stays w Member Number: " + str(countFST).rjust(6))
print("Processing Complete.")


#! python 3
#wvoInvoices.py

#libraries
import csv, codecs, datetime, shutil, os, openpyxl, pprint, pandas, xlsxwriter

#functions
def unicode_csv_reader(unicode_csv_data, dialect=csv.excel, **kwargs):
    # csv.py doesn't do Unicode; encode temporarily as UTF-8:
    csv_reader = csv.DictReader(utf_8_encoder(unicode_csv_data),
                            dialect=dialect, **kwargs)
    for row in csv_reader:
        # decode UTF-8 back to Unicode, cell by cell:
        yield [unicode(cell, 'utf-8') for cell in row]

def utf_8_encoder(unicode_csv_data):
    for line in unicode_csv_data:
        yield line.encode('utf-8')
        
#useful gobal values

types_of_encoding = ["utf8", "cp1252"]

datelabel = datetime.datetime.today().strftime("%m%d%y")
datetimelabel = datetime.datetime.today().strftime("%Y%m%d%H%M%S")
monthlabel = datetime.datetime.today().strftime("%b") + "-" + datetime.datetime.today().strftime("%y")

FP = 'C:\\Users\\dermer\\PythonScripts\\WVO - Katie\\'

#Get Latest Strategy
MasterFP = FP + 'WVO Redemptions Master.xlsx'
OlsonFP = FP + 'WEH_AwardNightRedemptionReport_' + datelabel + '.csv'
PropFP = FP + 'WVOPropertiesAndRates111517.csv' 
InvoiceFP = FP + 'Wyndham Rewards EH Interco Redemption Invoice - December 17.xlsx'

#1. Read in latest Olson WVO Redemptions
OlsonData = csv.DictReader(open(OlsonFP, 'r'))
OlsonIData = {} #re-write with award number as a key
for orow in OlsonData:  
    OlsonIData[int(orow['Award Number'])] = orow


#2. Read in Property Data
#to do: Go back and revisit opening XLS with instead of csv convert
PropData = csv.DictReader(open(PropFP, 'r'), delimiter='|')
PropIData = {} #re-write with Property Name as a key
for prow in PropData:
    PropIData[prow['Property Name 1'].upper()] = prow
         
#3. BackupWVO Master Redemptions File
BU = FP + 'Backup\\WVO Redemptions Master Backup ' + datetimelabel + '.xlsx'
shutil.copy(MasterFP, BU)
print('Backing up Redemption Master at: ' + BU)

#4. Read in current Master Redemption File
masterWB = openpyxl.load_workbook(MasterFP, data_only="TRUE")
masterWS = masterWB.active
MasterIData = {}
for row in masterWS.iter_rows(row_offset=1):

    billDate = row[26].value
    if isinstance(billDate,datetime.datetime):
        billDate = billDate.strftime("%b") + "-" + billDate.strftime("%y")

    MasterIData[row[15].value] = {'Unique Identifier':row[15].value,
                        'Member #':row[1].value,
                        'Member Level':row[2].value,
                        'Member Country of Residence':row[3].value,
                        'Redemption Date':row[4].value,
                        'Description':row[5].value,
                        'Site Name':row[6].value,
                        'Status':row[7].value,
                        'Arrival Date':row[8].value,
                        'Total Bedrooms':row[9].value,
                        'Total Nights':row[10].value,
                        'Points Per Award':row[11].value,
                        'Total Points Redeemed':row[12].value,
                        'PLUS Eligible':row[13].value,
                        'User ID':row[14].value,
                        'Award Number':row[15].value,
                        'Affiliation':row[16].value,
                        'Confirmation Number':row[17].value,
                        'Validation':row[18].value ,
                        'Comments':row[19].value ,
                        'iHotelier ID':row[20].value,
                        'SPE ID':row[21].value,
                        'Rate Schedule':row[22].value,
                        'Daily Reimbursement':row[23].value,
                        'Total Reimbursement':row[24].value,
                        'Redemption Processing Notes':row[25].value,
                        'Billed Date':billDate,
                        'Amount Billed':row[27].value,
                        'Invoice Processing Notes':row[28].value,
                        'Row Color':'white'}
masterWB.close()

#5. Scroll through Olson Redemptions and update master record
PrDt = datetime.date(2017, 1, 1)  #we don't need to process older records
SchDt = datetime.date(2017, 3, 5) #shift in reimbursement rate schedule
for orow in OlsonIData:
    r = OlsonIData[orow]
    #sts = ' ' + datelabel + ': '
    sts = ''
    rdta = r['Redemption Date'].split('/')
    rdt = datetime.date(int(rdta[2]), int(rdta[0]), int(rdta[1]))
    AwardType = "NONE"
    AwardMultiplier = 0
    AwardNightlyPoints = 0
    AwardRateKey = "NONE"
    AwardRedemptionRate = 0
    AwardSchedule = 2
    rowColor = "white"
    if rdt >= PrDt:
        #gather property data for the record
        if r['SIte Name'] in PropIData:
            #adding in site ids
            r['iHotelier ID'] = PropIData[r['SIte Name']]['iHotelierID']
            r['SPE ID'] = PropIData[r['SIte Name']]['SpeID']

            #determine award type, number of bedroomsn, nightly point cost
            # and total point cost of redemption, redemption rate as well
            #to do: python up this elif nonsense
            if "GO FAST" in r['Description'].upper():
                AwardType = "GO FAST"
                if any(x in r['Description'].upper() for x in ["1 BEDROOM","3000"]):
                    AwardMultiplier = 1
                    AwardNightlyPoints = 3000
                elif any(x in r['Description'].upper() for x in ["2 BEDROOM","6000"]):
                    AwardMultiplier = 2
                    AwardNightlyPoints = 6000
                elif any(x in r['Description'].upper() for x in ["3 BEDROOM","9000"]):
                    AwardMultiplier = 3
                    AwardNightlyPoints = 9000
                elif any(x in r['Description'].upper() for x in ["4 BEDROOM","12000"]):
                    AwardMultiplier = 4
                    AwardNightlyPoints = 12000
                else:
                    sts += "(3)No Room Count Match A! "
            elif "GO FREE" in r['Description'].upper():
                AwardType = "GO FREE"
                if any(x in r['Description'].upper() for x in ["1 BEDROOM","15000"]):
                    AwardMultiplier = 1
                    AwardNightlyPoints = 15000
                    AwardRateKey = "oneBedRoomv"
                elif any(x in r['Description'].upper() for x in ["2 BEDROOM","30000"]):
                    AwardMultiplier = 2
                    AwardNightlyPoints = 30000
                    AwardRateKey = "twoBedRoomv"
                elif any(x in r['Description'].upper() for x in ["3 BEDROOM","45000"]):
                    AwardMultiplier = 3
                    AwardNightlyPoints = 45000
                    AwardRateKey = "threeBedRoomv"
                elif any(x in r['Description'].upper() for x in ["4 BEDROOM","60000"]):
                    AwardMultiplier = 4
                    AwardNightlyPoints = 60000
                    AwardRateKey = "fourBedRoomv"
                else:
                    sts += "(3)No Room Count Match B! " + r['Description'].upper()

               #redmeption rate schedule retrieval
                if rdt >= SchDt:
                    if (AwardRateKey + '2') in PropIData[r['SIte Name']]:
                        #print('here:' + PropIData[r['SIte Name']]['Property Name 1'] + '|A')
                        AwardRedemptionRate = PropIData[r['SIte Name']][AwardRateKey + '2']
                    else:
                        sts += "(4)No Redemption Rate Schedule Match! "
                elif rdt < SchDt:
                    AwardSchedule = 1
                    if (AwardRateKey + '1') in PropIData[r['SIte Name']]:
                        #print('here:' + PropIData[r['SIte Name']][AwardRateKey + '1'] + '|B')
                        AwardRedemptionRate = PropIData[r['SIte Name']][AwardRateKey + '1']
                    else:
                        sts += "(4)No Redemption Rate Schedule Match! "

            else:
                sts += "(2)No Award Type Match! "
    
        else:
            sts += "(1)No Property Match! "
            r['iHotelier ID'] = None
            r['SPE ID'] = None

        r['Number of Rooms'] = AwardMultiplier
        r['Points Per Award'] = AwardNightlyPoints
        r['Number of Nights'] = int(r['Number of Nights'])
        r['Total Points Redeemed'] = (int(r['Number of Nights']) * int(r['Points Per Award']))
        r['Reimbursement Schedule'] = AwardSchedule
        r['Daily Reimbursement'] = AwardRedemptionRate
        #print(str(orow) + ' | ' + r['Number of Nights'] + ' | ' + str(AwardRedemptionRate) + ' | ' + sts)
        #print(str(orow) + ' | ' + AwardType + ' | ' + AwardRateKey + ' | ' + str(rdt) + ' | ' + str(SchDt) + ' | ' + PropIData[r['SIte Name']]['Property Name 1'] + ' | ' + sts)
        r['Total Reimbursement'] = (float(r['Number of Nights']) * float(AwardRedemptionRate))

        #now we start updating the master record
        if orow in MasterIData:
            #sts += "(6)Updating Existing Master Record | "
            if r['Status'] != MasterIData[orow]['Status']:
                sts += '(7) Status changed: ' +  MasterIData[orow]['Status'] + ' to ' + r['Status'] + ' | '
                rowColor = "#FFFF44"
        else:
            sts += "(8) Added | "
            rowColor = "#4444FF"
            MasterIData[orow] = {'Unique Identifier':None, 'Member #':None, 'Member Level':None, 'Member Country of Residence':None, 'Redemption Date':None, 'Description':None, 'Site Name':None, 'Status':None, 'Arrival Date':None, 'Total Bedrooms':None, 'Total Nights':None, 'Points Per Award':None, 'Total Points Redeemed':None, 'PLUS Eligible':None, 'User ID':None, 'Award Number':None, 'Affiliation':None, 'Confirmation Number':None, 'Validation':None , 'Comments':None , 'iHotelier ID':None, 'SPE ID':None, 'Rate Schedule':None, 'Daily Reimbursement':None, 'Total Reimbursement':None, 'Redemption Processing Notes':None, 'Billed Date':None, 'Amount Billed':None , 'Invoice Processing Notes':None, 'Row Color':None}
            
        #the following should not really change on existing records after we run the first time
        #to do: reevaluate, check out the get() function
        MasterIData[orow]['Unique Identifier'] = r['Award Number']
        MasterIData[orow]['Member #'] = r['Member #']
        MasterIData[orow]['Member Level'] = r['Member Level']
        MasterIData[orow]['Member Country of Residence'] = r['Member Country of Residence']
        MasterIData[orow]['Redemption Date'] = r['Redemption Date']
        MasterIData[orow]['Description'] = r['Description']
        MasterIData[orow]['Site Name'] = r['SIte Name']
        MasterIData[orow]['Status'] = r['Status']
        MasterIData[orow]['Arrival Date'] = r['Arrival Date']
        MasterIData[orow]['Total Bedrooms'] = r['Number of Rooms']
        MasterIData[orow]['Total Nights'] = r['Number of Nights']
        MasterIData[orow]['Points Per Award'] = r['Points Per Award']
        MasterIData[orow]['Total Points Redeemed'] = r['Total Points Redeemed']
        MasterIData[orow]['PLUS Eligible'] = r['PLUS Eligible']
        MasterIData[orow]['User ID'] = r['User ID']
        MasterIData[orow]['Award Number'] = r['Award Number']
        #MasterIData[orow]['Affiliation'] = r['Affiliation']
        MasterIData[orow]['Confirmation Number'] = r['Confirmation Number']
        #MasterIData[orow]['Validation'] = 
        #MasterIData[orow]['Comments'] = 
        MasterIData[orow]['iHotelier ID'] = r['iHotelier ID']
        MasterIData[orow]['SPE ID'] = r['SPE ID']
        MasterIData[orow]['Rate Schedule'] = r['Reimbursement Schedule']
        MasterIData[orow]['Daily Reimbursement'] = r['Daily Reimbursement']
        MasterIData[orow]['Total Reimbursement'] = r['Total Reimbursement']

        # ata some point we want to look into appending to existing notes - this will overwrite
        MasterIData[orow]['Redemption Processing Notes'] = sts
        #MasterIData[orow]['Billed Date'] = 
        #MasterIData[orow]['Amount Billed'] = 
        #MasterIData[orow]['Invoice Processing Notes'] = 
        MasterIData[orow]['Row Color'] = rowColor

#create a matching index for the master redemption record set
print("creating search index for master redemptions")
MasterISearch = {} #create a searchable index for the master redemption records
for srow in MasterIData:
    skey = str(MasterIData[srow]['Member #']) + str(MasterIData[srow]['iHotelier ID']) + str(MasterIData[srow]['Arrival Date']).replace("/","")
    MasterISearch [skey.upper()] = MasterIData[srow]['Award Number']


#Process invoice
#backup Invoice
BU2 = FP + 'Backup\\WVO Invoice Backup ' + datetimelabel + '.xlsx'
shutil.copy(InvoiceFP, BU2)
print('Backing up Invoice at: ' + BU2)

# 1. open invoice
iwb = openpyxl.load_workbook(InvoiceFP, data_only="TRUE")
iws = iwb.active
# 2. for each invoice record:
for row in iws.iter_rows(row_offset=9):
    rID = row[1].value
    InvoiceComment = ""
    if rID is not None:
        rSiteID = str(row[5].value)
        rReq = row[9].value
        rArrive = str(row[11].value)
        rConfirm = row[13].value
        rRes = row[15].value
        rMem = str(row[16].value)[-10:]
        lookup = rMem.upper() + rSiteID + rArrive[5:7] + rArrive[8:10] + rArrive[:4]
        #2a. Check for match in master redemption record
        if lookup in MasterISearch:

            mID = int(MasterISearch[lookup])
            
            row[20].value = MasterIData[mID]['Daily Reimbursement']
            
            row[22].value = "Yes" #On Olson Report

            #2b. If match compare reimbursement amounts
            mR = MasterIData[mID]['Total Reimbursement']
            row[21].value = mR #"Calculated Amount"
            
            if abs(mR - rReq) > 1:
                InvoiceComment = "Mismatch: " + str(mR) + '(WHG) - ' + str(rReq) + '(WVO) = ' + str(mR - rReq) + ' [' + MasterIData[mID]['Unique Identifier'] + ']' 
                #Update master record
                MasterIData[mID]['Row Color'] = '#FF4444'

                #Update invoice record
                row[19].value = "No"
                row[26].value = str(mR - rReq) #"Adjustment Amount"
                
            else:
                InvoiceComment = "Match: Move forward with payment" + ' [' + MasterIData[mID]['Unique Identifier'] + ']' 

                #Update master record
                MasterIData[mID]['Row Color'] = '#44FF44'
                MasterIData[mID]['Billed Date'] = monthlabel
                MasterIData[mID]['Amount Billed'] = rReq

                #Update invoice record
                row[19].value = "Yes"

            #2d. Update master redemption record with outcome
            MasterIData[mID]['Invoice Processing Notes'] = InvoiceComment
            
        else:
            row[19].value = "No"
            row[22].value = "No"
            InvoiceComment = "Cannot locate master redemption record: " + lookup

        #2c. Update invoice with outcome
        row[24].value = InvoiceComment

    
#re-write invoice XLS with processing outcomes
iwb.save(FP + "invoice review output " + str(datetimelabel) +  ".xlsx")
iwb.close()
        
print('Invoice has been reviewed and updated with outcome.')
            


#Write Redemption Master Data Back Out to Excel
workbook = xlsxwriter.Workbook(FP + 'WVO Redemptions Master.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1, 'bg_color': 'black', 'border':1, 'font_color':'white'})


counter = 0
row=0
col=0
#write header row
for k,v in (MasterIData.items()):
    for i in v.keys():
        worksheet.write(row, col, i, bold)
        col += 1
        counter += 1
    if counter > 0:
        break

#write records
row=1
col=0
for k,v in (MasterIData.items()):
    if 'Row Color' in v:
        clr = v['Row Color']
    else:
        clr = "white"

    cellf = workbook.add_format({'bg_color': clr, 'border':1})
       
    for i in v.values():
        worksheet.write(row, col, i, cellf)
        
        col += 1
        counter += 1
    row += 1
    col = 0

#Format Columns
worksheet.set_column('A:A', 16)
worksheet.set_column('B:B', 11)
worksheet.set_column('C:C', 13, None, {'hidden':1})
worksheet.set_column('D:D', 4, None, {'hidden':1})
worksheet.set_column('E:E', 16)
worksheet.set_column('F:F', 28)
worksheet.set_column('G:G', 35)
worksheet.set_column('H:H', 9)
worksheet.set_column('I:I', 11)
worksheet.set_column('J:J', 8)
worksheet.set_column('K:K', 8)
worksheet.set_column('L:L', 9)
worksheet.set_column('M:M', 9)
worksheet.set_column('N:N', 4)
worksheet.set_column('O:O', 8)
worksheet.set_column('P:P', 14)
worksheet.set_column('Q:Q', 25, None, {'hidden':1})
worksheet.set_column('R:R', 25)
worksheet.set_column('S:S', 9)
worksheet.set_column('T:T', 15)
worksheet.set_column('U:U', 10)
worksheet.set_column('V:V', 7)
worksheet.set_column('W:W', 4, None, {'hidden':1})
worksheet.set_column('X:X', 10)
worksheet.set_column('Y:Y', 10)
worksheet.set_column('Z:Z', 30)
worksheet.set_column('AA:AA', 10)
worksheet.set_column('AB:AB', 13)
worksheet.set_column('AC:AC', 50)
worksheet.set_column('AD:AD', 25, None, {'hidden':1})
workbook.close()

print('Master redemption List has been updated.')

   

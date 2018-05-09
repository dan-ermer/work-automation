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

FP = 'C:\\Users\\dermer\\AppData\\Local\\Programs\\Python\\WVO\\'

#Get Latest Strategy
MasterFP = FP + 'BlueThread Master Report.xlsx'
OlsonFP = FP + 'WEH_AwardNightRedemptionReport.csv'
PropFP = FP + 'Rates.csv' 
InvoiceFP = FP + 'Wyndham Rewards EH Interco Redemption Invoice - April.xlsx'

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

##    billDate = row[27].value
##    if isinstance(billDate,datetime.datetime):
##        billDate = billDate.strftime("%b") + "-" + billDate.strftime("%y")
    if row[15].value is not None:
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
                            'Award Number':int(row[15].value),
                            'Affiliation':row[16].value,
                            'Confirmation Number':row[17].value,
                            'Validation':row[18].value ,
                            'Comments':row[19].value ,
                            'iHotelier ID':int(row[20].value),
                            'SPE ID':int(row[21].value),
                            'Rate Schedule':row[22].value,
                            'Daily Reimbursement':float(row[23].value),
                            'Total Reimbursement':float(row[24].value),
                            'Redemption Processing Notes':row[25].value,
                            'Invoice Processing Notes':row[26].value,
                            'Billed Date':row[27].value,
                            'Amount Billed':row[28].value,
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
    adta = r['Arrival Date'].split('/')
    if len(adta) == 3:
        adt = datetime.date(int(adta[2]), int(adta[0]), int(adta[1]))
    else:
        adt = None
    
    AwardType = "NONE"
    AwardMultiplier = 0
    AwardNightlyPoints = 0
    AwardRateKey = "NONE"
    AwardRedemptionRate = 0
    AwardSchedule = 2
    rowColor = "white"
    if rdt >= PrDt:
        #Determine Award Type
        if "GO FAST" in r['Description'].upper():
            AwardType = "GO FAST"
            if any(x in r['Description'].upper() for x in ["1 BEDROOM","3000"]):
                AwardMultiplier = 1
                AwardNightlyPoints = 3000
            elif any(x in r['Description'].upper() for x in ["2 BEDROOM","6000"]):
                AwardMultiplier = 2
                AwardNightlyPoints = 6000
            else:
                sts += "(3)No Room Count Match A! "
        elif "GO FREE" in r['Description'].upper():
            AwardType = "GO FREE"
            if any(x in r['Description'].upper() for x in ["1 BEDROOM","15000"]):
                AwardMultiplier = 1
                AwardNightlyPoints = 15000
            elif any(x in r['Description'].upper() for x in ["2 BEDROOM","30000"]):
                AwardMultiplier = 2
                AwardNightlyPoints = 30000
            else:
                sts += "(3)No Room Count Match B! " + r['Description'].upper()
        else:
            sts += "(2)No Award Type Match! " + r['Description'].upper()


        #gather property data for the record
        snu = r['SIte Name'].upper()
        if snu in PropIData:
            #adding in site ids
            r['iHotelier ID'] = PropIData[snu]['iHotelier ID']
            r['SPE ID'] = PropIData[snu]['SPE ID']

            #determine redemption rate as well
            if (AwardType == "GO FREE") and  (AwardMultiplier == 1):
                AwardRedemptionRate = PropIData[snu]['1BRRate']
            elif (AwardType == "GO FREE") and  (AwardMultiplier == 2):
                AwardRedemptionRate = PropIData[snu]['2BRRate']
        else:
            sts += "(1)No Property Match! " + r['SIte Name']
            r['iHotelier ID'] = 0
            r['SPE ID'] = 0

        r['Number of Rooms'] = AwardMultiplier
        r['Points Per Award'] = AwardNightlyPoints
        r['Number of Nights'] = int(r['Number of Nights'])
        r['Total Points Redeemed'] = (int(r['Number of Nights']) * int(r['Points Per Award']))
        r['Reimbursement Schedule'] = AwardSchedule
        r['Daily Reimbursement'] = AwardRedemptionRate
        r['Total Reimbursement'] = (float(r['Number of Nights']) * float(AwardRedemptionRate))

        #now we start updating the master record
        if orow in MasterIData:
            #sts += "(6)Updating Existing Master Record | "
            if r['Status'] != MasterIData[orow]['Status']:
                sts += '(7) Status changed: ' +  MasterIData[orow]['Status'] + ' to ' + r['Status'] + ' | '
                rowColor = "#FFFF44"
        else:
            #new record
            sts += "(8) Added | "
            rowColor = "#4444FF"
            MasterIData[orow] = {'Unique Identifier':None, 'Member #':None, 'Member Level':None, 'Member Country of Residence':None, 'Redemption Date':None, 'Description':None, 'Site Name':None, 'Status':None, 'Arrival Date':None, 'Total Bedrooms':None, 'Total Nights':None, 'Points Per Award':None, 'Total Points Redeemed':None, 'PLUS Eligible':None, 'User ID':None, 'Award Number':None, 'Affiliation':None, 'Confirmation Number':None, 'Validation':None , 'Comments':None , 'iHotelier ID':None, 'SPE ID':None, 'Rate Schedule':None, 'Daily Reimbursement':None, 'Total Reimbursement':None, 'Redemption Processing Notes':None, 'Invoice Processing Notes':None, 'Billed Date':None, 'Amount Billed':None , 'Row Color':None}
            #to do: reevaluate, check out the get() function
            MasterIData[orow]['Unique Identifier'] = int(r['Award Number'])
            MasterIData[orow]['Member #'] = r['Member #']
            MasterIData[orow]['Member Level'] = r['Member Level']
            MasterIData[orow]['Member Country of Residence'] = r['Member Country of Residence']
            MasterIData[orow]['Redemption Date'] = rdt
            MasterIData[orow]['Description'] = r['Description']
            MasterIData[orow]['Site Name'] = r['SIte Name']
            MasterIData[orow]['Arrival Date'] = adt
            MasterIData[orow]['Total Bedrooms'] = r['Number of Rooms']
            MasterIData[orow]['Total Nights'] = r['Number of Nights']
            MasterIData[orow]['Points Per Award'] = r['Points Per Award']
            MasterIData[orow]['Total Points Redeemed'] = r['Total Points Redeemed']
            MasterIData[orow]['PLUS Eligible'] = r['PLUS Eligible']
            MasterIData[orow]['User ID'] = r['User ID']
            MasterIData[orow]['Award Number'] = int(r['Award Number'])
            MasterIData[orow]['Confirmation Number'] = r['Confirmation Number']
            MasterIData[orow]['iHotelier ID'] = int(r['iHotelier ID'])
            MasterIData[orow]['SPE ID'] = int(r['SPE ID'])
            MasterIData[orow]['Rate Schedule'] = r['Reimbursement Schedule']
            MasterIData[orow]['Daily Reimbursement'] = float(r['Daily Reimbursement'])
            MasterIData[orow]['Total Reimbursement'] = float(r['Total Reimbursement'])


        MasterIData[orow]['Status'] = r['Status']
        MasterIData[orow]['Redemption Processing Notes'] = sts
        MasterIData[orow]['Row Color'] = rowColor

#create a matching index for the master redemption record set
print("creating search index for master redemptions")
MasterISearch = {} #create a searchable index for the master redemption records
ctr = 0
for srow in MasterIData:
    #need a date tag - existing records are already dates, new records are just date like strings at this point and need to be handled differently
    #to do: maybe move up the date forcing at somepoint :-)
    if isinstance(MasterIData[srow]['Arrival Date'],datetime.datetime):
        madt = MasterIData[srow]['Arrival Date'].strftime("%m") + MasterIData[srow]['Arrival Date'].strftime("%d") + MasterIData[srow]['Arrival Date'].strftime("%Y")
    else:
        madt = str(MasterIData[srow]['Arrival Date']).replace("-","")
        
    skey = str(MasterIData[srow]['Member #']) + str(MasterIData[srow]['iHotelier ID']) + madt
    #keys may map to multiple redemption records so adding all matches to an array
    if not skey.upper() in MasterISearch:
        MasterISearch[skey.upper()] = [MasterIData[srow]['Award Number']]
    else:
        MasterISearch[skey.upper()].append(MasterIData[srow]['Award Number'])


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
    mID = 0
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
            #to do: See if there are multiple redemption records and deal
            if len(MasterISearch[lookup]) > 1:
                for i in MasterISearch[lookup]:
                   # print('i: ' + str(i))
                    if MasterIData[i]['Invoice Processing Notes'] is None:
                        mID = int(i)
                        break; 
            else: 
                mID = int(MasterISearch[lookup][0])

            if (mID == 0) or (MasterIData[mID]['Invoice Processing Notes'] is not None):
                InvoiceComment = "All matching redemption records have already been processed " + str(lookup)
            elif (MasterIData[mID]['Invoice Processing Notes'] == "Cancelled"):
                 InvoiceComment = "Matching redemption record is cancelled " + str(lookup)   
            else:
                #we have a matching redemption record, move forward with comparison
                row[24].value = MasterIData[mID]['Daily Reimbursement']
                row[26].value = "Yes" #On Olson Report

                #2b. If Compare reimbursement amounts
                mR = MasterIData[mID]['Total Reimbursement']
                row[25].value = mR #"Calculated Amount"
            
                if abs(mR - rReq) > 1:
                    InvoiceComment = "Mismatch: " + str(mR) + '(WHG) - ' + str(rReq) + '(WVO) = ' + str(mR - rReq) + ' [' + str(MasterIData[mID]['Unique Identifier']) + ']' 
                    #Update master record
                    MasterIData[mID]['Row Color'] = '#FF4444'

                    #Update invoice record
                    row[23].value = "No"
                    row[30].value = str(mR - rReq) #"Adjustment Amount"
                    
                else:
                    InvoiceComment = "Match: Move forward with payment" + ' [' + str(MasterIData[mID]['Unique Identifier']) + ']' 

                    #Update master record
                    MasterIData[mID]['Row Color'] = '#44FF44'
##                    MasterIData[mID]['Billed Date'] = monthlabel
##                    MasterIData[mID]['Amount Billed'] = rReq

                    #Update invoice record
                    row[23].value = "Yes"

            #2d. Update master redemption record with outcome
            #print('mID: ' + str(mID) + '  Notes: ' + InvoiceComment)
            if mID in MasterIData:
                MasterIData[mID]['Invoice Processing Notes'] = InvoiceComment
            
        else:
            row[23].value = "No"
            row[26].value = "No"
            InvoiceComment = "Cannot locate master redemption record: " + lookup

        #2c. Update invoice with outcome
        row[28].value = InvoiceComment

    
#re-write invoice XLS with processing outcomes
iwb.save(FP + "invoice review output " + str(datetimelabel) +  ".xlsx")
iwb.close()
        
print('Invoice has been reviewed and updated with outcome.')
            


#Write Redemption Master Data Back Out to Excel
workbook = xlsxwriter.Workbook(FP + 'WVO Redemptions Master ' + datetimelabel + '.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1, 'bg_color': 'black', 'border':1, 'font_color':'white'})
dfmt = workbook.add_format({'num_format':'mm/dd/yyyy'})

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
        if (col == 4) or (col == 8):
            cellf = workbook.add_format({'bg_color': clr, 'border':1, 'num_format':'mm/dd/yyyy'})
        elif (col == 27):
            cellf = workbook.add_format({'bg_color': clr, 'border':1, 'num_format':'yy-mmm'})
        else:
            cellf = workbook.add_format({'bg_color': clr, 'border':1})
            
        worksheet.write(row, col, i, cellf)    
        col += 1
        counter += 1
        cellf = None
    row += 1
    col = 0

#Format Columns
worksheet.set_column('A:A', 16)
worksheet.set_column('B:B', 11)
worksheet.set_column('C:C', 13, None, {'hidden':1})
worksheet.set_column('D:D', 4, None, {'hidden':1})
worksheet.set_column('E:E', 16, None)
worksheet.set_column('F:F', 28)
worksheet.set_column('G:G', 35)
worksheet.set_column('H:H', 9)
worksheet.set_column('I:I', 11, None)
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
worksheet.set_column('AA:AA', 30)
worksheet.set_column('AB:AB', 13)
worksheet.set_column('AC:AC', 13)
worksheet.set_column('AD:AD', 25, None, {'hidden':1})
workbook.close()

print('Master Redemption List has been updated.')

   

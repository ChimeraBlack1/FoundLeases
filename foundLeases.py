import math
import xlrd
import xlwt

def Device_Speed(speed):
  device_speed = {
    1: "Low",
    2: "Medium",
    3: "High",
    4: "Unknown"
  }
  return device_speed.get(speed)

def last_day(monthName):
  last_day = {
    "January": "1-31-2019",
    "February": "2-28-2019",
    "March": "3-31-2019",
    "April": "4-30-2019",
    "May": "5-31-2019",
    "June": "6-30-2019",
    "July": "7-31-2019",
    "August": "8-31-2019",
    "September": "9-30-2019",
    "October": "10-31-2019",
    "November": "11-30-2019",
    "December": "12-31-2019",
  }
  return last_day.get(monthName)

# PREVIOUS Month's Report
goodFile = False

while goodFile == False:
  fileToRead = input("Please enter the name of the PREVIOUS month's report)> ")
  if fileToRead == "exit" or fileToRead == "quit":
    print("ok, bye!")
    exit()
  else:
    prevReport = fileToRead + ".xlsm"
    try:
      prevwb = xlrd.open_workbook(prevReport)
      prevSheet = prevwb.sheet_by_index(0)
      goodFile = True
    except:
      print("I can't find that file, try again...")


# CURRENT month's report
goodReport = False

while goodReport == False:
  fileToRead2 = input("Please enter the name of the THIS month's report)> ")
  if fileToRead2 == "exit" or fileToRead2 == "quit":
    print("ok, bye!")
    exit()
  else:
    currentReport = fileToRead2 + ".xlsm"
    try:
      currentwb = xlrd.open_workbook(currentReport)
      currentSheet = currentwb.sheet_by_index(0)
      goodReport = True
    except:
      print("I can't find that file, try again...")


# PREVIOUS Sheet row count
PrevMonthSheetNumberGood = False

while PrevMonthSheetNumberGood == False:
  endOfPrevMonthSheetInput = input("How many cells are in the PREVIOUS month's sheet)> ")
  if endOfPrevMonthSheetInput == "exit" or endOfPrevMonthSheetInput == "quit":
    print("ok, bye!")
    exit()
  else:
    try:
      endOfPrevMonthSheet = int(endOfPrevMonthSheetInput)
      PrevMonthSheetNumberGood = True
    except:
      print("Not the right number, try again...)>")


# CURRENT sheet row count
ThisMonthSheetNumberGood = False

while ThisMonthSheetNumberGood == False:
  endOfThisMonthSheetInput = input("How many cells are in the THIS month's sheet)> ")
  if endOfThisMonthSheetInput == "exit" or endOfThisMonthSheetInput == "quit":
    print("ok, bye!")
    exit()
  else:
    try:
      endOfThisMonthSheet = int(endOfThisMonthSheetInput)
      ThisMonthSheetNumberGood = True
    except:
      print("Not the right number, try again...)>")

# Which Month
goodMonth = False
thisMonth = "default"

while goodMonth == False:
  whatMonth = input("What Month is this?)> ")
  if whatMonth == "exit" or whatMonth == "quit":
    print("ok, bye!")
    exit()
  else:
    try:
      thisMonth = whatMonth
      goodMonth = True
    except:
      print("Sorry that's not a month...)>")


#write to workbook
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('New Leases Found')
NewWorkbookName = "NewLeasesFound.xls"

#Titles
worksheet.write(0, 0, "Serial Number")
worksheet.write(0, 1, "Asset Price")
worksheet.write(0, 2, "Customer Name")
worksheet.write(0, 3, "Model Name")
worksheet.write(0, 4, "Address")
worksheet.write(0, 5, "Funder")
worksheet.write(0, 6, "Device Speed")
worksheet.write(0, 7, "Device Speed Name")
worksheet.write(0, 8, "Month")
worksheet.write(0, 9, "Month Name")

#Globals
# endOfPrevMonthSheet = 3569
# endOfThisMonthSheet = 3458

#Sherpa
SherpaReportAssetVolType = 8
SherpaReportSerialCol = 10
SherpaReportAssetPriceCol = 11
SherpaReportCustomerNameCol = 2
SherpaCustomerAddyCol = 12
SherpaReportCustomerModelCol = 9
SherpaReportTestFunderCol = 6
newWorkbookPointer = 1

#LOST
for x in range(1, endOfThisMonthSheet):
  found = False
  # get serial to test from Sherpa Report
  try:
    testSerial = currentSheet.cell_value(x,SherpaReportSerialCol)
    testAssetPrice = currentSheet.cell_value(x,SherpaReportAssetPriceCol)
    testCustomerName = currentSheet.cell_value(x,SherpaReportCustomerNameCol)
    testCustomerAddy = currentSheet.cell_value(x,SherpaCustomerAddyCol)
    testCustomerModel = currentSheet.cell_value(x,SherpaReportCustomerModelCol)
    testFunder = currentSheet.cell_value(x,SherpaReportTestFunderCol)
    testAssetVol = currentSheet.cell_value(x,SherpaReportAssetVolType)
    testSerial = int(testSerial)
  except:
    testSerial = str(testSerial)

  # look in the other report for the serial
  for y in range(1, endOfPrevMonthSheet):
    try:
      otherTestSerial = prevSheet.cell_value(y,SherpaReportSerialCol)
    except:
      continue

    if testSerial == "":
      worksheet.write(newWorkbookPointer, 0,"Blank Serial")
      worksheet.write(newWorkbookPointer, 1, testAssetPrice)
      worksheet.write(newWorkbookPointer, 2, testCustomerName)
      worksheet.write(newWorkbookPointer, 3, testCustomerModel)
      worksheet.write(newWorkbookPointer, 4, testCustomerAddy)
      worksheet.write(newWorkbookPointer, 5, testFunder)
      worksheet.write(newWorkbookPointer, 6, testAssetVol)
      newWorkbookPointer = newWorkbookPointer + 1
      continue
    if testSerial == otherTestSerial:
      found = True
      break
 
  #if found, go to next item
  if found == True:
    continue
  
  if found == False:
    worksheet.write(newWorkbookPointer, 0, testSerial)
    worksheet.write(newWorkbookPointer, 1, testAssetPrice)
    worksheet.write(newWorkbookPointer, 2, testCustomerName)
    worksheet.write(newWorkbookPointer, 3, testCustomerModel)
    worksheet.write(newWorkbookPointer, 4, testCustomerAddy)
    worksheet.write(newWorkbookPointer, 5, testFunder)
    worksheet.write(newWorkbookPointer, 6, testAssetVol)
    worksheet.write(newWorkbookPointer, 7, Device_Speed(testAssetVol))
    worksheet.write(newWorkbookPointer, 8, last_day(thisMonth))
    worksheet.write(newWorkbookPointer, 9, thisMonth)
    newWorkbookPointer = newWorkbookPointer + 1 

workbook.save(NewWorkbookName)
print("saved: " + str(NewWorkbookName))
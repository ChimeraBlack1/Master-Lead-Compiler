import excelCompare as ec

masterSheet = ec.OpenSheet("Leads_Tracking_Master.xlsx")
mgrSheet = ec.OpenSheet("02182020Rons_Leads.xlsx")

mastCol = 1
mastRow = 1
mgrCol = 1
mgrRow = 1

endOfMaster = ec.FindLastRow(masterSheet)
endOfMgr = ec.FindLastRow(mgrSheet)

masterDict = {}
mgrDict = {}
missing = []
missingDict = {}
missingCount = 0
mastLeadCount = 0

# value columns
assignedmgr_Col = 9
assignedrep_Col = 10
leadStatus_Col = 11
leadWon_Col = 12
reasonForLoss_Col = 13
clientQuote_Col = 14
requestedProduct_Col = 15
updated_Col = 16
comments_Col = 17
date_Col = 0
leadID_col = 1
source_Col = 2
method_Col = 3
interest_Col = 4
company_Col = 5 
contact_Col = 6
phone_Col = 7
email_Col = 8

for i in range(mastRow, endOfMaster):
  mastLead = ec.GetValue(masterSheet, i, mastCol)
  masterleadValues = {
    "xlRow": i,
    "assignedmgr_": ec.GetValue(masterSheet, i, assignedmgr_Col),
    "assignedrep_": ec.GetValue(masterSheet, i, assignedrep_Col),
    "leadStatus_": ec.GetValue(masterSheet, i, leadStatus_Col),
    "leadWon_": ec.GetValue(masterSheet, i, leadWon_Col),
    "reasonForLoss_": ec.GetValue(masterSheet, i, reasonForLoss_Col),
    "clientQuote_": ec.GetValue(masterSheet, i, clientQuote_Col),
    "requestedProduct_": ec.GetValue(masterSheet, i, requestedProduct_Col),
    "updated_": ec.GetValue(masterSheet, i, updated_Col),
    "comments_": ec.GetValue(masterSheet, i, comments_Col),
  }
  masterDict[mastLead] = masterleadValues
  mastLeadCount += 1

for i in range(mgrRow, endOfMgr):
  mgrLead = ec.GetValue(mgrSheet, i, mgrCol)
  if mgrLead in masterDict: 
    # get lead values
    leadValues = {
      "xlRow": masterDict[mgrLead]['xlRow'],
      "assignedmgr_": ec.GetValue(mgrSheet, i, assignedmgr_Col),
      "assignedrep_": ec.GetValue(mgrSheet, i, assignedrep_Col),
      "leadStatus_": ec.GetValue(mgrSheet, i, leadStatus_Col),
      "leadWon_": ec.GetValue(mgrSheet, i, leadWon_Col),
      "reasonForLoss_": ec.GetValue(mgrSheet, i, reasonForLoss_Col),
      "clientQuote_": ec.GetValue(mgrSheet, i, clientQuote_Col),
      "requestedProduct_": ec.GetValue(mgrSheet, i, requestedProduct_Col),
      "updated_": ec.GetValue(mgrSheet, i, updated_Col),
      "comments_": ec.GetValue(mgrSheet, i, comments_Col),
    }
    mgrDict[mgrLead] = leadValues
  else:
    missing.append(mgrLead)
    missingCount += 1
    missingValues = {
      "date_": ec.GetValue(mgrSheet, i, date_Col),
      "source_": ec.GetValue(mgrSheet, i, source_Col),
      "method_": ec.GetValue(mgrSheet, i, method_Col),
      "interest_": ec.GetValue(mgrSheet, i, interest_Col),
      "company_": ec.GetValue(mgrSheet, i, company_Col),
      "contact_": ec.GetValue(mgrSheet, i, contact_Col),
      "phone_": ec.GetValue(mgrSheet, i, phone_Col),
      "email_": ec.GetValue(mgrSheet, i, email_Col),
      "assignedmgr_": ec.GetValue(mgrSheet, i, assignedmgr_Col),
      "assignedrep_": ec.GetValue(mgrSheet, i, assignedrep_Col),
      "leadStatus_": ec.GetValue(mgrSheet, i, leadStatus_Col),
      "leadWon_": ec.GetValue(mgrSheet, i, leadWon_Col),
      "reasonForLoss_": ec.GetValue(mgrSheet, i, reasonForLoss_Col),
      "clientQuote_": ec.GetValue(mgrSheet, i, clientQuote_Col),
      "requestedProduct_": ec.GetValue(mgrSheet, i, requestedProduct_Col),
      "updated_": ec.GetValue(mgrSheet, i, updated_Col),
      "comments_": ec.GetValue(mgrSheet, i, comments_Col),
    }
    missingDict[mgrLead] = missingValues

#write to new xls wb
newWb = ec.Newb()
newSheet = ec.News(newWb, "Notes")

# write MASTER values to new wb
for i in masterDict:
  row = masterDict[i]['xlRow']
  if masterDict[i]['assignedmgr_'] != "":
    ec.SetValue(newSheet, masterDict[i]['assignedmgr_'], row, assignedmgr_Col)
    #write in serial number to check
    ec.SetValue(newSheet, i, row, comments_Col +1)
    ec.SetValue(newSheet, "From Master", row, comments_Col +3)
  if masterDict[i]['assignedrep_'] != "":
    ec.SetValue(newSheet, masterDict[i]['assignedrep_'], row, assignedrep_Col)
  if masterDict[i]['leadStatus_'] != "":
    ec.SetValue(newSheet, masterDict[i]['leadStatus_'], row, leadStatus_Col)
  if masterDict[i]['leadWon_'] != "":
    ec.SetValue(newSheet, masterDict[i]['leadWon_'], row, leadWon_Col)
  if masterDict[i]['reasonForLoss_'] != "":
    ec.SetValue(newSheet, masterDict[i]['reasonForLoss_'], row, reasonForLoss_Col)
  if masterDict[i]['clientQuote_'] != "":
    ec.SetValue(newSheet, masterDict[i]['clientQuote_'], row, clientQuote_Col)
  if masterDict[i]['requestedProduct_'] != "":
    ec.SetValue(newSheet, masterDict[i]['requestedProduct_'], row, requestedProduct_Col)
  if masterDict[i]['updated_'] != "":
    ec.SetValue(newSheet, masterDict[i]['updated_'], row, updated_Col)
  if masterDict[i]['comments_'] != "":
    ec.SetValue(newSheet, masterDict[i]['comments_'], row, comments_Col)


# write MGR values to new wb
for i in mgrDict:
  row = mgrDict[i]['xlRow']
  if mgrDict[i]['assignedmgr_']:
    ec.SetValue(newSheet, mgrDict[i]['assignedmgr_'], row, assignedmgr_Col)
    #write in serial number to check
    ec.SetValue(newSheet, i, row, comments_Col +1)
    ec.SetValue(newSheet, "From Manager", row, comments_Col +2)
  if mgrDict[i]['assignedrep_']:
    ec.SetValue(newSheet, mgrDict[i]['assignedrep_'], row, assignedrep_Col)
  if mgrDict[i]['leadStatus_']:
    ec.SetValue(newSheet, mgrDict[i]['leadStatus_'], row, leadStatus_Col)
  if mgrDict[i]['leadWon_']:
    ec.SetValue(newSheet, mgrDict[i]['leadWon_'], row, leadWon_Col)
  if mgrDict[i]['reasonForLoss_']:
    ec.SetValue(newSheet, mgrDict[i]['reasonForLoss_'], row, reasonForLoss_Col)
  if mgrDict[i]['clientQuote_']:
    ec.SetValue(newSheet, mgrDict[i]['clientQuote_'], row, clientQuote_Col)
  if mgrDict[i]['requestedProduct_']:
    ec.SetValue(newSheet, mgrDict[i]['requestedProduct_'], row, requestedProduct_Col)
  if mgrDict[i]['updated_']:
    ec.SetValue(newSheet, mgrDict[i]['updated_'], row, updated_Col)
  if mgrDict[i]['comments_']:
    ec.SetValue(newSheet, mgrDict[i]['comments_'], row, comments_Col)


# handle missing units
if missingCount > 0:
  print("missing: " + str(missingCount) + " units")
  print("They are: ")
  row = endOfMaster
  for i in missingDict:
    row += 1
    ec.SetValue(newSheet, missingDict[i]['date_'], row, date_Col)
    ec.SetValue(newSheet, i, row, leadID_col)
    ec.SetValue(newSheet, missingDict[i]['source_'], row, source_Col)
    ec.SetValue(newSheet, missingDict[i]['method_'], row, method_Col)
    ec.SetValue(newSheet, missingDict[i]['interest_'], row, interest_Col)
    ec.SetValue(newSheet, missingDict[i]['company_'], row, company_Col)
    ec.SetValue(newSheet, missingDict[i]['contact_'], row, contact_Col)
    ec.SetValue(newSheet, missingDict[i]['phone_'], row, phone_Col)
    ec.SetValue(newSheet, missingDict[i]['email_'], row, email_Col)
    ec.SetValue(newSheet, missingDict[i]['assignedmgr_'], row, assignedmgr_Col)
    ec.SetValue(newSheet, missingDict[i]['assignedrep_'], row, assignedrep_Col)
    ec.SetValue(newSheet, missingDict[i]['leadStatus_'], row, leadStatus_Col)
    ec.SetValue(newSheet, missingDict[i]['leadWon_'], row, leadWon_Col)
    ec.SetValue(newSheet, missingDict[i]['reasonForLoss_'], row, reasonForLoss_Col)
    ec.SetValue(newSheet, missingDict[i]['clientQuote_'], row, clientQuote_Col)
    ec.SetValue(newSheet, missingDict[i]['requestedProduct_'], row, requestedProduct_Col)
    ec.SetValue(newSheet, missingDict[i]['updated_'], row, updated_Col)
    ec.SetValue(newSheet, missingDict[i]['comments_'], row, comments_Col)
    print(i)

ec.Save(newWb, "test1")

"""
 Debug
"""
# print(masterDict)
# print(mgrDict)
# print(missingDict)
# print(missing)

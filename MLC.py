import excelCompare as ec

masterSheet = ec.OpenSheet("Leads_Tracking_Master.xlsx")
mgrSheet = ec.OpenSheet("02182020Stephens_Leads.xlsx")

mastCol = 1
mastRow = 1
mgrCol = 1
mgrRow = 1

endOfMaster = ec.FindLastRow(masterSheet)
endOfMgr = ec.FindLastRow(mgrSheet)

masterDict = {}
mgrDict = {}
missing = []
missingCount = 0
mastLeadCount = 0
found = 0

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

# print(masterDict)
# print(mgrDict)

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
    found += 1
  else:
    missingCount += 1
    missing.append(mgrLead)

#write to new xls wb
newWb = ec.Newb()
newSheet = ec.News(newWb, "Notes")

# print(masterDict)
# print(mgrDict)

# write MASTER values to new wb
for i in masterDict:
  row = masterDict[i]['xlRow']
  if masterDict[i]['assignedmgr_'] != "":
    ec.SetValue(newSheet, masterDict[i]['assignedmgr_'], row, assignedmgr_Col)
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

ec.Save(newWb)


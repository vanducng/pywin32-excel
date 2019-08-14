import win32com.client as win32
import os

def count_row(ws, col_name, row_add):
    row_inc = row_add
    while(ws.Range(col_name + str(row_inc)).Value != None and ws.Range(col_name + str(row_inc)).Value != ""):
        row_inc += 1
    return row_inc
        
def remove_sheet(workbook, sheet_name):
    for sheet in workbook.Sheets:
        if sheet_name == sheet.Name:
            workbook.Worksheets(sheet.Name).Delete()

input_file_name = "DT_MasterReport.xlsx"
output_file_name = "DT_MasterReport_Output.xlsx"

sheet_list = [("FC Report DT- Total DT", "Total DT"), 
             ("FC Report DT-North", "DT-North"),
             ("FC Report DT-Central", "Central"),
             ("FC Report DT-HCME", "HCME"),
             ("FC Report DT-MKD", "MKD")]

iExcel = win32.Dispatch("Excel.Application")
iExcel.Visible = True
iExcel.DisplayAlerts = False
iWb = iExcel.Workbooks.Open(os.path.join(os.getcwd(), input_file_name))

temp_sheet_name = "temp"
remove_sheet(iWb, temp_sheet_name)
wsTemp = iWb.Worksheets.Add()
wsTemp.Name = temp_sheet_name
wsTemp.Range("A3").Value = "FC Report DT"

max_row = 0
for i in range(len(sheet_list)):
    iSheet = iWb.Worksheets(sheet_list[i][0])
    report_name = sheet_list[i][1]
    last_row = count_row(iSheet, "A", 8)
    
    if i == 0:
        wsTemp.Range("A4:A" + str(last_row - 6)).Value = sheet_list[i][1]
        wsTemp.Range("B1:OW" + str(last_row - 5)).Value = iSheet.Range("A6:OV" + str(last_row)).Value
        max_row = last_row - 6
    else:
        wsTemp.Range("A" + str(max_row + 1) + ":A" + str((last_row - 8) + (max_row + 1))).Value = sheet_list[i][1]
        wsTemp.Range("B" + str(max_row + 1) + ":OW" + str((last_row - 8) + (max_row + 1))).Value = iSheet.Range("A9:OV" + str(last_row)).Value
        max_row = (last_row - 8) + (max_row + 1)

oExcel = win32.Dispatch("Excel.Application")
oExcel.Visible = False
oWb = oExcel.Workbooks.Add()
oSheet = oWb.Worksheets(1)
oSheet.Range("A1:OW" + str(max_row)).Value = wsTemp.Range("A1:OW" + str(max_row)).Value 
oWb.SaveAs(os.path.join(os.getcwd(), output_file_name))
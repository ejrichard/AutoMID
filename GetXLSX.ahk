#Requires AutoHotKey 2.0

;Get Spreadsheet of vehicle information
VehicleSpreadsheet := FileSelect(1,,"Select Vehicle Spreadsheet", "*.xlsx")
;Test to see if we got a file.
;MsgBox "The selected file is:  " VechicleSpreadsheet


Xl := ComObject("Excel.Application") 				            ;create a handle to a new excel application
Xl.visible := true
wrkbk1 := Xl.Workbooks.Open(VehicleSpreadsheet)

Test := wrkbk1.Sheets("Sheet1").Range("AN3").Value
MsgBox Test

;sheet := wrkbk1.worksheets.getItem("Sheet1")
;VehicleIDs := sheet.getrange("AN2:AN10")
;sheet := context.workbook.worksheets.getItem("Sheet1")

/*
NumberOfRows := wrkbk1.rows.count
MsgBox ("There are " NumberOfRows " rows in the spreadsheet.")
*/

;for c in wrkbk1.activesheet.Cells
;	msgbox Xl.Worksheets("Sheet1").VLookup(c.value, wrkbk1.activesheet.usedrange, 0)

;wrkbk1.Close := true
Xl.quit()


;Get Spreadsheet of vehicle information
VechicleSpreadsheet := FileSelect(1,,"Select Vehicle Spreadsheet", "*.xlsx")
;Test to see if we got a file.
;MsgBox "The selected file is:  " VechicleSpreadsheet


/*
Xl := ComObject("Excel.Application") 				            ;create a handle to a new excel application
wrkbk1:=Xl.Workbooks.Open(VehicleSpreadsheet)

;for c in wrkbk1.activesheet.Cells
;	msgbox Xl.Worksheets("Sheet1").VLookup(c.value, wrkbk1.activesheet.usedrange, 0)

wrkbk1.Close($true)
xl.quit()
*/


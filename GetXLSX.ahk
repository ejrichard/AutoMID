;
; Variable Setup
;

; Which Column holds the vehicle IDs?
RegNumColumn := "AN"
MakeColumn := "H"
ModelColumn := "I"


;Get Spreadsheet of vehicle information
VehicleSpreadsheet := FileSelect(1,,"Select Vehicle Spreadsheet", "*.xlsx")
;Test to see if we got a file.
;MsgBox "The selected file is:  " VechicleSpreadsheet


Xl := ComObject("Excel.Application") 				            ;create a handle to a new excel application
Xl.visible := true
wrkbk1 := Xl.Workbooks.Open(VehicleSpreadsheet)

RegNum := wrkbk1.Sheets("Sheet1").Range(RegNumColumn . "3").Value
Make := wrkbk1.Sheets("Sheet1").Range(MakeColumn . "3").Value
Model:=wrkbk1.Sheets("Sheet1").Range(ModelColumn . "3").Value
RegNumLength := StrLen(RegNum)

MsgBox Make . " " . Model . " has Registration Number " . RegNum . " with length " . RegNumLength . "."



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

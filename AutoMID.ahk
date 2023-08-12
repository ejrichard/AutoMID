;
; Variable Setup
;

; Which Column holds the vehicle IDs?
RegNumColumn := "AN"
MakeColumn := "H"
ModelColumn := "I"
EmptyRowCount := 0
StartRow := 2

;Get Spreadsheet of vehicle information
VehicleSpreadsheet := FileSelect(1,,"Select Vehicle Spreadsheet", "*.xlsx")
;Test to see if we got a file.
;MsgBox "The selected file is:  " VechicleSpreadsheet

Xl := ComObject("Excel.Application") 				            ;create a handle to a new excel application
Xl.visible := true
wrkbk1 := Xl.Workbooks.Open(VehicleSpreadsheet)

CurrentRow := StartRow

While EmptyRowCount < 20 {
	RegNum := wrkbk1.Sheets("Sheet1").Range(RegNumColumn . CurrentRow).Value
	Make := wrkbk1.Sheets("Sheet1").Range(MakeColumn . CurrentRow).Value
	Model:=wrkbk1.Sheets("Sheet1").Range(ModelColumn . CurrentRow).Value
	RegNumLength := StrLen(RegNum)

	CurrentRow := CurrentRow + 1
	
	if RegNumLength = 0
	{
	EmptyRowCount := EmptyRowCount + 1
	}
	else
	{
	MsgBox Make . " " . Model . " has Registration Number " . RegNum . " with length " . RegNumLength . "."
	EmptyRowCount := 0
	}
	
}

Xl.quit()

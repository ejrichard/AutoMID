;
; Variable Setup
;

; Which Column holds the vehicle IDs?
RegNumColumn := "AN"
MakeColumn := "H"
ModelColumn := "I"
EmptyRowCount := 0
EmptyRowThreshold := 1
StartRow := 5

;Get Spreadsheet of vehicle information
VehicleSpreadsheet := FileSelect(1,,"Select Vehicle Spreadsheet", "*.xlsx")
;Test to see if we got a file.
;MsgBox "The selected file is:  " VechicleSpreadsheet

Xl := ComObject("Excel.Application") 				            ;create a handle to a new excel application
Xl.visible := true
wrkbk1 := Xl.Workbooks.Open(VehicleSpreadsheet)

CurrentRow := StartRow

While EmptyRowCount < EmptyRowThreshold {
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
	EmptyRowCount := 0
	MsgBox Make . " " . Model . " has Registration Number " . RegNum . " with length " . RegNumLength . "."
	Run("https://ownvehicle.askMID.com")
	Sleep 3000
	SendInput("{tab 2}" RegNum "{tab}{space}{tab 2}")
	Sleep 3000
	SendText("`r")
	Sleep 20000
	SendInput("{home}")
	Sleep 1000
	CoordMode "Pixel"
	ImageSearch(&FoundYesX, &FoundYesY, 0, 0, A_ScreenWidth, A_ScreenHeight, ".\assets\Sample-YES.png")
	ImageSearch(&FoundNoX, &FoundNoY, 0, 0, A_ScreenWidth, A_ScreenHeight, ".\assets\Sample-NO.png")
	ImageSearch(&FoundInvalidX, &FoundInvalidY, 0, 0, A_ScreenWidth, A_ScreenHeight, ".\assets\Sample-INVALID.png")
	MsgBox "YES(" . FoundYesX . "," . FoundYesY . "), NO(" . FoundNoX . "," . FoundNoY . "), INVALID(" . FoundInvalidX . "," . FoundInvalidY . ")"  
	}
	
}

Xl.quit()

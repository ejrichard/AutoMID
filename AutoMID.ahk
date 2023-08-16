;
; Variable Setup
;

ResultDateString := A_YYYY . "-" . A_MM . "-" . A_DD . "-" . A_Hour . A_Min

ResultsOutput := A_Desktop . "\AutoMID Results " . ResultDateString . ".html"
RegNumColumn := "AN" ; Which Column holds the vehicle IDs?
MakeColumn := "H"
ModelColumn := "I"
EmptyRowCount := 0
EmptyRowThreshold := 20
StartRow := 2



;Get Spreadsheet of vehicle information
VehicleSpreadsheet := FileSelect(1,,"Select Vehicle Spreadsheet", "*.xlsx")

Xl := ComObject("Excel.Application") 				            ;create a handle to a new excel application
Xl.visible := true
wrkbk1 := Xl.Workbooks.Open(VehicleSpreadsheet)

CurrentRow := StartRow

InputBoxYellow := "0xffcc33"
ResultBoxYellow := "0xffcc08"
UninsuredXRed := "0xd20f0f"
InsuredCheckGreen := "0x0eac5b"
ErrorBoxRed := "0xff0000"

;Initialize Results File
FileAppend
(
"<html>
<head>
<title>AutoMID Results " ResultDateString "</title>
<style>
	table, td {
		border: 1px solid black;
		border-collapse: collapse;
	}
	th {
		overflow: auto;
		text-align: center;
	}
</style>
</head>
<body>
<h1>AutoMID Results " ResultDateString "</h1>
<table>
<th><tr><td>Insured Status</td><td>Registration Number</td><td>Make</td><td>Model</td></tr></th>
"
), ResultsOutput


ResultCheck(TargetWindow)
{
	WinActivate TargetWindow
	SendInput("{home}")
	CoordMode "Pixel", "Window"
	PixelSearch &ResultBoxX, &ResultBoxY, 0, 0, A_ScreenWidth, A_ScreenHeight, ResultBoxYellow
	if isNumber(ResultBoxX) and isNumber(ResultBoxY)
	{
		PixelSearch &InsuredCheckX, &InsuredCheckY, 0, 0, A_ScreenWidth, A_ScreenHeight, InsuredCheckGreen
		PixelSearch &UninsuredXX, &UninsuredXY, 0, 0, A_ScreenWidth, A_ScreenHeight, UninsuredXRed
		if isNumber(InsuredCheckX) and isNumber(UninsuredXX)
		{
		return "ERROR - CONFLICTING RESULTS"
		}
		else
		{
			if isNumber(UninsuredXX)
			{
				return "UNINSURED"
			}
			if isNumber(InsuredCheckX)
			{
				return "INSURED"
			}
		}
	}
	else
	{
		return "NO RESULT"
	}
}


Run("https://ownvehicle.askMID.com")
Sleep Random(1500,2500)
MIDWindow := WinExist("A")
Sleep 3000



While EmptyRowCount < EmptyRowThreshold {
	RegNum := wrkbk1.Sheets(1).Range(RegNumColumn . CurrentRow).Value
	Make := wrkbk1.Sheets(1).Range(MakeColumn . CurrentRow).Value
	Model:=wrkbk1.Sheets(1).Range(ModelColumn . CurrentRow).Value
	RegNumLength := StrLen(RegNum)

	if RegNumLength = 0
	{
	EmptyRowCount := EmptyRowCount + 1
	}
	else
	{
	EmptyRowCount := 0
	Sleep Random(100,600)
	SendInput("{home}")
	Sleep Random(1000,2500)
	CoordMode "Pixel"
	PixelSearch &InputBoxX, &InputBoxY, 0, 0, A_ScreenWidth, A_ScreenHeight, InputBoxYellow
	MouseClick "left", (InputBoxX + random(6,70)), (InputBoxY + Random(10,30))
	Sleep Random(100,600)
	SendInput(RegNum "{tab}{space}{tab 2}")
	Sleep Random(3000,5500)
	SendText("`r")
	Sleep Random(9000,13000) ;Wait for Captcha.
	SendInput("{tab 3}")
	Sleep Random(100,600)
	SendText("`r")
		
	Result := "NO RESULT"
	while Result = "NO RESULT"{
		Sleep Random(5000,7500) ;Wait for results page to load.
		Result := ResultCheck(MIDWindow)
		if Result = "NO RESULT"
		{
		MsgBox "Unable to find result.`nIf there is a Captcha, please complete it`nand then click 'Check This Vehicle'."
		}
	}
	
	FileAppend
	(
	"<tr><td>" Result "</td><td>" RegNum "</td><td>" Make "</td><td>" Model "</td></tr>
	"
	), ResultsOutput

	Sleep Random(2000,3000)
		
	SendInput("{F5}")
	}
	CurrentRow := CurrentRow + 1
}
;Close Out Results File
FileAppend "
(
</table>
</body>
</html>
)", ResultsOutput

Xl.quit()

Run (ResultsOutput)

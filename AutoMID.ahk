;
; Variable Setup
;

; Which Column holds the vehicle IDs?
RegNumColumn := "AN" 
; Which Column contains vehicle Make?
MakeColumn := "H"
; Which Column contains vehicle Model?
ModelColumn := "I"
; After how many empty rows should we assume that we have reached the end of the input data?
EmptyRowThreshold := 20
; Used to skip the header row.  Probably don't change this.
StartRow := 2
; Used to count how many consecutive empty rows have been passed.  This should start at zero.
EmptyRowCount := 0
; Used for output file naming to avoid collisions.  
ResultDateString := A_YYYY . "-" . A_MM . "-" . A_DD . "-" . A_Hour . A_Min
; Where do we put the output file? Can be customized, but do so carefully to avoid overwriting data.
ResultsOutput := A_Desktop . "\AutoMID Results " . ResultDateString . ".html"


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
CaptchaBlue := "0x1a73e8"

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
		padding: 3px 3px 3px 3px;
		border-collapse: collapse;
		vertical-align: top;
	}
    th {
	   	border: 1px solid black;
		position: sticky; 
	   top: 0; 
	   padding-right: 10px;
	   vertical-align: bottom;
	   text-align: left; 
	   background-color: white;
	}
</style>
</head>
<body>
<h1>AutoMID Results " ResultDateString "</h1>
<table>
<tr><th>Insured Status</th><th>Registration Number</th><th>Make</th><th>Model</th></tr>
"
), ResultsOutput


ResultCheck(TargetWindow)
{
	WinActivate TargetWindow
	WinMaximize TargetWindow
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
	WinActivate MIDWindow
	WinMaximize MIDWindow
	SendInput("{home}")
	Sleep Random(1000,2500)
	CoordMode "Pixel", "Window"
	PixelSearch &InputBoxX, &InputBoxY, 0, 0, A_ScreenWidth, A_ScreenHeight, InputBoxYellow
	MouseClick "left", (InputBoxX + random(6,80)), (InputBoxY + Random(10,20))
	Sleep Random(100,600)
	WinActivate MIDWindow
	WinMaximize MIDWindow
	SendInput(RegNum "{tab}{space}{tab 2}")
	Sleep Random(750,6000)
	SendText("`r")
	Sleep Random(9000,13000) ;Wait for Captcha.
	
	
	;Check for Captcha
	CaptchaCheck := true
	while CaptchaCheck = true
	{
		WinActivate MIDWindow
		WinMaximize MIDWindow
		CoordMode "Pixel", "Window"
		PixelSearch &CaptchaX, &CaptchaY, 0, 0, A_ScreenWidth, A_ScreenHeight, CaptchaBlue
		if isNumber(CaptchaX)
		{
			CaptchaCheck := true
		}
		else
		{
			CaptchaCheck := false
		}
		Sleep Random(750,2000)
	}
	
	WinActivate MIDWindow
	WinMaximize MIDWindow
	SendInput("{tab 3}")
	Sleep Random(100,600)
	SendText("`r")
			
	Result := "NO RESULT"
	while Result = "NO RESULT"{
		Sleep Random(5000,7500) ;Wait for results page to load.
		Result := ResultCheck(MIDWindow)
		if Result = "NO RESULT"
		{
		WinActivate MIDWindow
		WinMaximize MIDWindow
		MsgBox "Help!`n`n( 1 ) Confirm that Registration Number " . RegNum . " is entered properly.`n( 2 ) If there is a Captcha, please complete it.`n( 3 ) Click 'Check This Vehicle' at the bottom of the page.`n( 4 ) Click OK on this box`n`nAnd then wait."
		}
		if A_Index > 2
		{
			Result := "ERROR CHECKING"
		}
	}
	
	FileAppend
	(
	"<tr><td>" Result "</td><td>" RegNum "</td><td>" Make "</td><td>" Model "</td></tr>
	"
	), ResultsOutput

	Sleep Random(2000,3000)
	WinActivate MIDWindow
	WinMaximize MIDWindow
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

WinClose MIDWindow
Xl.quit()

Run (ResultsOutput)

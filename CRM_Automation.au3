
#include <Array.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>

; Create application object and open an example workbook
Local $oExcel = _Excel_Open()
If @error Then
	MsgBox($MB_SYSTEMMODAL, "","Unable to create excel object")
EndIf

Local $oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & "\customers.xlsx")

$iLastRow=$oWorkbook.ActiveSheet.Range("A1").SpecialCells($xlCellTypeLastCell).Row

; Read the values of a cell range on sheet 1 of the specified workbook
Local $aResult = _Excel_RangeRead($oWorkbook, 1, "A2:J" & $iLastRow)
If @error Then
	MsgBox($MB_SYSTEMMODAL, "","Unable to save the data in Array")
	Exit
EndIf

_Excel_Close($oExcel)

;open MYCRM
Local $strCRMPath=@ScriptDir &"\CRM.exe"


;Run CRM
Run($strCRMPath)

if @error Then
   MsgBox(0,"Error","Unable to open CRM App")
   Exit
EndIf


; Wait 10 seconds for the CRM window to appear.
Local $hCRMWnd = WinWait("[CLASS:WindowsForms10.Window.8.app.0.378734a]", "", 10)

Local $iDelay=100
Local $irow

for $irow=0 to $iLastRow-2

	Sleep($iDelay)
	;First name
	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a2", $aResult[$irow][0])

	;Last Name
	Sleep($iDelay)
	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a1", $aResult[$irow][1])

	;Gender
	Sleep($iDelay)
	Local $strGender=$aResult[$irow][2]
	If $strGender="Male" Then
		ControlClick($hCRMWnd, "", "WindowsForms10.BUTTON.app.0.378734a2")
	Else
		ControlClick($hCRMWnd, "", "WindowsForms10.BUTTON.app.0.378734a1")
	EndIf

	;Address1
	Sleep($iDelay)
	Local $strAddress=$aResult[$irow][3]

	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a12", $strAddress)

 	;Address2
 	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a11", "Address" & $irow)

	;City
	Sleep($iDelay)
	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a10", $aResult[$irow][4])

	;State
	Sleep($iDelay)

	ControlSend($hCRMWnd, "", "WindowsForms10.COMBOBOX.app.0.378734a1",$aResult[$irow][5])

	ControlSetText($hCRMWnd, "", "WindowsForms10.COMBOBOX.app.0.378734a1", $aResult[$irow][5])

	;Zip
	Sleep($iDelay)
	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a9", 110059)

	;Validate
	Sleep($iDelay)
	ControlClick($hCRMWnd, "", "WindowsForms10.BUTTON.app.0.378734a4")

	;Home Phone
	Sleep($iDelay)
	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a8", $aResult[$irow][6])

	;work Phone
	Sleep($iDelay)
	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a7", $aResult[$irow][7])

	;Mobile
	Sleep($iDelay)
	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a6", $aResult[$irow][7])

	;Personal email
	Sleep($iDelay)
	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a4", $aResult[$irow][8])

	;Work Email
	Sleep($iDelay)
	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a3", $aResult[$irow][9])

	;Active
	Sleep($iDelay)
	ControlClick($hCRMWnd, "", "WindowsForms10.BUTTON.app.0.378734a3")

	;Comments
	Sleep($iDelay)
	ControlSetText($hCRMWnd, "", "WindowsForms10.EDIT.app.0.378734a5","Comments " & $irow)

Next

MsgBox(0,"CRM","Done")

;Close window
Sleep(5000)
WinClose($hCRMWnd)

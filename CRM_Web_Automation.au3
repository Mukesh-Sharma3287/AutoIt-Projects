#include <Array.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <IE.au3>

; Create application object and open an example workbook
Local $iLastRow
Local $oExcel = _Excel_Open()
If @error Then
	MsgBox($MB_SYSTEMMODAL, "","Unable to create excel object")
EndIf

Local $oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & "\customers.xlsx")

$iLastRow=$oWorkbook.ActiveSheet.Range("A1").SpecialCells($xlCellTypeLastCell).Row

; Read the value of a cell range on sheet 2 of the specified workbook
Local $aResult = _Excel_RangeRead($oWorkbook, 1, "A2:J" & $iLastRow)
If @error Then
	MsgBox($MB_SYSTEMMODAL, "","Unable to save the data in Array")
	Exit
EndIf

_Excel_Close($oExcel)

Local $strUrl="https://rpacrm.bubbleapps.io/"
Local $oIE = _IECreate($strUrl, 0, 0 ,0)

$HWND = _IEPropertyGet($oIE, "hwnd")

WinSetState($HWND, "", @SW_MAXIMIZE)

_IEAction($oIE, "visible")

_IELoadWait($oIE)

;Local $oInputs= $oIE.document.GetElementsByTagName("input")
;Local $oInputs = _IETagNameGetCollection($oIE, "input")

Local $oInputs = $oIE.document.getElementsByClassName("bubble-element Input")

Local $iDelay=500
Local $irow

For $irow=0 to $iLastRow-2
	Local $iCount=0
	For $oInput In $oInputs
	   If $iCount=2 or $iCount=3 Then
		   Local $iGender
		   if $aResult[$irow][$iCount]="Male" Then
			  $iGender=1
		   Else
			  $iGender=2
		   EndIf
		   ;Local $oGender = $oIE.document.getElementsByClassName("radio radio-black")
		   Local $oGender= $oIE.document.GetElementsByTagName("label")
		   Local $oMale
		   Local $iGenderCount=1
		   for $oMale in $oGender
			  If $iGenderCount=$iGender Then
				  Sleep($iDelay)
				_IEAction($oMale, "click")
			  ElseIf $iGenderCount=$iGender Then
				Sleep($iDelay)
				_IEAction($oMale, "click")
			  EndIf
			  $iGenderCount=$iGenderCount+1
		   Next
			$iCount=$iCount+1
	   EndIf

;~ 	   ;State
	   If $iCount=5 Then
			Local $oState = $oIE.document.getElementsByClassName("bubble-element Dropdown")
			Local $oSelectState

			for $oSelectState in $oState
				 Sleep($iDelay)
				_IEAction($oSelectState, "focus")
				Sleep($iDelay)
				Send($aResult[$irow][$iCount])
			   ;_IEFormElementSetValue($oSelectState, $aResult[$irow][$iCount])
;~ 				MsgBox(0,"Country",$aResult[$irow][$iCount])
			Next
			$iCount=$iCount+1
	   EndIf

	   Sleep($iDelay)
	  _IEFormElementSetValue($oInput, $aResult[$irow][$iCount])
		$iCount=$iCount+1

	Next

	;Click on add button
	 Local $btns = $oIE.document.getElementsByClassName("bubble-element Button clickable-element")
	 Local $btnsAdd

	 Local $btnCount=1
	 for $btnsAdd in $btns
		if $btnCount=2 Then
			_IEAction($btnsAdd, "click")
		EndIf

		$btnCount=$btnCount+1

	 Next
Next


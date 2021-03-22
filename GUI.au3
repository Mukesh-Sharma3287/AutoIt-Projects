;Open KODA apps
;Run("C:\Users\mukesh.sharma12\Desktop\Auto IT Training Plan\autoit-v3\koda_1.7.3.0\FD.exe")
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Excel.au3>

;Globale variable declares
Global $ExlApp
Global $wbkBook
Global $FilePath
Global $ToolName="Data Entry"

#Region ### START Koda GUI section ### Form=
Global Const $FDataEntry = GUICreate("Data Entry", 360, 345, 369, 48)
$Emp = GUICtrlCreateLabel("Emp ID", 56, 24, 54, 20)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$txtEmpid = GUICtrlCreateInput("", 132, 24, 145, 24)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$Label1 = GUICtrlCreateLabel("Full Name", 56, 75, 74, 20)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$txtFullName = GUICtrlCreateInput("", 132, 75, 145, 24)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$Group1 = GUICtrlCreateGroup("Gender", 56, 128, 225, 57)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$optMale = GUICtrlCreateRadio("Male", 72, 144, 57, 25)
$optFemale = GUICtrlCreateRadio("Female", 160, 144, 73, 25)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$cmbLocation = GUICtrlCreateCombo("", 132, 208, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
GUICtrlSetData(-1, "Gurgaon|Pune|Jaipur|Mysore|Banglore")
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$Label2 = GUICtrlCreateLabel("Location", 56, 208, 63, 20)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
$Save = GUICtrlCreateButton("Submit", 104, 264, 145, 25)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
	    Case $Save
		    Save_Records()
	EndSwitch
WEnd


;handle HotKey event
HotKeySet("{ESC}","StopToolExecution")

Func Save_Records()

   Local $strEmpID
   Local $strFullName
   Local $strGender
   Local $strLocation
   Local $strSkills
   Local $iLastRow
   Local $iCounter

   ;check if file is place in Tool Folder----------------
   $FilePath= @ScriptDir & "\User_Details.xlsx"
   If Not FileExists($FilePath) Then
	  MsgBox(0,$ToolName,"File is not place in Tool Folder." & @CRLF & @CRLF &  "Tool Folder :-" & $FilePath )
	  Return
   EndIf

;~ 	  ;read data from controls and store in variables

	  $strEmpID=GUICtrlRead($txtEmpid)

	  $strFullName=GUICtrlRead($txtFullName)

	  $strLocation=GUICtrlRead($cmbLocation)


	  if GUICtrlRead($optMale)=1 Then
		 $strGender="Male"
	  Else
		  $strGender="Female"
	  EndIf

;~ 	  ;Validation
	  if $strEmpID=""  then
		 MsgBox(0,"Enter Data","Please enter Emp ID to proceed")
		 exit
	  ElseIf $strGender="" Then
		 MsgBox(0,"Enter Data","Please click on gender to proceed")
		 exit
	  ElseIf $strFullName="" Then
		  MsgBox(0,"Enter Data","Please enter full name to proceed")
		  exit
	  ElseIf  $strLocation="" Then
		  MsgBox(0,"Enter Data","Please select location to proceed")
		  exit
	  EndIf


	  ;call Open Excel Function------
	  Open_Excel_App()
	  $iLastRow = $wbkBook.ActiveSheet.Range("A1").SpecialCells($xlCellTypeLastCell).Row + 1

	  ;Write the stored values in Excel sheet
;~ 	  _Excel_RangeWrite($wbkBook,$wbkBook.Activesheet,GUICtrlRead($txtEmpid),"A" & $iLastRow)
;~ 	  _Excel_RangeWrite($wbkBook,$wbkBook.Activesheet,GUICtrlRead($txtFullName),"B" & $iLastRow)
;~ 	  _Excel_RangeWrite($wbkBook,$wbkBook.Activesheet,GUICtrlRead($strGender),"C" & $iLastRow)
;~ 	  _Excel_RangeWrite($wbkBook,$wbkBook.Activesheet,GUICtrlRead($cmbLocation),"D" & $iLastRow)

	  _Excel_RangeWrite($wbkBook,$wbkBook.Activesheet,$strEmpID,"A" & $iLastRow)
	  _Excel_RangeWrite($wbkBook,$wbkBook.Activesheet,$strGender,"B" & $iLastRow)
	  _Excel_RangeWrite($wbkBook,$wbkBook.Activesheet,$strFullName,"C" & $iLastRow)
	  _Excel_RangeWrite($wbkBook,$wbkBook.Activesheet,$strLocation,"D" & $iLastRow)

	  $wbkBook.Activesheet.Range("A:D").WrapText = True		;Wrap the Cell text

	    ;Store the Form Control values in the Variable
	  GUICtrlSetData ($txtEmpid,"")
	  GUICtrlSetData ($txtFullName,"")

	  MsgBox(0,"saved","Record has been saved successfully!!")

    ;call Close Excel Function------
    Close_Excel_App()

EndFunc

;Create Excel Applicaton & Open Workbook in it===================================================
Func Open_Excel_App()
   $ExlApp=_Excel_Open(False)	;create Excel Application

   If Not IsObj($ExlApp) Then
	  MsgBox(0,$ToolName,"Unable to create Excel Application. Please try again..")
	  Exit
   EndIf

   $wbkBook=_Excel_BookOpen($ExlApp,$FilePath,False,True)	;open Excel Database Workbook

   If Not IsObj($wbkBook) Then
	  MsgBox(0,$ToolName,"Unable to open Excel Database Workbook. Please try again..")
	  Exit
   EndIf

EndFunc

;Saves & Close Excel Application and Workbook====================================================
Func Close_Excel_App()
   _Excel_BookClose($wbkBook,True)		;close Excel WorkBook
   _Excel_Close($ExlApp,False,True)		;Close Excel Application
EndFunc

;Stop Tool execution=============================================================================
Func StopToolExecution()
   MsgBox (0,$ToolName,"Tool execution has been stopped by user.")
   ;call Close Excel Function------
   Close_Excel_App()
   ;-------------------------------
   Exit
EndFunc
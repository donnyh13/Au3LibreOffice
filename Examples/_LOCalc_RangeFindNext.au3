#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oSrchDesc, $oResult
	Local $aavData[3]
	Local $avRowData[8]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number and String Values I want in Column A to H.
	$avRowData[0] = 1 ; A8
	$avRowData[1] = 5 ; B8
	$avRowData[2] = "2a" ; C8
	$avRowData[3] = 7 ; D8
	$avRowData[4] = 1 ; E8
	$avRowData[5] = 2 ; F8
	$avRowData[6] = 1 ; G8
	$avRowData[7] = "b2" ; H8
	$aavData[0] = $avRowData

	$avRowData[0] = 10 ; A9
	$avRowData[1] = 20 ; B9
	$avRowData[2] = 30 ; C9
	$avRowData[3] = 5 ; D9
	$avRowData[4] = 77 ; E9
	$avRowData[5] = 24 ; F9
	$avRowData[6] = 89 ; G9
	$avRowData[7] = 58 ; H9
	$aavData[1] = $avRowData

	$avRowData[0] = "A String" ; A10
	$avRowData[1] = -1700 ; B10
	$avRowData[2] = 2000 ; C10
	$avRowData[3] = "A Different String" ; D10
	$avRowData[4] = 67 ; E10
	$avRowData[5] = 28 ; F10
	$avRowData[6] = 36 ; G10
	$avRowData[7] = "1992 as a String" ; H10
	$aavData[2] = $avRowData

	; Retrieve Cell range A8 to H10
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A8", "H10")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Search Descriptor, Backwards = False, Search Rows = True, Match Case = False, Search in = Values
	$oSrchDesc = _LOCalc_SearchDescriptorCreate($oSheet, False, True, False, $LOC_SEARCH_IN_VALUES)
	If @error Then _ERROR($oDoc, "Failed to create a Search descriptor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will perform a Find in the Sheet, looking for any cells that contain ""2""." & _
			" I will then set the background color of each individual result to a random background color.")

	; Set result to Null for my first call, so I can call find function in a loop. After the first call it will be overwritten with an Object.
	$oResult = Null

	Do
		; Perform a Find for the Entire Sheet, one at a time, Search for any cells containing 2.
		$oResult = _LOCalc_RangeFindNext($oSheet, $oSrchDesc, "2", $oResult)
		If @error Then _ERROR($oDoc, "Failed to perform FindNext for the Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		If IsObj($oResult) Then
			; Set the Cell Background color to a Random value.
			_LOCalc_CellBackColor($oResult, Random($LO_COLOR_BLACK, $LO_COLOR_WHITE, 1), False)
			If @error Then _ERROR($oDoc, "Failed to set Cell Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

			MsgBox($MB_OK + $MB_TOPMOST, Default, "The Search was successful, I found a Cell containing ""2"" and set the background color.")

		Else
			MsgBox($MB_OK + $MB_TOPMOST, Default, "The Search was successful, but I did not find a Cell containing ""2"".")
			ExitLoop
		EndIf
	Until Not IsObj($oResult)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc

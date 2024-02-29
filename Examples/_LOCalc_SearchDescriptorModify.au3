#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oCell, $oSrchDesc, $oResult, $oColumn
	Local $aavData[3]
	Local $avRowData[8]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Fill my arrays with the desired Number and String Values I want in Column A to H.
	$avRowData[0] = "Seventy" ; A8
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
	$avRowData[4] = "Seventy Seven" ; E9
	$avRowData[5] = 24 ; F9
	$avRowData[6] = 89 ; G9
	$avRowData[7] = 58 ; H9
	$aavData[1] = $avRowData

	$avRowData[0] = "A Seven" ; A10
	$avRowData[1] = -1700 ; B10
	$avRowData[2] = 2000 ; C10
	$avRowData[3] = "A Different Seven" ; D10
	$avRowData[4] = 1 ; E10
	$avRowData[5] = "a lowercase seven" ; F10
	$avRowData[6] = 36 ; G10
	$avRowData[7] = "1992 as a String" ; H10
	$aavData[2] = $avRowData

	; Retrieve Cell range A8 to H10
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A8", "H10")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell "A2"
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "A2")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set Cell A2 to a formula
	_LOCalc_CellFormula($oCell, "=B8 + D9")
	If @error Then _ERROR($oDoc, "Failed to set Cell formula. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell "C3"
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "C3")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set Cell C3 to a formula
	_LOCalc_CellFormula($oCell, "=SUM(G8:G10)")
	If @error Then _ERROR($oDoc, "Failed to set Cell formula. Error:" & @error & " Extended:" & @extended)

	; Set the columns A to H's width to Optimal.
	For $i = 0 To 7
		; Retrieve Column's Object
		$oColumn = _LOCalc_RangeColumnGetObjByPosition($oSheet, $i)
		If @error Then _ERROR($oDoc, "Failed to retrieve Column Object by Position. Error:" & @error & " Extended:" & @extended)

		; Set Column's width to optimal.
		_LOCalc_RangeColumnWidth($oColumn, True)
		If @error Then _ERROR($oDoc, "Failed to set Cell width to Optimal. Error:" & @error & " Extended:" & @extended)

	Next

	; Create a Search Descriptor, Backwards = True, Search Rows = False, Match Case = True, Search in = Values, Entire Cell and Regular Expressions = False, Wildcards = True.
	$oSrchDesc = _LOCalc_SearchDescriptorCreate($oSheet, True, False, True, $LOC_SEARCH_IN_VALUES, False, False, True)
	If @error Then _ERROR($oDoc, "Failed to create a Search descriptor. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will perform a Find in the Sheet, looking for the first cell that contain ""Seven*""." & _
			" If found, I will  set the background color of each result to a random background color.")

	; Perform a Find for the Entire Sheet, Search for any cells containing Seven*, ("*" is a wildcard meaning any sequence of characters)
	$oResult = _LOCalc_RangeFindNext($oSheet, $oSrchDesc, "Seven*")
	If @error Then _ERROR($oDoc, "Failed to perform Replace for the Cell Range. Error:" & @error & " Extended:" & @extended)

	If IsObj($oResult) Then
		; Set the Cell Background color to a Random value.
		_LOCalc_CellBackColor($oResult, Random($LOC_COLOR_BLACK, $LOC_COLOR_WHITE, 1), False)
		If @error Then _ERROR($oDoc, "Failed to set Cell Background color. Error:" & @error & " Extended:" & @extended)

	EndIf

	MsgBox($MB_OK, "", "I will now modify the Search descriptor to search in Rows instead of Columns, forward instead of backwards, " & _
			"Use Regular Expressions instead of Wildcards, and search in Formulas. I will then perform a search for the Regular Expression ""\d{2}"", " & _
			"which will search for any cells containing two digits in their formulas.")

	; Modify the Search Descriptor, Backwards = False, Search Rows = True, skip Match Case, Search in = Formulas, Regular Expressions = True, Wildcards = False.
	_LOCalc_SearchDescriptorModify($oSrchDesc, False, True, Null, $LOC_SEARCH_IN_FORMULAS, Null, True, False)
	If @error Then _ERROR($oDoc, "Failed to create a Search descriptor. Error:" & @error & " Extended:" & @extended)

	; Perform a Find for the Entire Sheet, Search for any cells containing \d{2}.
	$oResult = _LOCalc_RangeFindNext($oSheet, $oSrchDesc, "\d{2}")
	If @error Then _ERROR($oDoc, "Failed to perform Replace for the Cell Range. Error:" & @error & " Extended:" & @extended)

	If IsObj($oResult) Then
		; Set the Cell Background color to a Random value.
		_LOCalc_CellBackColor($oResult, Random($LOC_COLOR_BLACK, $LOC_COLOR_WHITE, 1), False)
		If @error Then _ERROR($oDoc, "Failed to set Cell Background color. Error:" & @error & " Extended:" & @extended)

	EndIf

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc

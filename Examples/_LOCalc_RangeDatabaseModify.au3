#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCellRange, $oDataBase
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell range A1 to A5
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "A5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the range A1:A5 as a Database Range.
	$oDataBase = _LOCalc_RangeDatabaseAdd($oDoc, $oCellRange, "My AutoIt Range")
	If @error Then _ERROR($oDoc, "Failed to add Cell Range to list of Database Ranges. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_RangeDatabaseModify($oDoc, $oDataBase)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Database Range settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Database Range's current settings are as follows: " & @CRLF & _
			"The Range currently covered by this Database range is: " & _LOCalc_RangeGetAddressAsName($avSettings[0]) & @CRLF & _
			"The Database Range name is: " & $avSettings[1] & @CRLF & _
			"Is the top row considered a Header? True/False: " & $avSettings[2] & @CRLF & _
			"Is the last row considered a Totals row? True/False: " & $avSettings[3] & @CRLF & _
			"Will columns/Rows be added or removed when the range size is changed by an update operation? True/False: " & $avSettings[4] & @CRLF & _
			"If Columns/Rows are added or removed, will formatting be increased or decreased accordingly? True/False: " & $avSettings[5] & @CRLF & _
			"Are the contents of this range ignored when saving the document? True/False: " & $avSettings[6] & @CRLF & _
			"Is Auto Filter enabled for this Range? True/False: " & $avSettings[7] & @CRLF & @CRLF & _
			"I will now modify the Database Range's settings.")

	; Retrieve the Range B2:C5.
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "B2", "C5")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Database Range's settings, change the Range to cover B2:C5, set the name to "AU3LibreOffice Range", Column Header = False, Totals Row = True,
	; Add/Remove Cells = False, Keep formatting = False, Ignore contents of range when saving document = True, Auto Filter = True.
	_LOCalc_RangeDatabaseModify($oDoc, $oDataBase, $oCellRange, "AU3LibreOffice Range", False, True, False, False, True, True)
	If @error Then _ERROR($oDoc, "Failed to set Database Range settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_RangeDatabaseModify($oDoc, $oDataBase)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Database Range settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Database Range's new settings are as follows: " & @CRLF & _
			"The Range currently covered by this Database range is: " & _LOCalc_RangeGetAddressAsName($avSettings[0]) & @CRLF & _
			"The Database Range name is: " & $avSettings[1] & @CRLF & _
			"Is the top row considered a Header? True/False: " & $avSettings[2] & @CRLF & _
			"Is the last row considered a Totals row? True/False: " & $avSettings[3] & @CRLF & _
			"Will columns/Rows be added or removed when the range size is changed by an update operation? True/False: " & $avSettings[4] & @CRLF & _
			"If Columns/Rows are added or removed, will formatting be increased or decreased accordingly? True/False: " & $avSettings[5] & @CRLF & _
			"Are the contents of this range ignored when saving the document? True/False: " & $avSettings[6] & @CRLF & _
			"Is Auto Filter enabled for this Range? True/False: " & $avSettings[7])

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc

#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oDoc2, $oSheet, $oSheet2, $oCellRange
	Local $aavData[3]
	Local $avRowData[2]
	Local $sFilePathName, $sPath

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, $oDoc2, $sPath, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Sheet named "New Sheet".
	$oSheet = _LOCalc_SheetAdd($oDoc, "New Sheet")
	If @error Then _ERROR($oDoc, $oDoc2, $sPath, "Failed to create a new Sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Activate the new Sheet
	_LOCalc_SheetActivate($oDoc, $oSheet)
	If @error Then _ERROR($oDoc, $oDoc2, $sPath, "Failed to activate the Sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill my arrays with the desired Number and String Values I want in Column A to B.
	$avRowData[0] = "Seventy" ; A1
	$avRowData[1] = 5 ; B1
	$aavData[0] = $avRowData

	$avRowData[0] = 10 ; A2
	$avRowData[1] = 20 ; B2
	$aavData[1] = $avRowData

	$avRowData[0] = "Testing" ; A3
	$avRowData[1] = -1700 ; B3
	$aavData[2] = $avRowData

	; Retrieve Cell range A1 to B3
	$oCellRange = _LOCalc_RangeGetCellByName($oSheet, "A1", "B3")
	If @error Then _ERROR($oDoc, $oDoc2, $sPath, "Failed to retrieve Cell Range Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the range with Data
	_LOCalc_RangeData($oCellRange, $aavData)
	If @error Then _ERROR($oDoc, $oDoc2, $sPath, "Failed to fill Cell Range. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now save this Document to the desktop folder then open a new document and link this Sheet from Document 1 into the new document.")

	$sFilePathName = @DesktopDir & "\TestExportDoc_" & @MDAY & ".ods"

	; Save The Document To Desktop Directory.
	$sPath = _LOCalc_DocSaveAs($oDoc, $sFilePathName)
	If @error Then _ERROR($oDoc, $oDoc2, $sPath, "Failed to Save the Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a New, visible, Blank Libre Office Document.
	$oDoc2 = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, $oDoc2, $sPath, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Link the Sheet "New Sheet" from Document 1 into this Document, I'll use a function to retrieve the Sheet's name, but I could also just call the Sheet name as "New Sheet"
	; Use Link mode Normal
	$oSheet2 = _LOCalc_SheetLink($oDoc, $oDoc2, _LOCalc_SheetName($oDoc, $oSheet), $LOC_SHEET_LINK_MODE_NORMAL, True)
	If @error Then _ERROR($oDoc, $oDoc2, $sPath, "Failed to Link Calc Sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Activate the new Sheet
	_LOCalc_SheetActivate($oDoc2, $oSheet2)
	If @error Then _ERROR($oDoc, $oDoc2, $sPath, "Failed to activate Sheet. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have linked the Sheet from Document 1 into Document 2, the new sheet is called: " & _LOCalc_SheetName($oDoc2, $oSheet2))

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, $oDoc2, $sPath, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the document.
	_LOCalc_DocClose($oDoc2, False)
	If @error Then _ERROR($oDoc, $oDoc2, $sPath, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Delete the file.
	FileDelete($sPath)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
If @error Then Return _ERROR($oDoc, $oDoc2, $sPath, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $oDoc2, $sPath, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	If IsObj($oDoc2) Then _LOCalc_DocClose($oDoc2, False)
	If IsString($sPath) Then FileDelete($sPath)
	Exit
EndFunc

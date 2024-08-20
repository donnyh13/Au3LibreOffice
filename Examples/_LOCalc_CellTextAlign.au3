#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell
	Local $avSettings[0]
	Local $iMicrometers

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve Cell B2
	$oCell = _LOCalc_RangeGetCellByName($oSheet, "B2")
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Object. Error:" & @error & " Extended:" & @extended)

	; Insert some text into Cell B2
	_LOCalc_CellString($oCell, "Some Text")
	If @error Then _ERROR($oDoc, "Failed to set Cell Text. Error:" & @error & " Extended:" & @extended)

	; Convert 1/4" to Micrometers
	$iMicrometers = _LOCalc_ConvertToMicrometer(0.25)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended)

	; Set the Cell's Text Alignment settings to, Horizontal Alignment = $LOC_CELL_ALIGN_HORI_LEFT, Vertical Align = $LOC_CELL_ALIGN_VERT_MIDDLE,
	; Indent = 1/4"
	_LOCalc_CellTextAlign($oCell, $LOC_CELL_ALIGN_HORI_LEFT, $LOC_CELL_ALIGN_VERT_MIDDLE, $iMicrometers)
	If @error Then _ERROR($oDoc, "Failed to set the Cell's settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_CellTextAlign($oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Cell's settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Cell's current Text Alignment settings are as follows: " & @CRLF & _
			"The Horizontal Alignment is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Vertical Alignment is (See UDF Constants): " & $avSettings[1] & @CRLF & _
			"The amount of Indentation is (in Micrometers): " & $avSettings[2])

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
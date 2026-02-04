#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oCellStyle
	Local $avSettings[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Object for Default Cell Style.
	$oCellStyle = _LOCalc_CellStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve the Object for Cell Style named ""Default"". Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Cell Style's Border Width to $LOC_BORDERWIDTH_THICK for all four sides, and $LOC_BORDERWIDTH_THIN for the diagonal borders.
	_LOCalc_CellStyleBorderWidth($oCellStyle, $LOC_BORDERWIDTH_THICK, $LOC_BORDERWIDTH_THICK, $LOC_BORDERWIDTH_THICK, $LOC_BORDERWIDTH_THICK, $LOC_BORDERWIDTH_THIN, $LOC_BORDERWIDTH_THIN)
	If @error Then _ERROR($oDoc, "Failed to set the Cell Style's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Cell Style's Border color to $LO_COLOR_BRICK for all four sides, and $LO_COLOR_BLUE for the diagonal borders.
	_LOCalc_CellStyleBorderColor($oCellStyle, $LO_COLOR_BRICK, $LO_COLOR_BRICK, $LO_COLOR_BRICK, $LO_COLOR_BRICK, $LO_COLOR_BLUE, $LO_COLOR_BLUE)
	If @error Then _ERROR($oDoc, "Failed to set the Cell Style's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_CellStyleBorderColor($oCellStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Cell Style's current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Default Cell Style's Border color settings are as follows: " & @CRLF & _
			"Top Border color (as a RGB Color Integer): " & $avSettings[0] & @CRLF & _
			"Bottom Border color (as a RGB Color Integer): " & $avSettings[1] & @CRLF & _
			"Left Border color (as a RGB Color Integer): " & $avSettings[2] & @CRLF & _
			"Right Border color (as a RGB Color Integer): " & $avSettings[3] & @CRLF & _
			"Top-Left to Bottom-Right Diagonal Border color (as a RGB Color Integer): " & $avSettings[4] & @CRLF & _
			"Bottom-Left to Top-Right Diagonal Border color (as a RGB Color Integer): " & $avSettings[5])

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

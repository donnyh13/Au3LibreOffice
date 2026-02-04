#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oCellStyle
	Local $avColor[0]

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Object for Default Cell Style.
	$oCellStyle = _LOCalc_CellStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve the Object for Cell Style named ""Default"". Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Cell Style's Background color to Blue.
	_LOCalc_CellStyleBackColor($oCellStyle, $LO_COLOR_BLUE)
	If @error Then _ERROR($oDoc, "Failed to set Cell Style Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Cell Style's current background color setting, Return will be an array in order of function parameters.
	$avColor = _LOCalc_CellStyleBackColor($oCellStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell Style's Background color. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, """Default"" Cell Style's Background color settings are: " & @CRLF & _
			"The Cell Style's Background color is (as a RGB Color Integer): " & $avColor[0] & @CRLF & _
			"Is the Cell Style's Background color transparent? True/False: " & $avColor[1])

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

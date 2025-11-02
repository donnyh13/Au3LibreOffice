#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oCellStyle
	Local $avSettings[0]
	Local $iMicrometers, $iMicrometers2

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Object for Default Cell Style.
	$oCellStyle = _LOCalc_CellStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve the Object for Cell Style named ""Default"". Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Cell Style's Border Width to $LOC_BORDERWIDTH_THICK for all four sides.
	_LOCalc_CellStyleBorderWidth($oCellStyle, $LOC_BORDERWIDTH_THICK, $LOC_BORDERWIDTH_THICK, $LOC_BORDERWIDTH_THICK, $LOC_BORDERWIDTH_THICK)
	If @error Then _ERROR($oDoc, "Failed to set the Cell Style's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert 1/4" to Micrometers
	$iMicrometers = _LO_ConvertToMicrometer(0.25)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set "Default" Cell Style Border padding to 1/4"
	_LOCalc_CellStyleBorderPadding($oCellStyle, $iMicrometers)
	If @error Then _ERROR($oDoc, "Failed to set Cell style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_CellStyleBorderPadding($oCellStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Cell Style's current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Default Cell Style's Border Padding settings are as follows (in Micrometers): " & @CRLF & _
			"The ""All"" value is: " & $avSettings[0] & @CRLF & _
			"Top Border Width is: " & $avSettings[1] & @CRLF & _
			"Bottom Border Width is: " & $avSettings[2] & @CRLF & _
			"Left Border Width is: " & $avSettings[3] & @CRLF & _
			"Right Border Width is: " & $avSettings[4])

	; Convert 1/2" to Micrometers
	$iMicrometers2 = _LO_ConvertToMicrometer(0.5)
	If @error Then _ERROR($oDoc, "Failed to convert from inches to Micrometers. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set "Default" Cell Style Border padding to, Top and Right, 1/4", Bottom and left, 1/2".
	_LOCalc_CellStyleBorderPadding($oCellStyle, Null, $iMicrometers, $iMicrometers2, $iMicrometers2, $iMicrometers)
	If @error Then _ERROR($oDoc, "Failed to set Cell style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOCalc_CellStyleBorderPadding($oCellStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Cell Style's current settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Default Cell Style's Border Padding settings are as follows (in Micrometers): " & @CRLF & _
			"The ""All"" value is: " & $avSettings[0] & @CRLF & _
			"Top Border Width is: " & $avSettings[1] & @CRLF & _
			"Bottom Border Width is: " & $avSettings[2] & @CRLF & _
			"Left Border Width is: " & $avSettings[3] & @CRLF & _
			"Right Border Width is: " & $avSettings[4])

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

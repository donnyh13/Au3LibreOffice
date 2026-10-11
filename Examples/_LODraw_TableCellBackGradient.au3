#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oTable, $oCell
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Current active page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current active page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a new Table.
	$oTable = _LODraw_TableInsert($oPage, 5000, 4000, 3, 3)
	If @error Then _ERROR($oDoc, "Failed to insert a Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve top left Table Cell Object
	$oCell = _LODraw_TableCellGetObjByPosition($oTable, 0, 0)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table cell Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Cell Gradient settings to: Preset Gradient name = $LOD_GRAD_NAME_BLUE_TOUCH
	_LODraw_TableCellBackGradient($oDoc, $oCell, $LOD_GRAD_NAME_BLUE_TOUCH)
	If @error Then _ERROR($oDoc, "Failed to set Cell settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Cell settings. Return will be an array in order of function parameters.
	$avSettings = _LODraw_TableCellBackGradient($oDoc, $oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cell's Gradient settings are as follows: " & @CRLF & _
			"The Gradient name is: " & $avSettings[0] & @CRLF & _
			"The type of Gradient is, (see UDF constants): " & $avSettings[1] & @CRLF & _
			"The number of steps to increment color is: " & $avSettings[2] & @CRLF & _
			"The horizontal offset percentage for the gradient is: " & $avSettings[3] & @CRLF & _
			"The vertical offset percentage for the gradient is: " & $avSettings[4] & @CRLF & _
			"The rotation angle for the gradient is, in degrees: " & $avSettings[5] & @CRLF & _
			"The percentage of area not covered by the ending color is: " & $avSettings[6] & @CRLF & _
			"The starting color is (as a RGB Color Integer): " & $avSettings[7] & @CRLF & _
			"The ending color is (as a RGB Color Integer): " & $avSettings[8] & @CRLF & _
			"The starting color intensity percentage is: " & $avSettings[9] & @CRLF & _
			"The ending color intensity percentage is: " & $avSettings[10])

	; Modify the Cell's Gradient settings to: skip pre-set gradient name, Gradient type = $LOD_GRAD_TYPE_SQUARE, increment steps = 150,
	; horizontal (X) offset = 25%, vertical offset (Y) = 56%, rotational angle = 135 degrees, percentage not covered by "From" color = 50%
	; Starting color = $LO_COLOR_ORANGE, Ending color = $LO_COLOR_TEAL, Starting color intensity = 100%, ending color intensity = 68%
	_LODraw_TableCellBackGradient($oDoc, $oCell, Null, $LOD_GRAD_TYPE_SQUARE, 150, 25, 56, 135, 50, $LO_COLOR_ORANGE, $LO_COLOR_TEAL, 100, 68)
	If @error Then _ERROR($oDoc, "Failed to set Cell settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Cell's settings. Return will be an array in order of function parameters.
	$avSettings = _LODraw_TableCellBackGradient($oDoc, $oCell)
	If @error Then _ERROR($oDoc, "Failed to retrieve Cell settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Cell's Gradient settings are as follows: " & @CRLF & _
			"The Gradient name is: " & $avSettings[0] & @CRLF & _
			"The type of Gradient is, (see UDF constants): " & $avSettings[1] & @CRLF & _
			"The number of steps to increment color is: " & $avSettings[2] & @CRLF & _
			"The horizontal offset percentage for the gradient is: " & $avSettings[3] & @CRLF & _
			"The vertical offset percentage for the gradient is: " & $avSettings[4] & @CRLF & _
			"The rotation angle for the gradient is, in degrees: " & $avSettings[5] & @CRLF & _
			"The percentage of area not covered by the ending color is: " & $avSettings[6] & @CRLF & _
			"The starting color is (as a RGB Color Integer): " & $avSettings[7] & @CRLF & _
			"The ending color is (as a RGB Color Integer): " & $avSettings[8] & @CRLF & _
			"The starting color intensity percentage is: " & $avSettings[9] & @CRLF & _
			"The ending color intensity percentage is: " & $avSettings[10])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LODraw_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LODraw_DocClose($oDoc, False)
	Exit
EndFunc

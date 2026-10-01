#include <MsgBoxConstants.au3>

#include "..\LibreOfficeDraw.au3"

Example()

Func Example()
	Local $oDoc, $oPage, $oTextBox, $oTextCursor
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LODraw_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Draw Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Page.
	$oPage = _LODraw_PageCurrent($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve current page. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Textbox.
	$oTextBox = _LODraw_ShapeTextBoxInsert($oPage, 10500, 5000, -1, 1500)
	If @error Then _ERROR($oDoc, "Failed to  insert a Text Box. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in the Textbox.
	$oTextCursor = _LODraw_ShapeCreateTextCursor($oTextBox)
	If @error Then _ERROR($oDoc, "Failed to create a Text Cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LODraw_CursorInsertString($oTextCursor, "Created by AutoIt!")
	If @error Then _ERROR($oDoc, "Failed to insert some text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set animation type to Scroll alternate, Begin on the left and go towards the right, Don't begin with the text visible, leave the text visible when the animation finishes,
	; Repeat 5 times, increment the text by 23 units, interpret the increment in pixels, and delay each animation cycle by 250 ms.
	_LODraw_ShapeTextAttrAnimation($oTextBox, $LOD_ANIMATION_TYPE_SCROLL_ALTERNATE, $LOD_ANIMATION_DIR_RIGHT, False, True, 5, 23, True, 250)
	If @error Then _ERROR($oDoc, "Failed to modify Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current Shape settings. Return will be an array in order of function parameters.
	$avSettings = _LODraw_ShapeTextAttrAnimation($oTextBox)
	If @error Then _ERROR($oDoc, "Failed to retrieve Shape settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Text Box's settings are as follows: " & @CRLF & _
			"The Animation type is (See UDF Constants): " & $avSettings[0] & @CRLF & _
			"The Animation direction is (See UDF Constants): " & $avSettings[1] & @CRLF & _
			"Is the Text visible and inside shape when the effect is applied? True/False: " & $avSettings[2] & @CRLF & _
			"Is the Text visible and inside shape when the effect is finished? True/False: " & $avSettings[3] & @CRLF & _
			"How many time will the animation repeat?: " & $avSettings[4] & @CRLF & _
			"How much is the text incremented by?: " & $avSettings[5] & @CRLF & _
			"Is the increment measured in Pixels? True/False: " & $avSettings[6] & @CRLF & _
			"How much delay is between each animation repeat? (In Milliseconds): " & $avSettings[7])

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

#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $aiReturn, $aiBackup

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Backup the current document coordinates. Return will be an array in order of function parameters.
	$aiBackup = _LOImpress_DocPosAndSize($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve document position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the X coordinate to 50, Y coordinate to 150, Width to 500, Height to 600
	_LOImpress_DocPosAndSize($oDoc, 50, 150, 500, 600)
	If @error Then _ERROR($oDoc, "Failed to set document settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current document coordinates. Return will be an array in order of function parameters.
	$aiReturn = _LOImpress_DocPosAndSize($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve document position. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The document's current position and size is as follows: " & @CRLF & _
			"X Coordinate = " & $aiReturn[0] & @CRLF & _
			"Y Coordinate = " & $aiReturn[1] & @CRLF & _
			"The document's width, in pixels, is: " & $aiReturn[2] & @CRLF & _
			"The document's height, in pixels, is: " & $aiReturn[3])

	; Restore the original Document size.
	_LOImpress_DocPosAndSize($oDoc, $aiBackup[0], $aiBackup[1], $aiBackup[2], $aiBackup[3])
	If @error Then _ERROR($oDoc, "Failed to set document settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc

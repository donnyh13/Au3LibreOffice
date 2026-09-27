#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $asStyles
	Local $sStyles = ""

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Shape Style called AutoIt-Test
	_LOImpress_ShapeStyleCreate($oDoc, "AutoIt-Test")
	If @error Then _ERROR($oDoc, "Failed to create a Shape Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of all user-created Shape Style names.
	$asStyles = _LOImpress_ShapeStylesGetNames($oDoc, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($asStyles) - 1
		$sStyles &= $asStyles[$i] & @CRLF
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I found " & UBound($asStyles) & " User-Created Shape Style(s). With the following name(s):" & @CRLF & @CRLF & $sStyles)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc

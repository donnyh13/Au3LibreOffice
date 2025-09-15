#include <Array.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc
	Local $asFonts

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array list of font names
	$asFonts = _LOImpress_FontsGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Array of font names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "There were " & @extended & " fonts found. I will now display the results using _ArrayDisplay. There will be four columns, " & @CRLF & _
			"-the first column contains the font name, " & @CRLF & _
			"-the second column contains the style name, " & @CRLF & _
			"-the third column contains the Font weight (Bold) value, (see constants)," & @CRLF & _
			"-the fourth column contains the font slant (Italic), (See constants).")

	_ArrayDisplay($asFonts, Default, Default, 0, Default, "Font Name|Style Name|Font Weight|Font Slant")

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

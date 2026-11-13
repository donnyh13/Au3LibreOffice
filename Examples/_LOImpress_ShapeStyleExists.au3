#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oStyle
	Local $bExists

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Shape Style named "AutoIt-Test"
	$oStyle = _LOImpress_ShapeStyleCreate($oDoc, "AutoIt-Test")
	If @error Then _ERROR($oDoc, "Failed to create a Shape Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; See if the new Style exists.
	$bExists = _LOImpress_ShapeStyleExists($oDoc, "AutoIt-Test")
	If @error Then _ERROR($oDoc, "Failed to query for a Shape Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Does the Shape Style ""AutoIt-Test"" exist? True/False: " & $bExists)

	; Delete the Shape Style named "AutoIt-Test", Force it to be deleted replacing anywhere it's used with "Filled Green"
	_LOImpress_ShapeStyleDelete($oDoc, $oStyle, True, "Filled Green")
	If @error Then _ERROR($oDoc, "Failed to delete a shape style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; See if the new Style exists.
	$bExists = _LOImpress_ShapeStyleExists($oDoc, "AutoIt-Test")
	If @error Then _ERROR($oDoc, "Failed to query for a Shape Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Now does the Shape Style ""AutoIt-Test"" exist? True/False: " & $bExists)

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

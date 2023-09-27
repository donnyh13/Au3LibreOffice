#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc
	Local $iFormatKey

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a New Number Format Key.
	$iFormatKey = _LOWriter_FormatKeyCreate($oDoc, "#,##0.000")
	If (@error > 0) Then _ERROR("Failed to create a Format Key. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I created a new Number format key. Its Format Key number is: " & $iFormatKey & " It looks like this: " & _
			_LOWriter_FormatKeyGetString($oDoc, $iFormatKey) & @CRLF & @CRLF & "Press Ok to delete it.")

	;Delete the Format Key.
	_LOWriter_FormatKeyDelete($oDoc, $iFormatKey)
	If (@error > 0) Then _ERROR("Failed to delete a Format Key. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Does the document still have the Number format I created? True/False: " & _LOWriter_FormatKeyExists($oDoc, $iFormatKey))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

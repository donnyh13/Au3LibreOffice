#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text before I set the character style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text before I set the Character Style to ""Emphasis"". ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now set the current character style to ""Emphasis.""")

	; Set the character style to Emphasis using the ViewCursor.
	_LOWriter_CharStyleSet($oDoc, $oViewCursor, "Emphasis")
	If @error Then _ERROR($oDoc, "Failed to set the Character style. Error:" & @error & " Extended:" & @extended)

	; Insert some text after I set the character style.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some emphasized text.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc

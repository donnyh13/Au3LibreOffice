#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $aoEndnotes

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Endnote at the end of this line. ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert a Endnote at the ViewCursor
	_LOWriter_EndnoteInsert($oDoc, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert a Endnote. Error:" & @error & " Extended:" & @extended)

	; Insert some more text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, " More Text.")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert another Endnote at the ViewCursor.
	_LOWriter_EndnoteInsert($oDoc, $oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to insert a Endnote. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to delete the first Endnote.")

	; Retrieve an array of Endnote Objects.
	$aoEndnotes = _LOWriter_EndnotesGetList($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve an array of Endnote objects. Error:" & @error & " Extended:" & @extended)

	; Delete the first Endnote returned
	_LOWriter_EndnoteDelete($aoEndnotes[0])
	If @error Then _ERROR($oDoc, "Failed to delete a Endnote. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc   ;==>Example

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc   ;==>_ERROR

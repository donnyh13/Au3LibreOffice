
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text at the Viewcursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text to demonstrate inserting text into a document." & @CR & "This is a New line. Note:" & _
			" In Libre Office, Autoit's @CR is used for a paragraph break (such as is inserted by pressing Enter key), " & @LF & _
			"and @LF is used for a Line break, Such as when you press Shift + Enter. If you use the standard Autoit @CRLF," & @CRLF & _
			" this double space, with a paragraph break and a line break results. Try turning on Formatting marks to see the difference, normally " & _
			"done by pressing CTRL + F10.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc




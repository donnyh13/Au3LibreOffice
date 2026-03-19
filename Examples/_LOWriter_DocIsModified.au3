
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $bReturn

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Check if the document has been modified since being saved or created.
	$bReturn = _LOWriter_DocIsModified($oDoc)
	If (@error > 0) Then _ERROR("Failed to query document status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Has the document been modified since being created or saved? True/False: " & $bReturn)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text for demonstration.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Check if the document has been modified since being saved or created.
	$bReturn = _LOWriter_DocIsModified($oDoc)
	If (@error > 0) Then _ERROR("Failed to query document status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Now has the document been modified since being created or saved? True/False: " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc




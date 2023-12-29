
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oPar
	Local $aoPars[0]

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The First Paragraph" & @CR & "I'm going to delete this paragraph." & @CR & _
			"A Third Paragraph." & @CR & " A Fourth Paragraph.")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Create a list of Paragraph Objects
	$aoPars = _LOWriter_ParObjCreateList($oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve array of paragraphs. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to delete the second paragraph.")

	$oPar = $aoPars[1]

	;Delete the second paragraph.
	_LOWriter_ParObjDelete($oPar)
	If (@error > 0) Then _ERROR("Failed to delete the paragraph. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc


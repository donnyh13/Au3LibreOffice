#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Footnote at the end of this line. ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert a Footnote at the ViewCursor.
	_LOWriter_FootnoteInsert($oDoc, $oViewCursor)
	If @error Then _ERROR("Failed to insert a Footnote. Error:" & @error & " Extended:" & @extended)

	; Modify the Document's Footnote continuation settings to: End of Footnote continuation notice = "Cont. On Page ", Next Page Continuation notice =
	; "Cont. from Page "
	_LOWriter_FootnoteSettingsContinuation($oDoc, "Cont. On Page ", "Cont. from Page ")
	If @error Then _ERROR("Failed to modify Footnote settings. Error:" & @error & " Extended:" & @extended)

	; Retrieve the current Footnote settings. Return will be an array in order of function parameters.
	$avSettings = _LOWriter_FootnoteSettingsContinuation($oDoc)
	If @error Then _ERROR("Failed to retrieve Footnote settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Document's current Footnote Continuation settings are as follows: " & @CRLF & _
			"The text that appears at the end of a Footnote to indicate it is continued is: " & $avSettings[0] & @CRLF & _
			"The text that appears at the beginning of a Footnote to indicate it was continued is: " & $avSettings[1] & @CRLF & _
			"Note, Libre Office automatically inserts a Page number after the continuation notices, so where ""Page"" appears, there would be a page number " & _
			"directly after it.")

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc


#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iCursorType

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)

	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Get the type of Cursor.
	$iCursorType = _LOWriter_CursorGetType($oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve Cursor Object type. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The possible cursor type values are: " & @CRLF & _
			"$LOW_CURTYPE_TEXT_CURSOR (1)" & @CRLF & _
			"$LOW_CURTYPE_TABLE_CURSOR (2)" & @CRLF & _
			"$LOW_CURTYPE_VIEW_CURSOR (3)")

	;Display a message depending on what type the cursor is.
	Switch $iCursorType

		Case $LOW_CURTYPE_TEXT_CURSOR
			MsgBox($MB_OK, "", "The Cursor Type is a Text Cursor, with an integer value of : " & $iCursorType & " — Or $LOW_CURTYPE_TEXT_CURSOR")
		Case $LOW_CURTYPE_TABLE_CURSOR
			MsgBox($MB_OK, "", "The Cursor Type is a Table Cursor, with an integer value of : " & $iCursorType & " — Or $LOW_CURTYPE_TABLE_CURSOR")
		Case $LOW_CURTYPE_VIEW_CURSOR
			MsgBox($MB_OK, "", "The Cursor Type is a View Cursor, with an integer value of : " & $iCursorType & " — Or $LOW_CURTYPE_VIEW_CURSOR")
		Case Else
			MsgBox($MB_OK, "", "Something went wrong.")
	EndSwitch

	;Close the document.
	_LOWriter_DocClose($oDoc, False)

	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Following Error codes returned: Error:" & _
			@error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc




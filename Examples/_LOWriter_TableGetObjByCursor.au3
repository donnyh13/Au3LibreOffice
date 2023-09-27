#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oTableNewObj

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I am going to insert a table named ""AutoitTest"".")

	;Create a Table, 2 rows, 2 columns
	$oTable = _LOWriter_TableCreate($oDoc, 2, 2, Null, Null, "AutoitTest")
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now move the ViewCursor up into the table so that I can retrieve the Table Object again.")

	;Move the View Cursor up one line, which will put it in a Table cell.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_UP, 1, False)

	;Retrieve the Table Object again using the ViewCursor.
	$oTableNewObj = _LOWriter_TableGetObjByCursor($oDoc, $oViewCursor)
	If (@error > 0) Then _ERROR("Failed to retrieve Text Table Object using View Cursor. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now delete the table.")

	;Delete the Table.
	_LOWriter_TableDelete($oDoc, $oTableNewObj)
	If (@error > 0) Then _ERROR("Failed to delete Text Table. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

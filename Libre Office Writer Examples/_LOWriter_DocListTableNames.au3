
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oTable, $oTable2
	Local $asTables

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create the Table, 5 rows, 3 columns.
	$oTable = _LOWriter_TableCreate($oDoc, 5, 3)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document at the View Cursor's location.
	$oTable = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert a paragraph break to separate the tables.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR)
	If (@error > 0) Then _ERROR("Failed to insert Text. Error:" & @error & " Extended:" & @extended)

	;Create another Table, 2 rows, 3 columns.
	$oTable2 = _LOWriter_TableCreate($oDoc, 2, 3)
	If (@error > 0) Then _ERROR("Failed to create Text Table. Error:" & @error & " Extended:" & @extended)

	;Insert the Table into the document at the View Cursor's location.
	$oTable2 = _LOWriter_TableInsert($oDoc, $oViewCursor, $oTable2)
	If (@error > 0) Then _ERROR("Failed to insert Text Table. Error:" & @error & " Extended:" & @extended)

	;Retrieve Array of Table names
	$asTables = _LOWriter_DocListTableNames($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve list of Text Tables. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The document contains the following Table names:" & @CR & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($asTables) -1
		;Insert the Frame names.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $asTables[$i] & @CR)
		If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)
	Next

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc




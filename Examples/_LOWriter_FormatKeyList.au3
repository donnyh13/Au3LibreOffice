
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iResults
	Local $avKeys

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a new DateFormatKey H:MM:SS (Hour, Minute, Second, AM/PM, Month Day Name(Day) Year
	_LOWriter_DateFormatKeyCreate($oDoc, "H:MM:SS AM/PM -- M/NNN(D)/YYYY")
	If (@error > 0) Then _ERROR("Failed to create a Date/Time Format Key. Error:" & @error & " Extended:" & @extended)

	;Create a New Number Format Key.
	_LOWriter_FormatKeyCreate($oDoc, "#,##0.000")
	If (@error > 0) Then _ERROR("Failed to create a Format Key. Error:" & @error & " Extended:" & @extended)

	;Retrieve an Array of Format Keys. With Boolean value of whether each is a UserCreated key or not., search for all Format Key types.
	$avKeys = _LOWriter_FormatKeyList($oDoc, True, False, $LOW_FORMAT_KEYS_ALL)
	If (@error > 0) Then _ERROR("Failed to retrieve an array of Date/Time Format Keys. Error:" & @error & " Extended:" & @extended)
	$iResults = @extended

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Format Key" & Chr(9) & Chr(9) & "Format Key String" & Chr(9) & Chr(9) & "Is User Created?" & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To $iResults - 1
		;List the keys in the document, separate each column by tabs.
		_LOWriter_DocInsertString($oDoc, $oViewCursor, $avKeys[$i][0] & Chr(9) & Chr(9) & Chr(9) & $avKeys[$i][1] & Chr(9) & Chr(9) & Chr(9) & $avKeys[$i][2] & @CR)
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


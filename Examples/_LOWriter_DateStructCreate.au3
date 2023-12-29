
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iDateFormatKey
	Local $tDateStruct1, $tDateStruct2

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a Date Structure, leave it blank so it will be set to the current date/Time.
	$tDateStruct1 = _LOWriter_DateStructCreate()
	If (@error > 0) Then _ERROR("Failed to create a Date structure. Error:" & @error & " Extended:" & @extended)

	;Create a second Date Structure, Year = 1992, Month = 03, Day = 28, Hour = 15, Minute = 43, Second = 25, Nano = 765, UTC =True
	$tDateStruct2 = _LOWriter_DateStructCreate(1992, 03, 28, 15, 43, 25, 765, True)
	If (@error > 0) Then _ERROR("Failed to create a Date structure. Error:" & @error & " Extended:" & @extended)

	;Create or retrieve a DateFormat Key, Hour, Minute, Second, AM/PM, Month Day Year
	$iDateFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "HH:MM:SS AM/PM MM/DD/YYYY")
	If (@error > 0) Then _ERROR("Failed to create a Date Format Key. Error:" & @error & " Extended:" & @extended)

	;Insert a Date and Time text Field at the View Cursor. Fixed = True, Set the Date to the first DateStruct I created, and set DateTime Format Key to
	;the Key I created.
	_LOWriter_FieldDateTimeInsert($oDoc, $oViewCursor, False, True, $tDateStruct1, Null, Null, $iDateFormatKey)
	If (@error > 0) Then _ERROR("Failed to insert a field. Error:" & @error & " Extended:" & @extended)

	;Insert 2 newlines.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR)
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Insert another Date and Time text Field at the View Cursor. Fixed = True, Set the Date to the Second DateStruct I created, and set DateTime Format
	;Key to the Key I created.
	_LOWriter_FieldDateTimeInsert($oDoc, $oViewCursor, False, True, $tDateStruct2, Null, Null, $iDateFormatKey)
	If (@error > 0) Then _ERROR("Failed to insert a field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc


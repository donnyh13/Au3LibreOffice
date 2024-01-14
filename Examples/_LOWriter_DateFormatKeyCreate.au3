#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iFormatKey

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a new DateFormatKey H:MM:SS (Hour, Minute, Second, AM/PM, Month Day Name(Day) Year
	$iFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "H:MM:SS AM/PM -- M/NNN(D)/YYYY")
	If @error Then _ERROR($oDoc, "Failed to create a Date/Time Format Key. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I created a new DateTime format key, its key number is: " & $iFormatKey & " and it looks like this: " & _
			_LOWriter_DateFormatKeyGetString($oDoc, $iFormatKey) & @CRLF & @CRLF & "Press Ok to insert a Date Field using this key into the document.")

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a Date and Time text Field at the View Cursor. using the DateTime Format Key I just created.
	_LOWriter_FieldDateTimeInsert($oDoc, $oViewCursor, False, Null, Null, Null, Null, $iFormatKey)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc

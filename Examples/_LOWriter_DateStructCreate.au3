#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iDateFormatKey
	Local $tDateStruct1, $tDateStruct2

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_ViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure, leave it blank so it will be set to the current date/Time.
	$tDateStruct1 = _LOWriter_DateStructCreate()
	If @error Then _ERROR($oDoc, "Failed to create a Date structure. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a second Date Structure, Year = 1992, Month = 03, Day = 28, Hour = 15, Minute = 43, Second = 25, Nanoseconds = 765, UTC =True
	$tDateStruct2 = _LOWriter_DateStructCreate(1992, 03, 28, 15, 43, 25, 765, True)
	If @error Then _ERROR($oDoc, "Failed to create a Date structure. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create or retrieve a DateFormat Key, Hour, Minute, Second, AM/PM, Month Day Year
	$iDateFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "HH:MM:SS AM/PM MM/DD/YYYY")
	If @error Then _ERROR($oDoc, "Failed to create a Date Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Date and Time text Field at the View Cursor. Fixed = True, Set the Date to the first DateStruct I created, and set DateTime Format Key to
	; the Key I created.
	_LOWriter_FieldDateTimeInsert($oDoc, $oViewCursor, False, True, $tDateStruct1, Null, Null, $iDateFormatKey)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert 2 newlines.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert another Date and Time text Field at the View Cursor. Fixed = True, Set the Date to the Second DateStruct I created, and set DateTime Format
	; Key to the Key I created.
	_LOWriter_FieldDateTimeInsert($oDoc, $oViewCursor, False, True, $tDateStruct2, Null, Null, $iDateFormatKey)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc

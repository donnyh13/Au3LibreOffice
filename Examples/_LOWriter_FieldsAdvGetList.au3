#include <Array.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iDateFormatKey
	Local $avFields
	Local $tDateStruct

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a Regular field at the end of this line, it wont be listed.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure, Year = 1992, Month = 03, Day = 28, Hour = 15, Minute = 43, Second = 25, Nanoseconds = 765, UTC =True
	$tDateStruct = _LOWriter_DateStructCreate(1992, 03, 28, 15, 43, 25, 765, True)
	If @error Then _ERROR($oDoc, "Failed to create a Date structure. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create or retrieve a DateFormat Key, Hour, Minute, Second, AM/PM, Month Day Year
	$iDateFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "HH:MM:SS AM/PM MM/DD/YYYY")
	If @error Then _ERROR($oDoc, "Failed to create a Date Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Date and Time text Field at the View Cursor. Fixed = True, Set the Date to my previously created DateStruct,and set DateTime Format Key to the  first
	; Key I created.
	_LOWriter_FieldDateTimeInsert($oDoc, $oViewCursor, False, True, $tDateStruct, Null, Null, $iDateFormatKey)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert 2 newlines.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a Regular field at the end of this line, it wont be listed.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Count Field at the View Cursor. Set count type to $LOW_FIELD_COUNT_TYPE_WORDS, Overwrite to False, and Number format to
	; $LOW_NUM_STYLE_CHARS_LOWER_LETTER
	_LOWriter_FieldStatCountInsert($oDoc, $oViewCursor, $LOW_FIELD_COUNT_TYPE_WORDS, False, $LOW_NUM_STYLE_ARABIC)
	If @error Then _ERROR($oDoc, "Failed to insert a Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert 2 newlines.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a Regular field at the end of this line, it wont be listed.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Combined character Field at the View Cursor. Insert the Characters "ABCDEF
	_LOWriter_FieldCombCharInsert($oDoc, $oViewCursor, False, "ABCDEF")
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert 2 newlines.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a Doc info field at the end of this line, it wont be listed.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Doc Info Title Field at the View Cursor. Set is Fixed = True, Title = "This is a Title Field."
	_LOWriter_FieldDocInfoTitleInsert($oDoc, $oViewCursor, False, True, "Title Field.")
	If @error Then _ERROR($oDoc, "Failed to insert a Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert 2 newlines.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "Please insert a User Field here.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I have no functions currently that insert any ""Advanced"" Fields, if you would like a demonstration, please follow to prompts, First go to: " & @CRLF & _
			"Insert, Field, More Fields, Variables, Click on ""User Field"" and enter ""Test"" as ""name"", and 1234 as value, then click insert. Press Ok on this MsgBox.")

	; Insert 2 newlines.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "Please insert a DDE Field here.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Next go to: " & @CRLF & _
			"Insert, Field, More Fields, Variables, Click on ""DDE Field"" and enter ""Test"" as ""name"", and 1234 as value, then click insert. Press Ok on this MsgBox.")

	; Insert 2 newlines.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "Please insert a Database Name Field here.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Next go to: " & @CRLF & _
			"Insert, Field, More Fields, Database, Click on ""Database Name"" and drop down ""Bibliography"", select ""Biblio"", then click insert. Press Ok on this MsgBox.")

	; Retrieve an array of Advanced Fields. The Doc Info Field, and regular fields, wont be listed in this array.
	$avFields = _LOWriter_FieldsAdvGetList($oDoc, $LOW_FIELD_ADV_TYPE_ALL)
	If @error Then _ERROR($oDoc, "Failed to search for Fields. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_ArrayDisplay($avFields)

	; Retrieve an array of Advanced Fields again, this time only list $LOW_FIELD_ADV_TYPE_USER and $LOW_FIELD_ADV_TYPE_DATABASE_NAME type fields.
	$avFields = _LOWriter_FieldsAdvGetList($oDoc, BitOR($LOW_FIELD_ADV_TYPE_DATABASE_NAME, $LOW_FIELD_ADV_TYPE_USER))
	If @error Then _ERROR($oDoc, "Failed to search for Fields. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	_ArrayDisplay($avFields)

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

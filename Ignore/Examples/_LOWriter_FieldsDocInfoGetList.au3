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
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Regular field at the end of this line, it wont be listed.--> ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Create a Date Structure, Year = 1992, Month = 03, Day = 28, Hour = 15, Minute = 43, Second = 25, Nano = 765, UTC =True
	$tDateStruct = _LOWriter_DateStructCreate(1992, 03, 28, 15, 43, 25, 765, True)
	If @error Then _ERROR("Failed to create a Date structure. Error:" & @error & " Extended:" & @extended)

	; Create or retrieve a DateFormat Key, Hour, Minute, Second, AM/PM, Month Day Year
	$iDateFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "HH:MM:SS AM/PM MM/DD/YYYY")
	If @error Then _ERROR("Failed to create a Date Format Key. Error:" & @error & " Extended:" & @extended)

	; Insert a Date and Time text Field at the View Cursor. Fixed = True, Set the Date to my previously created DateStruct,and set DateTime Format Key to the  first
	; Key I created.
	_LOWriter_FieldDateTimeInsert($oDoc, $oViewCursor, False, True, $tDateStruct, Null, Null, $iDateFormatKey)
	If @error Then _ERROR("Failed to insert a field. Error:" & @error & " Extended:" & @extended)

	; Insert 2 newlines.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Regular field at the end of this line, it wont be listed.--> ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert a Count Field at the View Cursor. Set count type to $LOW_FIELD_COUNT_TYPE_WORDS, Overwrite to False, and Number format to
	; $LOW_NUM_STYLE_CHARS_LOWER_LETTER
	_LOWriter_FieldStatCountInsert($oDoc, $oViewCursor, $LOW_FIELD_COUNT_TYPE_WORDS, False, $LOW_NUM_STYLE_ARABIC)
	If @error Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	; Insert 2 newlines.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Regular field at the end of this line, it wont be listed.--> ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert a Combined character Field at the View Cursor. Insert the Characters "ABCDEF
	_LOWriter_FieldCombCharInsert($oDoc, $oViewCursor, False, "ABCDEF")
	If @error Then _ERROR("Failed to insert a field. Error:" & @error & " Extended:" & @extended)

	; Insert 2 newlines.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Doc info field at the end of this line.--> ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert a Doc Info Title Field at the View Cursor. Set is Fixed = True, Title = "This is a Title Field."
	_LOWriter_FieldDocInfoTitleInsert($oDoc, $oViewCursor, False, True, "Title Field.")
	If @error Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	; Insert 2 newlines.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Doc info field at the end of this line.--> ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert a Document Information Create Author Field at the View Cursor. Set is Fixed = False
	_LOWriter_FieldDocInfoCreateAuthInsert($oDoc, $oViewCursor, False, False)
	If @error Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	; Insert 2 newlines.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, @CR & @CR)
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a Doc info field at the end of this line.--> ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert a Document Information Creation Date/Time Field at the View Cursor. Set is Fixed = True, and Date Format Key to the one I just created.
	_LOWriter_FieldDocInfoCreateDateTimeInsert($oDoc, $oViewCursor, False, True, $iDateFormatKey)
	If @error Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	; Retrieve an array of Doc Info Fields. The Doc Info Field wont be listedin this array.
	$avFields = _LOWriter_FieldsDocInfoGetList($oDoc, $LOW_FIELD_DOCINFO_TYPE_ALL)
	If @error Then _ERROR("Failed to search for Fields. Error:" & @error & " Extended:" & @extended)

	_ArrayDisplay($avFields)

	; Retrieve an array of Doc info Fields again, this time only list $LOW_FIELD_DOCINFO_TYPE_CREATE_AUTH and $LOW_FIELD_DOCINFO_TYPE_TITLE type fields.
	$avFields = _LOWriter_FieldsDocInfoGetList($oDoc, BitOR($LOW_FIELD_DOCINFO_TYPE_CREATE_AUTH, $LOW_FIELD_DOCINFO_TYPE_TITLE))
	If @error Then _ERROR("Failed to search for Fields. Error:" & @error & " Extended:" & @extended)

	_ArrayDisplay($avFields)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $iDateFormatKey
	Local $avSettings
	Local $tDateStruct

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure, set the Date to Year, 1992, 03, 28, 12:00
	$tDateStruct = _LOWriter_DateStructCreate(1992, 03, 28, 12, 00)
	If @error Then _ERROR($oDoc, "Failed to create a Date structure. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Document's Created by settings to, Created by = "Nik", Date to the previously created Day Structure.
	_LOWriter_DocGenPropPrint($oDoc, "Nik", $tDateStruct)
	If @error Then _ERROR($oDoc, "Failed to modify Document settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_CursorViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create or retrieve a DateFormat Key, Hour, Minute, Second, AM/PM, Month Day Year
	$iDateFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "H:MM:SS AM/PM MM/DD/YYYY")
	If @error Then _ERROR($oDoc, "Failed to create a Date Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Document Information Printed By Date/Time Field at the View Cursor. Set is Fixed = True, and Date Format Key to the one I just created.
	$oField = _LOWriter_FieldDocInfoPrintDateTimeInsert($oDoc, $oViewCursor, False, True, $iDateFormatKey)
	If @error Then _ERROR($oDoc, "Failed to insert a Field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to modify the Doc Info Field settings.")

	; Create or retrieve a different DateFormat Key, two-digit year, Day, Month, Hour Minute, Second.
	$iDateFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "YY/DDD/MMM HH:MM:SS")
	If @error Then _ERROR($oDoc, "Failed to create a Date Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Doc Info Printed by Date/Time Field settings. Set Fixed to False, Set the Date/Time Format to the new Key I just created.
	_LOWriter_FieldDocInfoPrintDateTimeModify($oDoc, $oField, False, $iDateFormatKey)
	If @error Then _ERROR($oDoc, "Failed to modify field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current Field settings. Return will be an array, with elements in order of function parameters.
	$avSettings = _LOWriter_FieldDocInfoPrintDateTimeModify($oDoc, $oField)
	If @error Then _ERROR($oDoc, "Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Doc Info Field settings are: " & @CRLF & _
			"Is the content of this field fixed? True/ False: " & $avSettings[0] & @CRLF & _
			"The Date/Time format key used to display the date is: " & $avSettings[1] & " And looks like: " & @CRLF & _
			_LOWriter_DateFormatKeyGetString($oDoc, $avSettings[1]))

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

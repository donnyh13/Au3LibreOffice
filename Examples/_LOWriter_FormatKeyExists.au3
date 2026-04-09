#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $iFormatKey, $i2ndFormatKey, $iResults
	Local $avKeys
	Local $bExists

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new DateFormatKey H:MM:SS (Hour, Minute, Second, AM/PM, Month Day Name(Day) Year
	$iFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "H:MM:SS AM/PM -- M/NNN(D)/YYYY")

	; Check if the new key exists.
	$bExists = _LOWriter_FormatKeyExists($oDoc, $iFormatKey)
	If @error Then _ERROR($oDoc, "Failed to search for a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If ($bExists = True) Then
		MsgBox($MB_OK + $MB_TOPMOST, Default, "I created a new DateTime format key.")

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "I Failed to create a new DateTime format key.")
	EndIf

	; Create a New Number Format Key.
	$i2ndFormatKey = _LOWriter_FormatKeyCreate($oDoc, "#,##0.000")
	If @error Then _ERROR($oDoc, "Failed to create a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Check if the new key exists, searching in Number Format Keys only.
	$bExists = _LOWriter_FormatKeyExists($oDoc, $i2ndFormatKey, $LOW_FORMAT_KEYS_NUMBER)
	If @error Then _ERROR($oDoc, "Failed to search for a Format Key. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	If ($bExists = True) Then
		MsgBox($MB_OK + $MB_TOPMOST, Default, "I created a new Number format key.")

	Else
		MsgBox($MB_OK + $MB_TOPMOST, Default, "I Failed to create a new Number format key.")
	EndIf

	; Retrieve an Array of Format Keys. User created ones only.
	$avKeys = _LOWriter_FormatKeysGetList($oDoc, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve an array of Date/Time Format Keys. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	$iResults = @extended

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_CursorViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "Format Key" & Chr(9) & Chr(9) & "Format Key String" & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To $iResults - 1
		; List the keys in the document, separate each column by tabs.
		_LOWriter_CursorInsertString($oDoc, $oViewCursor, $avKeys[$i][0] & Chr(9) & Chr(9) & Chr(9) & $avKeys[$i][1] & @CR)
		If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	Next

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

#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $asNumStyles, $asNumStylesDisplay
	Local $sStyles = ""

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Numbering Style names.
	$asNumStyles = _LOWriter_NumStylesGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Numbering style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Numbering Style names.
	$asNumStylesDisplay = _LOWriter_NumStylesGetNames($oDoc, False, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Numbering style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now insert a list of available Numbering styles. There are " & @extended & " results.")

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The available Numbering Styles in this document are:" & @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To (UBound($asNumStyles) - 1)
		If ($asNumStyles[$i] <> $asNumStylesDisplay[$i]) Then
			$sStyles &= $asNumStyles[$i] & @LF & "(Display Name: " & $asNumStylesDisplay[$i] & ")" & @CR & @CR

		Else
			$sStyles &= $asNumStyles[$i] & @CR & @CR
		EndIf
	Next

	; Insert the Style names.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, $sStyles)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply List 3 Numbering Style.
	_LOWriter_NumStyleCurrent($oDoc, $oViewCursor, "List 3")
	If @error Then _ERROR($oDoc, "Failed to apply Numbering Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Numbering Style names that are applied to the document
	$asNumStyles = _LOWriter_NumStylesGetNames($oDoc, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Numbering style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Numbering Style names that are applied to the document
	$asNumStylesDisplay = _LOWriter_NumStylesGetNames($oDoc, False, True, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Numbering style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now insert a list of used Numbering styles, if any. There are " & @extended & " results.")

	; Move the View Cursor to the end of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_End)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Numbering Styles currently in use in this document are:" & @CR & @CR, True)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sStyles = ""

	For $i = 0 To (UBound($asNumStyles) - 1)
		If ($asNumStyles[$i] <> $asNumStylesDisplay[$i]) Then
			$sStyles &= $asNumStyles[$i] & @LF & "(Display Name: " & $asNumStylesDisplay[$i] & ")" & @CR & @CR

		Else
			$sStyles &= $asNumStyles[$i] & @CR & @CR
		EndIf
	Next

	; Insert the Style names.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, $sStyles)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc

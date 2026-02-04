#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $asParStyles, $asParStylesDisplay
	Local $sStyles = ""

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Paragraph Style names.
	$asParStyles = _LOWriter_ParStylesGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Paragraph style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Paragraph Style Display names.
	$asParStylesDisplay = _LOWriter_ParStylesGetNames($oDoc, False, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Paragraph style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now insert a list of available Paragraph styles. There are " & @extended & " results.")

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Paragraph Styles available in this document are:" & @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To (UBound($asParStyles) - 1)
		If ($asParStyles[$i] <> $asParStylesDisplay[$i]) Then
			$sStyles &= $asParStyles[$i] & @LF & "(Display Name: " & $asParStylesDisplay[$i] & ")" & @CR & @CR

		Else
			$sStyles &= $asParStyles[$i] & @CR & @CR
		EndIf
	Next

	; Insert the Style names.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, $sStyles)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Paragraph Style names that are applied to the document
	$asParStyles = _LOWriter_ParStylesGetNames($oDoc, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Paragraph style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Paragraph Style display names that are applied to the document
	$asParStylesDisplay = _LOWriter_ParStylesGetNames($oDoc, False, True, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of Paragraph style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now display a list of used Paragraph styles. There is " & @extended & " result.")

	; Move the View Cursor to the end of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_End)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Paragraph Styles currently in use in this document are:" & @CR & @CR, True)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sStyles = ""

	For $i = 0 To (UBound($asParStyles) - 1)
		If ($asParStyles[$i] <> $asParStylesDisplay[$i]) Then
			$sStyles &= $asParStyles[$i] & @LF & "(Display Name: " & $asParStylesDisplay[$i] & ")" & @CR & @CR

		Else
			$sStyles &= $asParStyles[$i] & @CR & @CR
		EndIf
	Next

	; Insert the Paragraph Style names.
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

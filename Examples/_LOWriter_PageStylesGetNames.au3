#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $asPageStyles, $asPageStylesDisplay
	Local $sStyles = ""

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Page Style names.
	$asPageStyles = _LOWriter_PageStylesGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of page style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Page Style display names.
	$asPageStylesDisplay = _LOWriter_PageStylesGetNames($oDoc, False, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of page style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now insert a list of available Page styles. There are " & @extended & " results.")

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The available Page Styles in this document are:" & @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To (UBound($asPageStyles) - 1)
		If ($asPageStyles[$i] <> $asPageStylesDisplay[$i]) Then
			$sStyles &= $asPageStyles[$i] & @LF & "(Display Name: " & $asPageStylesDisplay[$i] & ")" & @CR & @CR

		Else
			$sStyles &= $asPageStyles[$i] & @CR & @CR
		EndIf
	Next

	; Insert the Style names.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, $sStyles & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the Page Style "Left Page"
	_LOWriter_PageStyleCurrent($oDoc, $oViewCursor, "Left Page")
	If @error Then _ERROR($oDoc, "Failed to Apply Page Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START, 1, False)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Page Style names that are applied to the document
	$asPageStyles = _LOWriter_PageStylesGetNames($oDoc, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of page style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Array of Page Style display names that are applied to the document
	$asPageStylesDisplay = _LOWriter_PageStylesGetNames($oDoc, False, True, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve array of page style names. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now insert a list of used Page styles. There is " & @extended & " result.")

	; Move the View Cursor to the end of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_End)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Page Styles currently in use in this document are:" & @CR & @CR, True)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sStyles = ""

	For $i = 0 To (UBound($asPageStyles) - 1)
		If ($asPageStyles[$i] <> $asPageStylesDisplay[$i]) Then
			$sStyles &= $asPageStyles[$i] & @LF & "(Display Name: " & $asPageStylesDisplay[$i] & ")" & @CR & @CR

		Else
			$sStyles &= $asPageStyles[$i] & @CR & @CR
		EndIf
	Next

	; Insert the Style names.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, $sStyles & @CR)
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

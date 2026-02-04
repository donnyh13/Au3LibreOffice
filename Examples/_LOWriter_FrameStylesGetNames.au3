#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oFrame
	Local $asFrameStyles, $asFrameStylesDisplay
	Local $sStyles = ""

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a new Frame Style named "Test Style"
	_LOWriter_FrameStyleCreate($oDoc, "Test Style")
	If @error Then _ERROR($oDoc, "Failed to create a Frame Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert a Frame with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Frame
	$oFrame = _LOWriter_FrameCreate($oDoc, $oViewCursor, Null, 3000, 3000)
	If @error Then _ERROR($oDoc, "Failed to insert a Text Frame. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the Frame Style to "Labels"
	_LOWriter_FrameStyleCurrent($oDoc, $oFrame, "Labels")
	If @error Then _ERROR($oDoc, "Failed to set the Text Frame style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all available Frame Style Names.
	$asFrameStyles = _LOWriter_FrameStylesGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame style list. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all available Frame Style display Names.
	$asFrameStylesDisplay = _LOWriter_FrameStylesGetNames($oDoc, False, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame style list. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now insert a list of all available Frame styles. There are " & @extended & " results.")

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The available Frame Styles in this document are:" & @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To (UBound($asFrameStyles) - 1)
		If ($asFrameStyles[$i] <> $asFrameStylesDisplay[$i]) Then
			$sStyles &= $asFrameStyles[$i] & @LF & "(Display Name: " & $asFrameStylesDisplay[$i] & ")" & @CR & @CR

		Else
			$sStyles &= $asFrameStyles[$i] & @CR & @CR
		EndIf
	Next

	; Insert the Style names.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, $sStyles)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all user-created Frame Style Names.
	$asFrameStyles = _LOWriter_FrameStylesGetNames($oDoc, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame style list. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all user-created Frame Style display Names.
	$asFrameStylesDisplay = _LOWriter_FrameStylesGetNames($oDoc, True, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame style list. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now insert a list of all user-created Frame styles. There are " & @extended & " results.")

	; Move the View Cursor to the end of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_End)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the right once, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Frame Styles that are user-created are:" & @CR & @CR, True)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sStyles = ""

	For $i = 0 To (UBound($asFrameStyles) - 1)
		If ($asFrameStyles[$i] <> $asFrameStylesDisplay[$i]) Then
			$sStyles &= $asFrameStyles[$i] & @LF & "(Display Name: " & $asFrameStylesDisplay[$i] & ")" & @CR & @CR

		Else
			$sStyles &= $asFrameStyles[$i] & @CR & @CR
		EndIf
	Next

	; Insert the Style names.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, $sStyles)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all applied Frame Style Names.
	$asFrameStyles = _LOWriter_FrameStylesGetNames($oDoc, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame style list. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all applied Frame Style display Names.
	$asFrameStylesDisplay = _LOWriter_FrameStylesGetNames($oDoc, False, True, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve Frame style list. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "I will now insert a list of all Frame styles used in this document. There are " & @extended & " results.")

	; Move the View Cursor to the end of the document
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_End)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the start of the document, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_START, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the View Cursor to the right once, select the text.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_RIGHT, 1, True)
	If @error Then _ERROR($oDoc, "Failed to move ViewCursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "The Frame Styles currently in use in this document are:" & @CR & @CR, True)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sStyles = ""

	For $i = 0 To (UBound($asFrameStyles) - 1)
		If ($asFrameStyles[$i] <> $asFrameStylesDisplay[$i]) Then
			$sStyles &= $asFrameStyles[$i] & @LF & "(Display Name: " & $asFrameStylesDisplay[$i] & ")" & @CR & @CR

		Else
			$sStyles &= $asFrameStyles[$i] & @CR & @CR
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

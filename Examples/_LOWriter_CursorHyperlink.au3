#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $avSettings
	Local $sPath

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_CursorViewCursorGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "This (U)ser-(D)efined-(F)unction was created using Autoit v3©.")
	If @error Then _ERROR($oDoc, "Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the cursor back one character.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_LEFT, 1, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the cursor back ten characters, selecting them.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_LEFT, 10, True)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the hyperlink settings for the text "Autoit v3©", to the website URL of "http://www.autoitscript.com/site/autoit/", and hyperlink name AutoIt.
	_LOWriter_CursorHyperlink($oViewCursor, "http://www.autoitscript.com/site/autoit/", "AutoIt")
	If @error Then _ERROR($oDoc, "Failed to set hyperlink settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOWriter_CursorHyperlink($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The selected text's current hyperlink settings are as follows: " & @CRLF & _
			"The Hyperlink address is: " & $avSettings[0] & @CRLF & _
			"The Hyperlink name is: " & $avSettings[1])

	; Return the cursor back to the end.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_END)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some more text
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "Support for AutoIt can be obtained by E-Mail.")
	If @error Then _ERROR($oDoc, "Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the cursor back one character.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_LEFT, 1, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the cursor back six characters, selecting them.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_LEFT, 6, True)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the hyperlink settings for the text "E-Mail", to the mail-to URL of "mailto:support@autoitscrip.com?subject=Please help". (I intentionally misquoted the email address).
	_LOWriter_CursorHyperlink($oViewCursor, "mailto:support@autoitscrip.com?subject=Please help")
	If @error Then _ERROR($oDoc, "Failed to set hyperlink settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOWriter_CursorHyperlink($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The selected text's current hyperlink settings are as follows: " & @CRLF & _
			"The Hyperlink address is: " & $avSettings[0] & @CRLF & _
			"The Hyperlink name is: " & $avSettings[1])

	; Return the cursor back to the end.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_END)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some more text
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, @CR & "You are currently running this example.")
	If @error Then _ERROR($oDoc, "Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the cursor back one character.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_LEFT, 1, False)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Move the cursor back seven characters, selecting them.
	_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GO_LEFT, 7, True)
	If @error Then _ERROR($oDoc, "Error performing cursor Move. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Convert the path to URL notation.
	$sPath = _LO_PathConvert(@ScriptFullPath, $LO_PATHCONV_OFFICE_RETURN)
	If @error Then _ERROR($oDoc, "Failed to convert path. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the hyperlink settings for the text "example", the the path of the current example.
	_LOWriter_CursorHyperlink($oViewCursor, $sPath)
	If @error Then _ERROR($oDoc, "Failed to set hyperlink settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avSettings = _LOWriter_CursorHyperlink($oViewCursor)
	If @error Then _ERROR($oDoc, "Failed to retrieve the selected text's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The selected text's current hyperlink settings are as follows: " & @CRLF & _
			"The Hyperlink address is: " & $avSettings[0] & @CRLF & _
			"The Hyperlink name is: " & $avSettings[1])

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

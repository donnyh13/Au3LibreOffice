#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oPageStyle, $oHeaderTextCursor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text at the ViewCursor.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "Some text for demonstration purposes." & @CR & @CR)
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Turn Headers on for this page style.
	_LOWriter_PageStyleHeader($oPageStyle, True)
	If @error Then _ERROR($oDoc, "Failed to turn Headers on for this Page Style. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a text cursor for the page style header.
	$oHeaderTextCursor = _LOWriter_DocHeaderGetTextCursor($oPageStyle, True)
	If @error Then _ERROR($oDoc, "Failed to create a TextCursor in the Page Style's Header. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text in the Text Cursor.
	_LOWriter_DocInsertString($oDoc, $oHeaderTextCursor, "Some text in the Page's Header." & @CR)
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

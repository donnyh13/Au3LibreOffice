#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $sChoices
	Local $asItems[5], $asNewItems[3]
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Fill the Array with Strings for our List Items.
	$asItems[0] = "Option 1"
	$asItems[1] = "Option 2"
	$asItems[2] = "Option 3"
	$asItems[3] = "Option 4"
	$asItems[4] = "Option 5"

	; Insert a Input List Field at the View Cursor. Use the Array created above to fill the List, Set the name to "Pick One", Set the selected item
	; to be "Option 3".
	$oField = _LOWriter_FieldInputListInsert($oDoc, $oViewCursor, False, $asItems, "Pick One", "Option 3")
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to modify the Input List Field.")

	; Fill the second new Array
	$asNewItems[0] = "Choice 1"
	$asNewItems[1] = "Choice 3"
	$asNewItems[2] = "Choice 2"

	; Modify the Input List Field settings. Change our input List to contain our new array of options. The New Name is: "Three Choices", and
	; the selected item is "Choice 2"
	_LOWriter_FieldInputListModify($oField, $asNewItems, "Three Choices", "Choice 2")
	If @error Then _ERROR($oDoc, "Failed to modify field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Pick an Option in the Input List and then press ok.")

	; Retrieve current Field settings.
	$avSettings = _LOWriter_FieldInputListModify($oField)
	If @error Then _ERROR($oDoc, "Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; convert the Array into a String.
	For $i = 0 To UBound($avSettings[0]) - 1
		$sChoices &= @CRLF & ($avSettings[0])[$i]
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Field settings are: " & @CRLF & _
			"The Input List's available choices are: " & $sChoices & @CRLF & _
			"The Input List's name is: " & $avSettings[1] & @CRLF & _
			"The currently selected item is: " & $avSettings[2])

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

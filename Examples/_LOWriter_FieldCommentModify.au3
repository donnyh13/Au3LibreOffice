#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $sDateTime
	Local $avSettings, $avDate
	Local $tDateStruct

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text.
	_LOWriter_CursorInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Date Structure, Year = 1844, Month = 10, Day = 22, Hour = 8, minutes = 14, Seconds = 0 , Nanoseconds = 0, UTC= True.
	$tDateStruct = _LOWriter_DateStructCreate(1844, 10, 22, 8, 14, 0, 0, True)
	If @error Then _ERROR($oDoc, "Failed to create a Date structure. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert a Comment Field at the View Cursor. Set content to "This is a note", Author to "Daniel", Date to my previous Date Structure I created.
	; Initials to "D.", Author to "A Name", Resolved = True
	$oField = _LOWriter_FieldCommentInsert($oDoc, $oViewCursor, False, "This is a note", "Daniel", $tDateStruct, "D.", "A Name", True)
	If @error Then _ERROR($oDoc, "Failed to insert a field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press Ok to modify the comment Field.")

	; Create a new Date Structure, leaving all blank will create a Date Structure with today's date.
	$tDateStruct = _LOWriter_DateStructCreate()
	If @error Then _ERROR($oDoc, "Failed to create a Date structure. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Modify the Comment Field settings. Set content to "New Content", Author to "Anon", Date to my new Date Structure I created.
	; Initials to "A.", Author to "A-Non-Mouse", Resolved = False
	_LOWriter_FieldCommentModify($oDoc, $oField, "New Content", "Anon", $tDateStruct, "A.", "A-Non-Mouse", False)
	If @error Then _ERROR($oDoc, "Failed to modify field. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve current Field settings.
	$avSettings = _LOWriter_FieldCommentModify($oDoc, $oField)
	If @error Then _ERROR($oDoc, "Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; convert the Date Struct to an Array, and then into a String.
	$avDate = _LOWriter_DateStructModify($avSettings[2])
	If @error Then _ERROR($oDoc, "Failed to retrieve Date structure properties. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To UBound($avDate) - 1
		If IsBool($avDate[$i]) Then
			If ($avDate[$i] = True) Then
				$sDateTime &= " UTC"

			Else
				; Skip UTC setting
			EndIf

		Else
			$sDateTime &= $avDate[$i] & ":"
		EndIf
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The current Field settings are: " & @CRLF & _
			"The Comment's content is: " & $avSettings[0] & @CRLF & _
			"The Comment's Author is: " & $avSettings[1] & @CRLF & _
			"The Comment's Creation date is: " & $sDateTime & @CRLF & _
			"The Comment's Author's Initials are: " & $avSettings[3] & @CRLF & _
			"The Comment's Author's Name is: " & $avSettings[4] & @CRLF & _
			"Is the Comment resolved? True/False: " & $avSettings[5])

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

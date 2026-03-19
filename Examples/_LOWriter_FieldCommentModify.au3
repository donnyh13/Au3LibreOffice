
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $sDateTime
	Local $avSettings, $avDate
	Local $tDateStruct

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	;Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If (@error > 0) Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	;Create a Date Structure, Year = 1844, Month = 10, Day = 22, Hour = 8, minutes = 14, Seconds = 0 , NanoSeconds = 0, UTC= True.
	$tDateStruct = _LOWriter_DateStructCreate(1844, 10, 22, 8, 14, 0, 0, True)
	If (@error > 0) Then _ERROR("Failed to create a Date structure. Error:" & @error & " Extended:" & @extended)

	;Insert a Comment Field at the View Cursor. Set content to "This is a note", Author to "Daniel", Date to my previouse Date Structure I created.
	;Initials to "D.", Author to "A Name", Resolved = True
	$oField = _LOWriter_FieldCommentInsert($oDoc, $oViewCursor, False, "This is a note", "Daniel", $tDateStruct, "D.", "A Name", True)
	If (@error > 0) Then _ERROR("Failed to insert a field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the comment Field.")

	;Create a new Date Structure, leaving all blank will create a Date Structure with today's date.
	$tDateStruct = _LOWriter_DateStructCreate()
	If (@error > 0) Then _ERROR("Failed to create a Date structure. Error:" & @error & " Extended:" & @extended)

	;Modify the Comment Field settings. Set content to "New Content", Author to "Anon", Date to my new Date Structure I created.
	;Initials to "A.", Author to "A-Non-Mouse", Resolved = False
	_LOWriter_FieldCommentModify($oDoc, $oField, "New Content", "Anon", $tDateStruct, "A.", "A-Non-Mouse", False)
	If (@error > 0) Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Field settings.
	$avSettings = _LOWriter_FieldCommentModify($oDoc, $oField)
	If (@error > 0) Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	;convert the Date Struct to an Array, and then into a String.
	$avDate = _LOWriter_DateStructModify($avSettings[2])
	If (@error > 0) Then _ERROR("Failed to retrieve Date structure properties. Error:" & @error & " Extended:" & @extended)

	For $i = 0 To UBound($avDate) - 1
		If IsBool($avDate[$i]) Then
			If ($avDate[$i] = True) Then
				$sDateTime &= " UTC"
			Else
				;Skip UTC setting
			EndIf
		Else
			$sDateTime &= $avDate[$i] & ":"
		EndIf
	Next

	MsgBox($MB_OK, "", "The current Field settings are: " & @CRLF & _
			"The Comment's content is: " & $avSettings[0] & @CRLF & _
			"The Comment's Author is: " & $avSettings[1] & @CRLF & _
			"The Comment's Creation date is: " & $sDateTime & @CRLF & _
			"The Comment's Author's Initials are: " & $avSettings[3] & @CRLF & _
			"The Comment's Author's Name is: " & $avSettings[4] & @CRLF & _
			"Is the Comment resolved? True/False: " & $avSettings[5])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc




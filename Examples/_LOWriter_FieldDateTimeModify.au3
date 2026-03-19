
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $iDateFormatKey
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

	;Create a Date Structure, leave it blank so it will be set to the current date/Time.
	$tDateStruct = _LOWriter_DateStructCreate()
	If (@error > 0) Then _ERROR("Failed to create a Date structure. Error:" & @error & " Extended:" & @extended)

	;Create or retrieve a DateFormat Key, Hour, Minute, Second, Month Day Year
	$iDateFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "H:MM:SS AM/PM MM/DD/YYYY")
	If (@error > 0) Then _ERROR("Failed to create a Date Format Key. Error:" & @error & " Extended:" & @extended)

	;Insert a Date and Time text Field at the View Cursor., Fixed = False, Set the Date to my previously created DateStruct, Is Date = True,
	;Offset (In Days since I set Date to True) = -1 meaning minus one day, and set DateTime Format Key to the Key I just created.
	$oField = _LOWriter_FieldDateTimeInsert($oDoc, $oViewCursor, False, False, $tDateStruct, True, -1, $iDateFormatKey)
	If (@error > 0) Then _ERROR("Failed to insert a field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Date/Time Field.")

	;Create a new Date Structure, Year = 1992, Month = 4, Day = 28, Hour = 12, Minute == 00 , Sec == 00.
	$tDateStruct = _LOWriter_DateStructCreate(1992, 04, 28, 12, 00, 00)
	If (@error > 0) Then _ERROR("Failed to create a Date structure. Error:" & @error & " Extended:" & @extended)

	;Create or retrieve a DateFormat Key,  Month Day Year Hour, Minute, Second, AM/PM
	$iDateFormatKey = _LOWriter_DateFormatKeyCreate($oDoc, "MM/DD/YYYY H:MM:SS AM/PM")
	If (@error > 0) Then _ERROR("Failed to create a Date Format Key. Error:" & @error & " Extended:" & @extended)

	;Modify the Date/Time Field settings. Fixed= True, Modify the Date to the one just created, Is Date = False, Off set (in minutes) = 20, Use my new
	;DateFormat Key.
	_LOWriter_FieldDateTimeModify($oDoc, $oField, True, $tDateStruct, False, 20, $iDateFormatKey)
	If (@error > 0) Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	;Retrieve current Field settings. Return will be an Array with elements in the order of function parameters.
	$avSettings = _LOWriter_FieldDateTimeModify($oDoc, $oField)
	If (@error > 0) Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	;convert the Date Struct to an Array, and then into a String.
	$avDate = _LOWriter_DateStructModify($avSettings[1])
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
			"Is the Date/Time Fixed at that time? Ture/False: " & $avSettings[0] & @CRLF & _
			"The Date/Time Field is set to the current Date and Time: " & $sDateTime & @CRLF & _
			"Is this set as a Date, and not a time? True/False: " & $avSettings[2] & @CRLF & _
			"The Offset is set to: " & $avSettings[3] & @CRLF & _
			"The DateTime Format Key used is: " & $avSettings[4] & " Which looks like: " & _LOWriter_DateFormatKeyGetString($oDoc, $avSettings[4]))

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc





#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $sDateTime
	Local $avReturn, $avDate
	Local $tDateStruct

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Create a Date Structure, Year = 1844, Month = 10, Day = 22, Hour = 8, minutes = 14, Seconds = 0 , NanoSeconds = 0, UTC= True.
	$tDateStruct = _LOWriter_DateStructCreate(1844, 10, 22, 8, 14, 0, 0, True)
	If (@error > 0) Then _ERROR("Failed to create a Date structure. Error:" & @error & " Extended:" & @extended)

	;Set the Document's template settings to, Template name = "AutoIt", Template URL to a fake Path, C:\Folder1\Folder2\Test.ott
	;Date to the previously created Day Structure.
	_LOWriter_DocGenPropTemplate($oDoc, "AutoIt", "C:\Folder1\Folder2\Test.ott", $tDateStruct)
	If (@error > 0) Then _ERROR("Failed to modify Document settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Document's description. Return will be an Array in order of function parameters.
	$avReturn = _LOWriter_DocGenPropTemplate($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve document information. Error:" & @error & " Extended:" & @extended)

	;convert the Date Struct to an Array, and then into a String.
	$avDate = _LOWriter_DateStructModify($avReturn[2])
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

	MsgBox($MB_OK, "", "The document was created using a template named: " & $avReturn[0] & @CRLF & _
			"At the following location: " & $avReturn[1] & @CRLF & _
			"The document was created from this template at the following Date and Time: " & $sDateTime)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc


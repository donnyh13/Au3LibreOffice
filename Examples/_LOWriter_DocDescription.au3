
#include "LibreOfficeWriter.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc
	Local $sKeywords
	Local $asKeywords[3] = ["Autoit", "LibreOffice", "Writer"]
	Local $avReturn, $asReturnedKeywords

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Set the Document's Description settings, Title = "Autoit Example", Subject = "Doc Description Demonstration", Keywords to Keywords Array,
	;Comments To two lines of comments.
	_LOWriter_DocDescription($oDoc, "Autoit Example", "Doc Description Demonstration", $asKeywords, "This is a comment." & @CR & "This is a second comment line.")
	If (@error > 0) Then _ERROR("Failed to modify Document settings. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Document's description. Return will be an Array in order of function parameters.
	$avReturn = _LOWriter_DocDescription($oDoc)
	If (@error > 0) Then _ERROR("Failed to retrieve document information. Error:" & @error & " Extended:" & @extended)

	$asReturnedKeywords = $avReturn[2]
	;Convert the Keyword Array to a String, separate each element with a @CRLF
	For $i = 0 To UBound($asReturnedKeywords) - 1
		$sKeywords = $sKeywords & $asReturnedKeywords[$i] & @CRLF
	Next

	MsgBox($MB_OK, "", "The document's description properties are: " & @CRLF & _
			"The Document's title is: " & $avReturn[0] & @CRLF & _
			"The Document's Subject is: " & $avReturn[1] & @CRLF & _
			"The Keywords for this document are: " & @CRLF & $sKeywords & _
			"The Comments for this document are: " & $avReturn[3])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc


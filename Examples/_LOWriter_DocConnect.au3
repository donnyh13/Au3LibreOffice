#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oDoc2, $oDoc3
	Local $iUserChoice
	Local $sDocName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, $oDoc2, $oDoc3, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Connect to the Current Document.
	$oDoc2 = _LOWriter_DocConnect($LO_DOC_CONNECT_MODE_CURRENT)
	If (@error > 0) Or Not IsObj($oDoc2) Then _ERROR($oDoc, $oDoc2, $oDoc3, "Failed to Connect to Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Doc Name.
	$sDocName = _LOWriter_DocGetName($oDoc2, False)
	If @error Then _ERROR($oDoc, $oDoc2, $oDoc3, "Failed to retrieve Writer Document name. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$iUserChoice = MsgBox($MB_YESNO, "Close?", "I have connected to the current Document, which has the following title: " & $sDocName & @CRLF & @CRLF & _
			"Would you like to connect again to the Document using this same name and close it?")

	If ($iUserChoice = $IDYES) Then
		; Connect to the document.
		$oDoc3 = _LOWriter_DocConnect($LO_DOC_CONNECT_MODE_SEARCH_NAME, $sDocName)
		If (@error > 0) Or Not IsObj($oDoc3) Then _ERROR($oDoc, $oDoc2, $oDoc3, "Failed to Connect to Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

		; Close the document, don't save changes.
		_LOWriter_DocClose($oDoc3, False)
		If @error Then _ERROR($oDoc, $oDoc2, $oDoc3, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	EndIf

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, $oDoc2, $oDoc3, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $oDoc2, $oDoc3, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	If IsObj($oDoc2) Then _LOWriter_DocClose($oDoc2, False)
	If IsObj($oDoc3) Then _LOWriter_DocClose($oDoc3, False)
	Exit
EndFunc

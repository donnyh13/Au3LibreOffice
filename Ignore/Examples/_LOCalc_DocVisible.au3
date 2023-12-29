#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc
	Local $bReturn

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok and I will make the new document I just opened invisible.")

	; Make the document invisible by setting visible to False
	_LOCalc_DocVisible($oDoc, False)
	If (@error > 0) Then
		_LOCalc_DocClose($oDoc, False)
		_ERROR("Failed to change Document visibility settings. Error:" & @error & " Extended:" & @extended)
	EndIf

	; Test if the document is Visible
	$bReturn = _LOCalc_DocVisible($oDoc)
	If @error Then _ERROR("Failed to retrieve document status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is the document currently visible? True/False: " & $bReturn & @CRLF & @CRLF & _
			"Press Ok to make the document visible again.")

	; Make the document visible by setting visible to True
	_LOCalc_DocVisible($oDoc, True)
	If (@error > 0) Then
		_LOCalc_DocClose($oDoc, False)
		_ERROR("Failed to change Document visibility settings. Error:" & @error & " Extended:" & @extended)
	EndIf

	; Test if the document is Visible
	$bReturn = _LOCalc_DocVisible($oDoc)
	If @error Then _ERROR("Failed to retrieve document status. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Is the document now visible? True/False: " & $bReturn)

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

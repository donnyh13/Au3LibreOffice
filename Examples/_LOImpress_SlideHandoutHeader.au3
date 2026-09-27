#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $oDoc, $oHandout
	Local $avSettings

	; Create a New, visible, Blank LibreOffice Document.
	$oDoc = _LOImpress_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Impress Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the handout Page's Object.
	$oHandout = _LOImpress_SlideHandoutGetObj($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Handout page's Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set some of the Handout Page's Header settings
	_LOImpress_SlideHandoutHeader($oHandout, True, "Hello, from AutoIt!", True, False, Null, $LOI_FIELD_DATE_FMT_DOW_MMM_DD_YYYY)
	If @error Then _ERROR($oDoc, "Failed to set Page's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Handout Page's Header settings. Return will be an array in order of function parameters.
	$avSettings = _LOImpress_SlideHandoutHeader($oHandout)
	If @error Then _ERROR($oDoc, "Failed to retrieve Page's settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Handout Page's Header settings are as follows: " & @CRLF & _
			"Is the Header field enabled? True/False: " & $avSettings[0] & @CRLF & _
			"The Header text to display is (if any): " & $avSettings[1] & @CRLF & _
			"Is the Date/Time field enabled? True/False: " & $avSettings[2] & @CRLF & _
			"Is the Date/Time fixed? True/False: " & $avSettings[3] & @CRLF & _
			"If the Date/Time is fixed, the value to display is: " & $avSettings[4] & @CRLF & _
			"The format to display the Date/Time in is (See UDF Constants): " & $avSettings[5])

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOImpress_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Close the background LibreOffice instance if all Documents are closed.
	_LO_Terminate()
	If @error Then Return _ERROR($oDoc, "Failed to Terminate LibreOffice. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOImpress_DocClose($oDoc, False)
	Exit
EndFunc

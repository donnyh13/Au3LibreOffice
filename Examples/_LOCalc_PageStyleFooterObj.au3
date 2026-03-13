#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle, $oFooter, $oTextCursor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default page Style Object.
	$oPageStyle = _LOCalc_PageStyleGetObj($oDoc, "Default")
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve Right Page Page Style Footer Object.
	$oFooter = _LOCalc_PageStyleFooterObj($oPageStyle, Null, Default)
	If @error Then _ERROR($oDoc, "Failed to retrieve Footer Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Text Cursor in Right Page Footer, in the "Center" Area.
	$oTextCursor = _LOCalc_PageStyleFooterCreateTextCursor($oFooter, True, False, True)
	If @error Then _ERROR($oDoc, "Failed to create a Text Cursor Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Insert some text into the Footer Center Area
	_LOCalc_TextCursorInsertString($oTextCursor, " This is some text!")
	If @error Then _ERROR($oDoc, "Failed to insert text. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Select the word "Text!" and set it to Bold.
	_LOCalc_TextCursorMove($oTextCursor, $LOC_TEXTCUR_GO_LEFT, 5, True)
	If @error Then _ERROR($oDoc, "Failed to move text cursor. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Set the selected text to Bold.
	_LOCalc_TextCursorFont($oTextCursor, Null, Null, Null, $LOC_CHAR_WEIGHT_BOLD)
	If @error Then _ERROR($oDoc, "Failed to set Text format. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Apply the modified Footer Object to the Page Style, Apply it to both First Page and Right Page, which will apply it to Both Left and Right because Same content on Left and Right option is True.
	_LOCalc_PageStyleFooterObj($oPageStyle, $oFooter, $oFooter)
	If @error Then _ERROR($oDoc, "Failed to set Footer Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	MsgBox($MB_OK + $MB_TOPMOST, Default, "To see the new text, switch to Page Styles in the UI, go to Modify the ""Default"" page style, select Footer, select Edit, and look at ""Footer (First)"" and ""Footer(Rest)"" tabs." & @CRLF & _
			"Press ok to close the document.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOCalc_DocClose($oDoc, False)
	Exit
EndFunc

#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $avPageStyleSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the Default Page Style's Object, to modify its settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Standard")
	If @error Then _ERROR($oDoc, "Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; If Libre Office version is higher or equal to 7.2 then set Gutter settings.
	If (_LO_VersionGet(True) >= 7.2) Then
		; Set Page layout to, $LOW_PAGE_LAYOUT_MIRRORED, Numbering format to $LOW_NUM_STYLE_CHARS_UPPER_LETTER_N, Reference Paragraph style to
		; "Standard", Gutter on Right to False, Gutter At top to False, Background covers margins to True.
		_LOWriter_PageStyleLayout($oDoc, $oPageStyle, $LOW_PAGE_LAYOUT_MIRRORED, $LOW_NUM_STYLE_CHARS_UPPER_LETTER_N, "Standard", False, False, True)
		If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	Else ; Set all other settings, except the Gutter settings.
		; Set layout to, $LOW_PAGE_LAYOUT_MIRRORED, Numbering format to $LOW_NUM_STYLE_CHARS_UPPER_LETTER_N, Reference Paragraph style to
		; "Standard", Background covers margins to True.
		_LOWriter_PageStyleLayout($oDoc, $oPageStyle, $LOW_PAGE_LAYOUT_MIRRORED, $LOW_NUM_STYLE_CHARS_UPPER_LETTER_N, "Standard", Null, Null, True)
		If @error Then _ERROR($oDoc, "Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
	EndIf

	; Retrieve the current settings. Return will be an array with element values in order of function parameters.
	$avPageStyleSettings = _LOWriter_PageStyleLayout($oDoc, $oPageStyle)
	If @error Then _ERROR($oDoc, "Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; If Libre Office version is higher or equal to 7.2 then display the Gutter margin setting.
	If (_LO_VersionGet(True) >= 7.2) Then
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Layout settings are as follows: " & @CRLF & _
				"The current Page Layout is, (see UDF constants): " & $avPageStyleSettings[0] & @CRLF & _
				"The Numbering format used is, (See UDF constants): " & $avPageStyleSettings[1] & @CRLF & _
				"The Reference Paragraph Style name is: " & $avPageStyleSettings[2] & @CRLF & _
				"Gutter is on the right? True/False: " & $avPageStyleSettings[3] & @CRLF & _
				"Gutter is on the top? True/False: " & $avPageStyleSettings[4] & @CRLF & _
				"Background covers the margins? True/False: " & $avPageStyleSettings[5] & @CRLF & _
				"The paper tray to use, when printing this document is: " & $avPageStyleSettings[6])

	Else ; Display all other margin settings, except the Gutter margin.
		MsgBox($MB_OK + $MB_TOPMOST, Default, "The Page Style's current Layout settings are as follows: " & @CRLF & _
				"The current Page Layout is, (see UDF constants): " & $avPageStyleSettings[0] & @CRLF & _
				"The Numbering format used is, (See UDF constants): " & $avPageStyleSettings[1] & @CRLF & _
				"The Reference Paragraph Style name is: " & $avPageStyleSettings[2] & @CRLF & _
				"Background covers the margins? True/False: " & $avPageStyleSettings[3] & @CRLF & _
				"The paper tray to use, when printing this document is: " & $avPageStyleSettings[4])
	EndIf

	MsgBox($MB_OK + $MB_TOPMOST, Default, "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)
EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOWriter_DocClose($oDoc, False)
	Exit
EndFunc

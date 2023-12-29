
#include "LibreOfficeWriter.au3"
#include "LibreOfficeWriterConstants.au3"
#include <MsgBoxConstants.au3>

Example()

Func Example()
	Local $oDoc, $oPageStyle
	Local $avPageStyleSettings

	;Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If (@error > 0) Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	;Retrieve the Default Page Style's Object, to modify it's settings.
	$oPageStyle = _LOWriter_PageStyleGetObj($oDoc, "Default Page Style")
	If (@error > 0) Then _ERROR("Failed to retrieve Page Style Object. Error:" & @error & " Extended:" & @extended)

	;If Libre Office version is higher or equal to 7.2 then set Gutter settings.
	If (_LOWriter_VersionGet(True) >= 7.2) Then

		;Set Page layout to, $LOW_PAGE_LAYOUT_MIRRORED, Numbering format to $LOW_NUM_STYLE_CHARS_UPPER_LETTER_N, Reference Paragraph style to
		;"Default Paragraph Style", Gutter on Right to False, Gutter At top to False, Background covers margins to True.
		_LOWriter_PageStyleLayout($oDoc, $oPageStyle, $LOW_PAGE_LAYOUT_MIRRORED, $LOW_NUM_STYLE_CHARS_UPPER_LETTER_N, "Default Paragraph Style", False, False, True)
		If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)

	Else ;Set all other settins, except the Gutter settings.
		;Set layout to, $LOW_PAGE_LAYOUT_MIRRORED, Numbering format to $LOW_NUM_STYLE_CHARS_UPPER_LETTER_N, Reference Paragraph style to
		;"Default Paragraph Style", Background covers margins to True.
		_LOWriter_PageStyleLayout($oDoc, $oPageStyle, $LOW_PAGE_LAYOUT_MIRRORED, $LOW_NUM_STYLE_CHARS_UPPER_LETTER_N, "Default Paragraph Style", Null, Null, True)
		If (@error > 0) Then _ERROR("Failed to modify Page Style settings. Error:" & @error & " Extended:" & @extended)
	EndIf

	;Retrieve the current settings. Return will be an array with element values in order of function parameter.
	$avPageStyleSettings = _LOWriter_PageStyleLayout($oDoc, $oPageStyle)
	If (@error > 0) Then _ERROR("Failed to retrieve the Page style settings. Error:" & @error & " Extended:" & @extended)

	;If Libre Office version is higher or equal to 7.2 then display the Gutter margin setting.
	If (_LOWriter_VersionGet(True) >= 7.2) Then
		MsgBox($MB_OK, "", "The Page Style's current Layout settings are as follows: " & @CRLF & _
				"The current Page Layout is, (see UDF constants): " & $avPageStyleSettings[0] & @CRLF & _
				"The Numbering format used is, (See UDF constants): " & $avPageStyleSettings[1] & @CRLF & _
				"The Reference Paragraph Style name is: " & $avPageStyleSettings[2] & @CRLF & _
				"Gutter is on the right? True/False: " & $avPageStyleSettings[3] & @CRLF & _
				"Gutter is on the top? True/False: " & $avPageStyleSettings[4] & @CRLF & _
				"Background covers the margins? True/False: " & $avPageStyleSettings[5] & @CRLF & _
				"The paper tray to use, when printing this document is: " & $avPageStyleSettings[6])

	Else ; Display all other margin settings, except the Gutter margin.
		MsgBox($MB_OK, "", "The Page Style's current Layout settings are as follows: " & @CRLF & _
				"The current Page Layout is, (see UDF constants): " & $avPageStyleSettings[0] & @CRLF & _
				"The Numbering format used is, (See UDF constants): " & $avPageStyleSettings[1] & @CRLF & _
				"The Reference Paragraph Style name is: " & $avPageStyleSettings[2] & @CRLF & _
				"Background covers the margins? True/False: " & $avPageStyleSettings[3] & @CRLF & _
				"The paper tray to use, when printing this document is: " & $avPageStyleSettings[4])
	EndIf

	MsgBox($MB_OK, "", "Press ok to close the document.")

	;Close the document.
	_LOWriter_DocClose($oDoc, False)
	If (@error > 0) Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc




#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor
	Local $asTableStyles, $asTableStylesDisplay
	Local $sStyles = ""

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR($oDoc, "Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Create a Table, 3 columns, 5 rows.
	_LOWriter_TableCreate($oDoc, $oViewCursor, 3, 5, Null, Null, Null, "Elegant")
	If @error Then _ERROR($oDoc, "Failed to create Text Table. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all available Table Style Names.
	$asTableStyles = _LOWriter_TableStylesGetNames($oDoc)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table style list. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all available Table Style display Names.
	$asTableStylesDisplay = _LOWriter_TableStylesGetNames($oDoc, False, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table style list. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	For $i = 0 To (UBound($asTableStyles) - 1)
		If ($asTableStyles[$i] <> $asTableStylesDisplay[$i]) Then
			$sStyles &= $asTableStyles[$i] & @CRLF & "(Display Name: " & $asTableStylesDisplay[$i] & ")" & @CRLF & @CRLF

		Else
			$sStyles &= $asTableStyles[$i] & @CRLF & @CRLF
		EndIf
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The available Table Styles are:" & @CRLF & $sStyles)

	; Retrieve an Array of all Table Styles used in the Document.
	$asTableStyles = _LOWriter_TableStylesGetNames($oDoc, False, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table style list. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	; Retrieve an Array of all Table Styles used in the Document.
	$asTableStylesDisplay = _LOWriter_TableStylesGetNames($oDoc, False, True, True)
	If @error Then _ERROR($oDoc, "Failed to retrieve Table style list. Error:" & @error & " Extended:" & @extended & " On Line: " & @ScriptLineNumber)

	$sStyles = ""

	For $i = 0 To (UBound($asTableStyles) - 1)
		If ($asTableStyles[$i] <> $asTableStylesDisplay[$i]) Then
			$sStyles &= $asTableStyles[$i] & @CRLF & "(Display Name: " & $asTableStylesDisplay[$i] & ")" & @CRLF & @CRLF

		Else
			$sStyles &= $asTableStyles[$i] & @CRLF & @CRLF
		EndIf
	Next

	MsgBox($MB_OK + $MB_TOPMOST, Default, "The Table Styles used in this document are:" & @CRLF & $sStyles)

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

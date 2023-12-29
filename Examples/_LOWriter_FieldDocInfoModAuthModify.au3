#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $avSettings
	Local $tDateStruct

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Create a Date Structure, Year = 1844, Month = 10, Day = 22, Hour = 8, minutes = 14, Seconds = 0 , NanoSeconds = 0, UTC= True.
	$tDateStruct = _LOWriter_DateStructCreate(1844, 10, 22, 8, 14, 0, 0, True)
	If @error Then _ERROR("Failed to create a Date structure. Error:" & @error & " Extended:" & @extended)

	; Set the Document's Modification settings to, Modified by = "Daniel", Date to the previously created Day Structure, for this Field demonstration.
	_LOWriter_DocGenPropModification($oDoc, "Daniel", $tDateStruct)
	If @error Then _ERROR("Failed to modify Document settings. Error:" & @error & " Extended:" & @extended)

	; Insert a Document Information Modification Author Field at the View Cursor. Set is Fixed = False
	$oField = _LOWriter_FieldDocInfoModAuthInsert($oDoc, $oViewCursor, False, False)
	If @error Then _ERROR("Failed to insert a Field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the Doc Info Field settings.")

	; Modify the Doc Info Modification Author Field settings. Set Fixed to True, Set author to "Me".
	_LOWriter_FieldDocInfoModAuthModify($oField, True, "Me")
	If @error Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	; Retrieve current Field settings. Return will be an array, with elements in order of function parameters.
	$avSettings = _LOWriter_FieldDocInfoModAuthModify($oField)
	If @error Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Doc Info Field settings are: " & @CRLF & _
			"Is the content of this field fixed? True/ False: " & $avSettings[0] & @CRLF & _
			"The last person to modify this document was: " & $avSettings[1])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

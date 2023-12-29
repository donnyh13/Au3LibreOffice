#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeCalc.au3"

Example()

Func Example()
	Local $oDoc, $oSheet, $oCell
	Local $sFilePathName, $sPath

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOCalc_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Calc Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I will now save the new Calc Document to the desktop folder.")

	$sFilePathName = _TempFile(@DesktopDir & "\", "TestExportDoc_", ".ods")

	; save The New Blank Doc To Desktop Directory using a unique temporary name.
	$sPath = _LOCalc_DocSaveAs($oDoc, $sFilePathName)
	If @error Then _ERROR("Failed to Save the Calc Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have created and saved the document to your Desktop, found at the following Path: " _
			& $sPath & @CRLF & "Press Ok to write some data to it and then save the changes and close the document.")

	; Retrieve the presently active Sheet.
	$oSheet = _LOCalc_SheetGetActive($oDoc)
	If @error Then _ERROR("Failed to retrieve the currently active Sheet Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the top left most cell, 0, 0, or A1.
	$oCell = _LOCalc_SheetGetCellByPosition($oSheet, 0, 0)
	If @error Then _ERROR("Failed to retrieve A1 Cell Object. Error:" & @error & " Extended:" & @extended)

	; Set A1 Cell text to "A1"
	_LOCalc_CellText($oCell, "A1")
	If @error Then _ERROR("Failed to Set A1 Cell content. Error:" & @error & " Extended:" & @extended)

	; Save the changes to the document.
	_LOCalc_DocSave($oDoc)
	If @error Then _ERROR("Failed to Save the Calc Document. Error:" & @error & " Extended:" & @extended)

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have written and saved the data, and closed the document. I will now open it up again to show it worked.")

	; Open the document.
	$oDoc = _LOCalc_DocOpen($sPath)
	If @error Then _ERROR("Failed to open Calc Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Document was succesfully opened. Press OK to close and delete it.")

	; Close the document.
	_LOCalc_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Following Error codes returned: Error:" & _
			@error & " Extended:" & @extended)

	; Delete the file.
	FileDelete($sPath)
EndFunc   ;==>Example

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc   ;==>_ERROR

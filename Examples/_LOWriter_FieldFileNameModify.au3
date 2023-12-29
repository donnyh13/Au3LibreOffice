#include <File.au3>
#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor, $oField
	Local $sSavePath, $sPath
	Local $avSettings

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)
	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "ATemporaryDoc_", ".odt")

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOWriter_DocSaveAs($oDoc, $sSavePath, "", True)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text.
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "I have inserted a field at the end of this line.--> ")
	If @error Then _ERROR("Failed to insert text. Error:" & @error & " Extended:" & @extended)

	; Insert a Sender Field at the View Cursor. Fixed =False, format = $LOW_FIELD_FILENAME_NAME
	$oField = _LOWriter_FieldFileNameInsert($oDoc, $oViewCursor, False, False, $LOW_FIELD_FILENAME_NAME)
	If @error Then _ERROR("Failed to insert a field. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "Press Ok to modify the File Name Field.")

	; Modify the File Name Field settings. Skip Fixed, File Name Format = $LOW_FIELD_FILENAME_FULL_PATH
	_LOWriter_FieldFileNameModify($oField, Null, $LOW_FIELD_FILENAME_FULL_PATH)
	If @error Then _ERROR("Failed to modfiy field. Error:" & @error & " Extended:" & @extended)

	; Retrieve current Field settings. Return will be an Array with elements in the order of function parameters.
	$avSettings = _LOWriter_FieldFileNameModify($oField)
	If @error Then _ERROR("Failed to retrieve field settings. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The current Field settings are: " & @CRLF & _
			"Is the File Name Field's content fixed? True/False: " & $avSettings[0] & @CRLF & _
			"The File Name Field's display format is, (see UDF COnstants): " & $avSettings[1])

	MsgBox($MB_OK, "", "Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

	; Delete the file.
	FileDelete($sPath)
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

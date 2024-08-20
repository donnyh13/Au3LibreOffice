#include <MsgBoxConstants.au3>
#include <File.au3>

#include "..\LibreOfficeBase.au3"

Global $sPath

Example()

; Delete the file.
If IsString($sPath) Then FileDelete($sPath)

Func Example()
	Local $oDBase, $oDoc
	Local $sSavePath, $sName

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOBase_DocCreate(True, False)
	If @error Then Return _ERROR($oDoc, "Failed to Create a new Base Document. Error:" & @error & " Extended:" & @extended)

	; Create a unique file name
	$sSavePath = _TempFile(@TempDir & "\", "DocTestFile_", ".odb")

	; Set the Database type.
	_LOBase_DocDatabaseType($oDoc)
	If @error Then Return _ERROR($oDoc, "Failed to Set Base Document Database type. Error:" & @error & " Extended:" & @extended)

	; Save The New Blank Doc To Temp Directory.
	$sPath = _LOBase_DocSaveAs($oDoc, $sSavePath, True)
	If @error Then Return _ERROR($oDoc, "Failed to save the Base Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Database Object using the Save path.
	$oDBase = _LOBase_DatabaseGetObjByURL($sPath)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Base Document Database Object. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Database Name.
	$sName = _LOBase_DatabaseName($oDBase)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Database Name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have retrieved the Database Object for the new document I Created and saved, the Database Name is (which will be the save path, as it is not registered): " & @CRLF & _
			$sName & @CRLF & @CRLF & "Press ok to register the Database.")

	; Register the Database
	_LOBase_DatabaseRegisteredAdd($oDBase, "AutoIt_Database")
	If @error Then Return _ERROR($oDoc, "Failed to Register the Database. Error:" & @error & " Extended:" & @extended)

	; Retrieve the Database Name.
	$sName = _LOBase_DatabaseName($oDBase)
	If @error Then Return _ERROR($oDoc, "Failed to Retrieve the Database Name. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "The Database's Name is: " & @CRLF & _
			$sName & @CRLF & @CRLF & "Press ok to unregister the Database and delete the file.")

	; Unregister the Database
	_LOBase_DatabaseRegisteredRemoveByName("AutoIt_Database")
	If @error Then Return _ERROR($oDoc, "Failed to Unregister the Database. Error:" & @error & " Extended:" & @extended)

	; Close the document.
	_LOBase_DocClose($oDoc, False)
	If @error Then Return _ERROR($oDoc, "Failed to close opened L.O. Document. Error:" & @error & " Extended:" & @extended)

EndFunc

Func _ERROR($oDoc, $sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	If IsObj($oDoc) Then _LOBase_DocClose($oDoc, False)
EndFunc
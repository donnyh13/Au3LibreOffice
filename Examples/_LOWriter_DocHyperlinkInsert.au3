#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $oDoc, $oViewCursor

	; Create a New, visible, Blank Libre Office Document.
	$oDoc = _LOWriter_DocCreate(True, False)

	If @error Then _ERROR("Failed to Create a new Writer Document. Error:" & @error & " Extended:" & @extended)

	; Retrieve the document view cursor to insert text with.
	$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
	If @error Then _ERROR("Failed to retrieve the View Cursor Object for the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert some text
	_LOWriter_DocInsertString($oDoc, $oViewCursor, "This (U)ser-(D)efined-(F)unction was created using ")
	If @error Then _ERROR("Failed to insert text into the Writer Document. Error:" & @error & " Extended:" & @extended)

	; Insert a hyperlink. Set the display text to "Autoit v3©.", and the website URL to "http://www.autoitscript.com/site/autoit/"
	_LOWriter_DocHyperlinkInsert($oDoc, $oViewCursor, "Autoit v3©.", "http://www.autoitscript.com/site/autoit/")
	If @error Then _ERROR("Failed to insert a hyperlink into the Writer Document. Error:" & @error & " Extended:" & @extended)

	MsgBox($MB_OK, "", "I have inserted a hyperlink into the document. Press ok to close the document.")

	; Close the document.
	_LOWriter_DocClose($oDoc, False)
	If @error Then _ERROR("Failed to close opened L.O. Document. Following Error codes returned: Error:" & _
			@error & " Extended:" & @extended)

EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK, "Error", $sErrorText)
	Exit
EndFunc

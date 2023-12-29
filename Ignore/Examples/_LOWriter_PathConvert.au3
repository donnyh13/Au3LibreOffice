#include <MsgBoxConstants.au3>

#include "..\LibreOfficeWriter.au3"

Example()

Func Example()
	Local $sPCPath, $sOfficePath

	; An example of a computer path
	$sPCPath = "C:\A Folder With Spaces\FolderWithASemicolon;\{FolderWithBrackets}\TestDocument.odt"

	; An example of a Libre Office URL
	$sOfficePath = "file:///C:/A%20Folder%20With%20Spaces/FolderWithASemicolon%3B/%7BFolderWithBrackets%7D/TestDocument.odt"

	; AutoReturn From a Computer Path
	MsgBox($MB_OK, "Auto_Return -- Computer Path", "This is the result from converting a Computer path to Libre Office URL automatically." & @CRLF & _
			"Original Path: " & @CRLF & $sPCPath & @CRLF & "Result: " & @CRLF & _
			_LOWriter_PathConvert($sPCPath, $LOW_PATHCONV_AUTO_RETURN))

	; AutoReturn From a Libre Office URL
	MsgBox($MB_OK, "Auto_Return -- Libre Office URL", "This is the result from converting a Libre Office URL to Computer Path automatically." & @CRLF & _
			"Original Path: " & @CRLF & $sOfficePath & @CRLF & "Result: " & @CRLF & _
			_LOWriter_PathConvert($sOfficePath, $LOW_PATHCONV_AUTO_RETURN))

	; Return From a Libre Office URL to Computer Path conversion
	MsgBox($MB_OK, "PCPATH_RETURN -- Libre Office URL", "This is the result from converting a Libre Office URL to Computer Path." & @CRLF & _
			"Original Path: " & @CRLF & $sOfficePath & @CRLF & "Result: " & @CRLF & _
			_LOWriter_PathConvert($sOfficePath, $LOW_PATHCONV_PCPATH_RETURN))

	; Return From a Libre Office URL to Computer Path conversion when the path is already a computer path.
	MsgBox($MB_OK, "PCPATH_RETURN -- Computer Path", "This is the result from converting a Libre Office URL to Computer Path " & _
			"when the path is already a computer path." & @CRLF & _
			"Original Path: " & @CRLF & $sPCPath & @CRLF & "Result: " & @CRLF & _
			_LOWriter_PathConvert($sPCPath, $LOW_PATHCONV_PCPATH_RETURN))

	; Return From a Computer Path to Libre Office URL conversion
	MsgBox($MB_OK, "OFFICE_RETURN -- Computer Path", "This is the result from converting a Computer Path to Libre Office URL." & @CRLF & _
			"Original Path: " & @CRLF & $sPCPath & @CRLF & "Result: " & @CRLF & _
			_LOWriter_PathConvert($sPCPath, $LOW_PATHCONV_OFFICE_RETURN))

	; Return From a Computer Path to Libre Office URL conversion when the path is already a Libre Office path.
	MsgBox($MB_OK, "OFFICE_RETURN -- Libre Office URL", "This is the result from converting a Computer Path to Libre Office URL " & _
			"when the path is already a Libre Office URL." & @CRLF & _
			"Original Path: " & @CRLF & $sOfficePath & @CRLF & "Result: " & @CRLF & _
			_LOWriter_PathConvert($sOfficePath, $LOW_PATHCONV_OFFICE_RETURN))

EndFunc

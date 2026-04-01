#include <MsgBoxConstants.au3>

#include "..\LibreOffice_Helper.au3"

Example()

Func Example()
	Local $sPCPath, $sOfficePath

	; An example of a computer path
	$sPCPath = "C:\A Folder With Spaces\FolderWithASemicolon;\{FolderWithBrackets}\TestDocument.odt"

	; An example of a LibreOffice URL
	$sOfficePath = "file:///C:/A%20Folder%20With%20Spaces/FolderWithASemicolon%3B/%7BFolderWithBrackets%7D/TestDocument.odt"

	; AutoReturn From a Computer Path
	MsgBox($MB_OK + $MB_TOPMOST, Default, "Auto_Return -- Computer Path" & @CRLF & "This is the result from converting a Computer path to LibreOffice URL automatically." & @CRLF & @CRLF & _
			"Original Path: " & @CRLF & $sPCPath & @CRLF & @CRLF & "Result: " & @CRLF & _
			_LO_PathConvert($sPCPath, $LO_PATHCONV_AUTO_RETURN))

	; AutoReturn From a LibreOffice URL
	MsgBox($MB_OK + $MB_TOPMOST, Default, "Auto_Return -- LibreOffice URL" & @CRLF & "This is the result from converting a LibreOffice URL to Computer Path automatically." & @CRLF & @CRLF & _
			"Original Path: " & @CRLF & $sOfficePath & @CRLF & @CRLF & "Result: " & @CRLF & _
			_LO_PathConvert($sOfficePath, $LO_PATHCONV_AUTO_RETURN))

	; Return From a LibreOffice URL to Computer Path conversion
	MsgBox($MB_OK + $MB_TOPMOST, Default, "PCPATH_RETURN -- LibreOffice URL" & @CRLF & "This is the result from converting a LibreOffice URL to Computer Path." & @CRLF & @CRLF & _
			"Original Path: " & @CRLF & $sOfficePath & @CRLF & @CRLF & "Result: " & @CRLF & _
			_LO_PathConvert($sOfficePath, $LO_PATHCONV_PCPATH_RETURN))

	; Return From a LibreOffice URL to Computer Path conversion when the path is already a computer path.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "PCPATH_RETURN -- Computer Path" & @CRLF & "This is the result from converting a LibreOffice URL to Computer Path " & _
			"when the path is already a computer path." & @CRLF & @CRLF & _
			"Original Path: " & @CRLF & $sPCPath & @CRLF & @CRLF & "Result: " & @CRLF & _
			_LO_PathConvert($sPCPath, $LO_PATHCONV_PCPATH_RETURN))

	; Return From a Computer Path to LibreOffice URL conversion
	MsgBox($MB_OK + $MB_TOPMOST, Default, "OFFICE_RETURN -- Computer Path" & @CRLF & "This is the result from converting a Computer Path to LibreOffice URL." & @CRLF & @CRLF & _
			"Original Path: " & @CRLF & $sPCPath & @CRLF & @CRLF & "Result: " & @CRLF & _
			_LO_PathConvert($sPCPath, $LO_PATHCONV_OFFICE_RETURN))

	; Return From a Computer Path to LibreOffice URL conversion when the path is already a LibreOffice path.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "OFFICE_RETURN -- LibreOffice URL" & @CRLF & "This is the result from converting a Computer Path to LibreOffice URL " & _
			"when the path is already a LibreOffice URL." & @CRLF & @CRLF & _
			"Original Path: " & @CRLF & $sOfficePath & @CRLF & @CRLF & "Result: " & @CRLF & _
			_LO_PathConvert($sOfficePath, $LO_PATHCONV_OFFICE_RETURN))
EndFunc

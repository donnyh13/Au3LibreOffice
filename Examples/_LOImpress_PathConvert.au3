#include <MsgBoxConstants.au3>

#include "..\LibreOfficeImpress.au3"

Example()

Func Example()
	Local $sPCPath, $sOfficePath

	; An example of a computer path
	$sPCPath = "C:\A Folder With Spaces\FolderWithASemicolon;\{FolderWithBrackets}\TestDocument.odp"

	; An example of a Libre Office URL
	$sOfficePath = "file:///C:/A%20Folder%20With%20Spaces/FolderWithASemicolon%3B/%7BFolderWithBrackets%7D/TestDocument.odp"

	; Auto Return From a Computer Path
	MsgBox($MB_OK + $MB_TOPMOST, Default, "Auto_Return -- Computer Path" & @CRLF & "This is the result from converting a Computer path to Libre Office URL automatically." & @CRLF & _
			"Original Path: " & @CRLF & $sPCPath & @CRLF & "Result: " & @CRLF & _
			_LOImpress_PathConvert($sPCPath, $LOI_PATHCONV_AUTO_RETURN))

	; Auto Return From a Libre Office URL
	MsgBox($MB_OK + $MB_TOPMOST, Default, "Auto_Return -- Libre Office URL" & @CRLF & "This is the result from converting a Libre Office URL to Computer Path automatically." & @CRLF & _
			"Original Path: " & @CRLF & $sOfficePath & @CRLF & "Result: " & @CRLF & _
			_LOImpress_PathConvert($sOfficePath, $LOI_PATHCONV_AUTO_RETURN))

	; Return From a Libre Office URL to Computer Path conversion
	MsgBox($MB_OK + $MB_TOPMOST, Default, "PCPATH_RETURN -- Libre Office URL" & @CRLF & "This is the result from converting a Libre Office URL to Computer Path." & @CRLF & _
			"Original Path: " & @CRLF & $sOfficePath & @CRLF & "Result: " & @CRLF & _
			_LOImpress_PathConvert($sOfficePath, $LOI_PATHCONV_PCPATH_RETURN))

	; Return From a Libre Office URL to Computer Path conversion when the path is already a computer path.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "PCPATH_RETURN -- Computer Path" & @CRLF & "This is the result from converting a Libre Office URL to Computer Path " & _
			"when the path is already a computer path." & @CRLF & _
			"Original Path: " & @CRLF & $sPCPath & @CRLF & "Result: " & @CRLF & _
			_LOImpress_PathConvert($sPCPath, $LOI_PATHCONV_PCPATH_RETURN))

	; Return From a Computer Path to Libre Office URL conversion
	MsgBox($MB_OK + $MB_TOPMOST, Default, "OFFICE_RETURN -- Computer Path" & @CRLF & "This is the result from converting a Computer Path to Libre Office URL." & @CRLF & _
			"Original Path: " & @CRLF & $sPCPath & @CRLF & "Result: " & @CRLF & _
			_LOImpress_PathConvert($sPCPath, $LOI_PATHCONV_OFFICE_RETURN))

	; Return From a Computer Path to Libre Office URL conversion when the path is already a Libre Office path.
	MsgBox($MB_OK + $MB_TOPMOST, Default, "OFFICE_RETURN -- Libre Office URL" & @CRLF & "This is the result from converting a Computer Path to Libre Office URL " & _
			"when the path is already a Libre Office URL." & @CRLF & _
			"Original Path: " & @CRLF & $sOfficePath & @CRLF & "Result: " & @CRLF & _
			_LOImpress_PathConvert($sOfficePath, $LOI_PATHCONV_OFFICE_RETURN))
EndFunc

Func _ERROR($sErrorText)
	MsgBox($MB_OK + $MB_ICONERROR + $MB_TOPMOST, "Error", $sErrorText)
	Exit
EndFunc

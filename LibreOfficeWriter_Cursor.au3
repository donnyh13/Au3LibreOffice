#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel /tcl=1
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Writer
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Retrieving and manipulating a Cursor in L.O. Writer.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_CursorGetDataType
; _LOWriter_CursorGetStatus
; _LOWriter_CursorGetString
; _LOWriter_CursorGetType
; _LOWriter_CursorGoToRange
; _LOWriter_CursorHyperlink
; _LOWriter_CursorInsertControlChar
; _LOWriter_CursorInsertString
; _LOWriter_CursorMove
; _LOWriter_CursorParObjCopy
; _LOWriter_CursorParObjCreateList
; _LOWriter_CursorParObjDelete
; _LOWriter_CursorParObjPaste
; _LOWriter_CursorParObjSectionsGet
; _LOWriter_CursorTextCursorCreate
; _LOWriter_CursorViewCursorGetObj
; _LOWriter_CursorViewCursorGetPosition
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorGetDataType
; Description ...: Determines what type of Text data a Cursor is currently in.
; Syntax ........: _LOWriter_CursorGetType(ByRef $oDoc, ByRef $oCursor)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - A Cursor Object returned from any Cursor Object creation Or retrieval functions.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Return value will be one of the constants, $LOW_CURDATA_* as defined in LibreOfficeWriter_Constants.au3.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oCursor not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Cursor Data Type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Returns what type of data a cursor is currently located in, such as a TextTable, Footnote etc.
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorTextCursorCreate, _LOWriter_TableCellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_PageStyleHeaderCreateTextCursor, _LOWriter_PageStyleFooterCreateTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorGetDataType(ByRef $oDoc, ByRef $oCursor)
	Local $iCursorDataType

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iCursorDataType = __LOWriter_Internal_CursorGetDataType($oDoc, $oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCursorDataType)
EndFunc   ;==>_LOWriter_CursorGetDataType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorGetStatus
; Description ...: Retrieve the current status of a cursor.
; Syntax ........: _LOWriter_CursorGetStatus(ByRef $oCursor, $iFlag)
; Parameters ....: $oCursor             - A Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $iFlag               - The Requested status to return, see constants, $LOW_CURSOR_STAT_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
; Return values .: Success: Variable.
;                  @Error: 0, @Extended: 0, Return: Variable = Success. The requested status was successfully returned. See called flag for expected return type.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iFlag not an Integer.
;                  @Error: 1, @Extended: 3 = Flag called in $iFlag not available for "Text" cursor.
;                  @Error: 1, @Extended: 4 = Flag called in $iFlag not available for "Table" cursor.
;                  @Error: 1, @Extended: 5 = Flag called in $iFlag not available for "View" cursor.
;                  @Error: 1, @Extended: 6 = $oCursor unknown cursor type.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Cursor Type.
;                  @Error: 3, @Extended: 2 = Error retrieving Status for Text Cursor.
;                  @Error: 3, @Extended: 3 = Error retrieving Status for Table Cursor.
;                  @Error: 3, @Extended: 4 = Error retrieving Status for View Cursor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only certain flags work for certain types of cursors:
;                  # Text And View Cursor Status Flag Constants:
;                  - $LOW_CURSOR_STAT_IS_COLLAPSED
;                  # Text Cursor Status Flag Constants:
;                  - $LOW_CURSOR_STAT_IS_START_OF_WORD,
;                  - $LOW_CURSOR_STAT_IS_END_OF_WORD,
;                  - $LOW_CURSOR_STAT_IS_START_OF_SENTENCE,
;                  - $LOW_CURSOR_STAT_IS_END_OF_SENTENCE,
;                  - $LOW_CURSOR_STAT_IS_START_OF_PAR,
;                  - $LOW_CURSOR_STAT_IS_END_OF_PAR,
;                  # View Cursor Status Flag Constants:
;                  - $LOW_CURSOR_STAT_IS_START_OF_LINE,
;                  - $LOW_CURSOR_STAT_IS_END_OF_LINE,
;                  - $LOW_CURSOR_STAT_GET_PAGE,
;                  # Table Cursor Status Flag Constants:
;                  - $LOW_CURSOR_STAT_GET_RANGE_NAME
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorTextCursorCreate, _LOWriter_TableCellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_PageStyleHeaderCreateTextCursor, _LOWriter_PageStyleFooterCreateTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_CursorGetType
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorGetStatus(ByRef $oCursor, $iFlag)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType
	Local $vReturn
	Local $aiCommands[11]
	$aiCommands[$LOW_CURSOR_STAT_IS_COLLAPSED] = ".isCollapsed()"
	$aiCommands[$LOW_CURSOR_STAT_IS_START_OF_WORD] = ".isStartOfWord()"
	$aiCommands[$LOW_CURSOR_STAT_IS_END_OF_WORD] = ".isEndOfWord()"
	$aiCommands[$LOW_CURSOR_STAT_IS_START_OF_SENTENCE] = ".isStartOfSentence()"
	$aiCommands[$LOW_CURSOR_STAT_IS_END_OF_SENTENCE] = ".isEndOfSentence()"
	$aiCommands[$LOW_CURSOR_STAT_IS_START_OF_PAR] = ".isStartOfParagraph()"
	$aiCommands[$LOW_CURSOR_STAT_IS_END_OF_PAR] = ".isEndOfParagraph()"
	$aiCommands[$LOW_CURSOR_STAT_IS_START_OF_LINE] = ".isAtStartOfLine()"
	$aiCommands[$LOW_CURSOR_STAT_IS_END_OF_LINE] = ".isAtEndOfLine()"
	$aiCommands[$LOW_CURSOR_STAT_GET_PAGE] = ".getPage()"
	$aiCommands[$LOW_CURSOR_STAT_GET_RANGE_NAME] = ".getRangeName()"

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsInt($iFlag) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Switch $iCursorType
		Case $LOW_CURTYPE_TEXT_CURSOR
			If Not __LO_IntIsBetween($iFlag, $LOW_CURSOR_STAT_IS_COLLAPSED, $LOW_CURSOR_STAT_IS_END_OF_PAR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

			$vReturn = Execute("$oCursor" & $aiCommands[$iFlag])
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

			Return SetError($__LO_STATUS_SUCCESS, 0, $vReturn)

		Case $LOW_CURTYPE_TABLE_CURSOR
			If Not ($iFlag = $LOW_CURSOR_STAT_GET_RANGE_NAME) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

			$vReturn = Execute("$oCursor" & $aiCommands[$iFlag])
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

			Return SetError($__LO_STATUS_SUCCESS, 0, $vReturn)

		Case $LOW_CURTYPE_VIEW_CURSOR
			If Not __LO_IntIsBetween($iFlag, $LOW_CURSOR_STAT_IS_START_OF_LINE, $LOW_CURSOR_STAT_GET_PAGE, "", $LOW_CURSOR_STAT_IS_COLLAPSED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

			$vReturn = Execute("$oCursor" & $aiCommands[$iFlag])
			If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0)

			Return SetError($__LO_STATUS_SUCCESS, 0, $vReturn)

		Case Else

			Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0) ; unknown cursor data type.
	EndSwitch
EndFunc   ;==>_LOWriter_CursorGetStatus

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorGetString
; Description ...: Retrieve the string of text currently selected or contained in a paragraph object.
; Syntax ........: _LOWriter_CursorGetString(ByRef $oObj)
; Parameters ....: $oObj                - A Cursor Object returned from any Cursor Object creation or retrieval functions with Data selected, or a Paragraph Object returned from _LOWriter_CursorParObjCreateList function.
; Return values .: Success: String
;                  @Error: 0, @Extended: 0, Return: String = Success. The selected text in String format.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oObj not an Object.
;                  @Error: 1, @Extended: 2 = $oObj doesn't support Paragraph Properties service.
;                  @Error: 1, @Extended: 3 = $oObj is a TableCursor. Can only use View Cursor or Text Cursor.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Cursor type.
;                  @Error: 3, @Extended: 2 = Failed to retrieve String.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: LibreOffice documentation states that when used in LibreOffice Basic, GetString is limited to 64kb's in size. I do not know if the same limitation applies to any outside use of GetString (such as through Autoit).
;                  If there are multiple selections, the returned value will be an empty string ("").
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorTextCursorCreate, _LOWriter_TableCellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_PageStyleHeaderCreateTextCursor, _LOWriter_PageStyleFooterCreateTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorGetString(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sString

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.style.ParagraphProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If $oObj.supportsService("com.sun.star.text.TextCursor") Or $oObj.supportsService("com.sun.star.text.TextViewCursor") Then
		Local $iCursorType = __LOWriter_Internal_CursorGetType($oObj)
		If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	EndIf

	$sString = $oObj.getString()
	If Not IsString($sString) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $sString)
EndFunc   ;==>_LOWriter_CursorGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorGetType
; Description ...: Determine what type a Cursor Object is, such as a TableCursor, Text Cursor or a ViewCursor.
; Syntax ........: _LOWriter_CursorGetType(ByRef $oCursor)
; Parameters ....: $oCursor             - A Cursor Object returned from any Cursor Object creation Or retrieval functions.
; Return values .: Success: Integer.
;                  @Error: 0, @Extended: 0, Return: Integer = Success. Return value will be one of the Constants, $LOW_CURTYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCursor not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Cursor Type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Will also work for Paragraph object and paragraph section objects.
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorTextCursorCreate, _LOWriter_TableCellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_PageStyleHeaderCreateTextCursor, _LOWriter_PageStyleFooterCreateTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_CursorParObjCreateList, _LOWriter_CursorParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorGetType(ByRef $oCursor)
	Local $iCursorType

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $iCursorType)
EndFunc   ;==>_LOWriter_CursorGetType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorGoToRange
; Description ...: Moves a Text or View cursor to another View or Text Cursor Position or Range.
; Syntax ........: _LOWriter_CursorGoToRange(ByRef $oCursor, ByRef $oRange[, $bSelect = False])
; Parameters ....: $oCursor             - A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $oRange              - A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions to move $oCursor to.
;                  $bSelect             - [optional] Default is False. If True, the selection is expanded or created from original cursor location to Range location.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Cursor successfully moved to $oRange position.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCursor not an Object.
;                  @Error: 1, @Extended: 2 = $oRange not an Object.
;                  @Error: 1, @Extended: 3 = $bSelect not a Boolean.
;                  @Error: 1, @Extended: 4 = $oCursor not a Text or View Cursor.
;                  @Error: 1, @Extended: 5 = $oRange is a Table Cursor, and is not supported.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error determining $oCursor cursor type.
;                  @Error: 3, @Extended: 2 = Error determining $oRange cursor type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the Cursor being used as a range has anything selected, the selection will be selected in the Cursor called in $oCursor also.
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorTextCursorCreate, _LOWriter_TableCellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_PageStyleHeaderCreateTextCursor, _LOWriter_PageStyleFooterCreateTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_CursorParObjCreateList, _LOWriter_CursorParObjSectionsGet, _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorGoToRange(ByRef $oCursor, ByRef $oRange, $bSelect = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType, $iRangeType

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bSelect) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iRangeType = __LOWriter_Internal_CursorGetType($oRange)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If ($iCursorType <> $LOW_CURTYPE_TEXT_CURSOR) And ($iCursorType <> $LOW_CURTYPE_VIEW_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRangeType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oCursor.gotoRange($oRange, $bSelect)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CursorGoToRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorHyperlink
; Description ...: Set or Retrieve hyperlink settings at the cursor location.
; Syntax ........: _LOWriter_CursorHyperlink(ByRef $oCursor[, $sLinkAddress = Null[, $sLinkName = Null]])
; Parameters ....: $oCursor             - A Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $sLinkAddress        - [optional] Default is Null. The hyperlink address. See remarks.
;                  $sLinkName           - [optional] Default is Null. The hyperlink name.
; Return values .: Success: 1 or Array.
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Settings were successfully set.
;                  @Error: 0, @Extended: 1, Return: Array = Success. All optional parameters were called with Null, returning current settings in a 2 Element Array with values in order of function parameters.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCursor not an Object, and is not called with Default keyword.
;                  @Error: 1, @Extended: 2 = $oCursor is a TableCursor, and is not supported.
;                  @Error: 1, @Extended: 3 = $sLinkAddress not a String.
;                  @Error: 1, @Extended: 4 = $sLinkName not a String.
;                  --Property Setting Errors--
;                  @Error: 4, @Extended: ? = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sLinkAddress
;                  |                               2 = Error setting $sLinkName
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or by calling all other parameters with the Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  If no text is selected when setting or retrieving hyperlink settings, the whole word where the cursor is located is used.
;                  $sLinkAddress can be a website URL (www.abc.com), or a mail to link ("mailto:abc@email.ca?subject=Testing123". It can also be a path to a file ("C:\Users\xyz.docx").
;                  When setting $sLinkAddress to a computer path, it is recommended to convert it to URL notation using _LO_PathConvert. If you do not, a property setting error may be triggered because LibreOffice converts the path itself.
;                  When retrieving the current value for $sLinkAddress, the value will be directly returned from LibreOffice, if the current value is a path, it will be in URL format and will need to be converted using _LO_PathConvert.
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorTextCursorCreate, _LOWriter_TableCellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_PageStyleHeaderCreateTextCursor, _LOWriter_PageStyleFooterCreateTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_CursorInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorHyperlink(ByRef $oCursor, $sLinkAddress = Null, $sLinkName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $asLink[2]
	Local $iError = 0

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oCursor.supportsService("com.sun.star.text.TextTableCursor") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; Can't set hyperlink settings for a Table Cursor.

	If __LO_VarsAreNull($sLinkAddress, $sLinkName) Then
		__LO_ArrayFill($asLink, $oCursor.HyperLinkURL(), $oCursor.HyperLinkName())

		Return SetError($__LO_STATUS_SUCCESS, 1, $asLink)
	EndIf

	If ($sLinkAddress <> Null) Then
		If Not IsString($sLinkAddress) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oCursor.HyperLinkURL = $sLinkAddress
		$iError = ($oCursor.HyperLinkURL() = $sLinkAddress) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sLinkName <> Null) Then
		If Not IsString($sLinkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

		$oCursor.HyperLinkName = $sLinkName
		$iError = ($oCursor.HyperLinkName() = $sLinkName) ? ($iError) : (BitOR($iError, 2))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_CursorHyperlink

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorInsertControlChar
; Description ...: Insert a control character at the cursor position.
; Syntax ........: _LOWriter_CursorInsertControlChar(ByRef $oDoc, ByRef $oCursor, $iConChar[, $bOverwrite = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $iConChar            - (0-5) The control character to insert. See constants, $LOW_CON_CHAR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bOverwrite          - [optional] Default is False. If True, and the cursor object has text selected, it is overwritten, else if False, the character is inserted to the left of the selection.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Control Character was successfully inserted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oCursor not an Object.
;                  @Error: 1, @Extended: 3 = $iConChar not an Integer, less than 0 or greater than 5. See Constants, $LOW_CON_CHAR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 4 = $bOverwrite not a Boolean.
;                  @Error: 1, @Extended: 5 = $oCursor is a TableCursor. Can only use View Cursor or Text Cursor.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Cursor type.
;                  @Error: 3, @Extended: 2 = Error creating Text Cursor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorTextCursorCreate, _LOWriter_TableCellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_PageStyleHeaderCreateTextCursor, _LOWriter_PageStyleFooterCreateTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_CursorInsertString
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorInsertControlChar(ByRef $oDoc, ByRef $oCursor, $iConChar, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType
	Local $oTextCursor = $oCursor

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not __LO_IntIsBetween($iConChar, $LOW_CON_CHAR_PAR_BREAK, $LOW_CON_CHAR_APPEND_PAR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If ($iCursorType = $LOW_CURTYPE_VIEW_CURSOR) Then $oTextCursor = _LOWriter_CursorTextCursorCreate($oDoc, False, True)

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$oTextCursor.Text.insertControlCharacter($oTextCursor, $iConChar, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CursorInsertControlChar

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorInsertString
; Description ...: Insert a string at a cursor position.
; Syntax ........: _LOWriter_CursorInsertString(ByRef $oDoc, ByRef $oCursor, $sString[, $bOverwrite = False])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $sString             - A String to insert.
;                  $bOverwrite          - [optional] Default is False. If True, and the cursor object has text selected, the selection is overwritten, else if False, the string is inserted to the left of the selection. If there are multiple selections, the string is inserted to the left of the last selection, and none are overwritten.
; Return values .: Success: 1
;                  @Error: 0, @Extended: 0, Return: 1 = Success. String was successfully inserted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oCursor not an Object.
;                  @Error: 1, @Extended: 3 = $sString not a string..
;                  @Error: 1, @Extended: 4 = $bOverwrite not a Boolean.
;                  @Error: 1, @Extended: 5 = $oCursor is a TableCursor. Can only use View Cursor or Text Cursor.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error retrieving Cursor type.
;                  @Error: 3, @Extended: 2 = Error creating Text Cursor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To prevent accidental and unwanted newlines, @CRLF is automatically replaced with @CR to match LibreOffice's newline style.
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorTextCursorCreate, _LOWriter_TableCellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_PageStyleHeaderCreateTextCursor, _LOWriter_PageStyleFooterCreateTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorInsertString(ByRef $oDoc, ByRef $oCursor, $sString, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType
	Local $oTextCursor = $oCursor

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If ($iCursorType = $LOW_CURTYPE_VIEW_CURSOR) Then $oTextCursor = _LOWriter_CursorTextCursorCreate($oDoc, False, True)

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	; Exchange CRLF for CR to prevent errors.
	$sString = StringRegExpReplace($sString, @CRLF, @CR)

	$oTextCursor.Text.insertString($oTextCursor, $sString, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CursorInsertString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorMove
; Description ...: Move a Cursor object in a document. Also for creating/Expanding selections.
; Syntax ........: _LOWriter_CursorMove(ByRef $oCursor, $iMove[, $iCount = 1[, $bSelect = False]])
; Parameters ....: $oCursor             - A Cursor Object returned from any Cursor Object creation Or retrieval functions.
;                  $iMove               - The movement command. See remarks and Constants, $LOW_VIEWCUR_, $LOW_TEXTCUR_, $LOW_TABLECUR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iCount              - [optional] Default is 1. Number of movements to make. See remarks.
;                  $bSelect             - [optional] Default is False. Whether to select data during this cursor movement. See remarks.
; Return values .: Success: Boolean.
;                  @Error: 0, @Extended: ?, Return: Boolean = Success, Cursor object movement was processed successfully. Returning True if the full count of movements were successful, else False if none or only partially successful. @Extended set to number of successful movements. Or Page Number for "gotoPage" command. See Remarks
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCursor not an Object.
;                  @Error: 1, @Extended: 2 = $iMove not an Integer.
;                  @Error: 1, @Extended: 3 = $iMove mismatch with Cursor type. See Cursor Type/Move Type Constants, $LOW_VIEWCUR_, $LOW_TEXTCUR_, $LOW_TABLECUR_* as defined in LibreOfficeWriter_Constants.au3.
;                  @Error: 1, @Extended: 4 = $iCount not an Integer or is a negative.
;                  @Error: 1, @Extended: 5 = $bSelect not a Boolean.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Error determining cursor type.
;                  @Error: 3, @Extended: 2 = Error processing cursor move.
;                  @Error: 3, @Extended: 3 = Failed to retrieve Current Page number.
;                  @Error: 3, @Extended: 4 = $oCursor Object unknown cursor type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iMove may be called with any of the following constants depending on the Cursor type you are intending to move.
;                  Only some movements accept movement amounts (such as "goRight" 2) etc. Also only some accept creating/ extending a selection of text/ data. They will be specified below.
;                  To Clear /Unselect a current selection, you can input a move such as "goRight", 0, False.
;                  #Cursor Movement Constants which accept Number of Moves and Selecting:
;                  + ViewCursor
;                  - $LOW_VIEWCUR_GO_DOWN,
;                  - $LOW_VIEWCUR_GO_UP,
;                  - $LOW_VIEWCUR_GO_LEFT,
;                  - $LOW_VIEWCUR_GO_RIGHT,
;                  + TextCursor
;                  - $LOW_TEXTCUR_GO_LEFT,
;                  - $LOW_TEXTCUR_GO_RIGHT,
;                  - $LOW_TEXTCUR_GOTO_NEXT_WORD,
;                  - $LOW_TEXTCUR_GOTO_PREV_WORD,
;                  - $LOW_TEXTCUR_GOTO_NEXT_SENTENCE,
;                  - $LOW_TEXTCUR_GOTO_PREV_SENTENCE,
;                  - $LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH,
;                  - $LOW_TEXTCUR_GOTO_PREV_PARAGRAPH,
;                  + TableCursor
;                  - $LOW_TABLECUR_GO_LEFT,
;                  - $LOW_TABLECUR_GO_RIGHT,
;                  - $LOW_TABLECUR_GO_UP,
;                  - $LOW_TABLECUR_GO_DOWN,
;                  # Cursor Movements which accept Number of Moves Only:
;                  + ViewCursor
;                  - $LOW_VIEWCUR_JUMP_TO_NEXT_PAGE,
;                  - $LOW_VIEWCUR_JUMP_TO_PREV_PAGE,
;                  - $LOW_VIEWCUR_SCREEN_DOWN,
;                  - $LOW_VIEWCUR_SCREEN_UP,
;                  # Cursor Movements which accept Selecting Only:
;                  + ViewCursor
;                  - $LOW_VIEWCUR_GOTO_END_OF_LINE,
;                  - $LOW_VIEWCUR_GOTO_START_OF_LINE,
;                  - $LOW_VIEWCUR_GOTO_START,
;                  - $LOW_VIEWCUR_GOTO_END,
;                  + TextCursor
;                  - $LOW_TEXTCUR_GOTO_START,
;                  - $LOW_TEXTCUR_GOTO_END,
;                  - $LOW_TEXTCUR_GOTO_END_OF_WORD,
;                  - $LOW_TEXTCUR_GOTO_START_OF_WORD,
;                  - $LOW_TEXTCUR_GOTO_END_OF_SENTENCE,
;                  - $LOW_TEXTCUR_GOTO_START_OF_SENTENCE,
;                  - $LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH,
;                  - $LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH,
;                  + TableCursor
;                  - $LOW_TABLECUR_GOTO_START,
;                  - $LOW_TABLECUR_GOTO_END,
;                  # Cursor Movements which accept nothing and are done once per call:
;                  + ViewCursor
;                  - $LOW_VIEWCUR_JUMP_TO_FIRST_PAGE,
;                  - $LOW_VIEWCUR_JUMP_TO_LAST_PAGE,
;                  - $LOW_VIEWCUR_JUMP_TO_END_OF_PAGE,
;                  - $LOW_VIEWCUR_JUMP_TO_START_OF_PAGE,
;                  + TextCursor
;                  - $LOW_TEXTCUR_COLLAPSE_TO_START,
;                  - $LOW_TEXTCUR_COLLAPSE_TO_END,
;                  # Misc. Cursor Movements:
;                  + ViewCursor
;                  - $LOW_VIEWCUR_JUMP_TO_PAGE
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorTextCursorCreate, _LOWriter_TableCellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_PageStyleHeaderCreateTextCursor, _LOWriter_PageStyleFooterCreateTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_TableCreateCursor, _LOWriter_CursorGoToRange
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorMove(ByRef $oCursor, $iMove, $iCount = 1, $bSelect = False)
	Local $iCursorType
	Local $bMoved = False

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Switch $iCursorType
		Case $LOW_CURTYPE_TEXT_CURSOR
			$bMoved = __LOWriter_TextCursorMove($oCursor, $iMove, $iCount, $bSelect)

			Return SetError(@error, @extended, $bMoved)

		Case $LOW_CURTYPE_TABLE_CURSOR
			$bMoved = __LOWriter_TableCursorMove($oCursor, $iMove, $iCount, $bSelect)

			Return SetError(@error, @extended, $bMoved)

		Case $LOW_CURTYPE_VIEW_CURSOR
			$bMoved = __LOWriter_ViewCursorMove($oCursor, $iMove, $iCount, $bSelect)

			Return SetError(@error, @extended, $bMoved)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 4, 0) ; unknown cursor type.
	EndSwitch
EndFunc   ;==>_LOWriter_CursorMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorParObjCopy
; Description ...: "Copies" data selected by the ViewCursor, returning an Object for use in inserting later.
; Syntax ........: _LOWriter_CursorParObjCopy(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. Data was successfully copied, returning an Object for use in _LOWriter_CursorParObjPaste.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to Copy Selected Data, make sure Data is selected using the ViewCursor.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Data you desire to be copied MUST be selected with the ViewCursor, see _LOWriter_DocSelection.
;                  This function works essentially the same as Copy/ Ctrl+C, except it doesn't use your clipboard.
;                  The Object returned is used in _LOWriter_CursorParObjPaste to insert the data again.
;                  Copying data this way works for Tables, Images, frames and Text, including with direct formatting, etc.
;                  Data copied can be inserted into the same or another document.
; Related .......: _LOWriter_CursorParObjPaste, _LOWriter_DocSelection, _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorParObjCopy(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oObj = $oDoc.CurrentController.getTransferable() ; Copy
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oObj)
EndFunc   ;==>_LOWriter_CursorParObjCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorParObjCreateList
; Description ...: Return Objects for every paragraph contained in a specific section of a document.
; Syntax ........: _LOWriter_CursorParObjCreateList(ByRef $oCursor[, $bTableCheck = False])
; Parameters ....: $oCursor             - A Cursor Object returned from any Cursor Object creation Or retrieval functions. See Remarks
;                  $bTableCheck         - [optional] Default is False. If True, returned array will be 2 dimensional, with the second column indicating if the paragraph object is a Table (True) or not (False).
; Return values .: Success: 1D or 2D Array
;                  @Error: 0, @Extended: ?, Return: Array = Success. Returning an Array of Paragraph Objects, @Extended is set to the number of results.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCursor not an Object.
;                  @Error: 1, @Extended: 2 = $bTableCheck not a Boolean.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create Enumeration of Paragraphs.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $oCursor can be either a ViewCursor or a TextCursor, the paragraphs are enumerated for the area the cursor is currently within, for example, the ViewCursor is currently in a Table, the enumeration of paragraphs would be for the Cell the cursor was presently in.
;                  In the main document the enumeration would be for the entire Text Body, in the Header, it would for the that Header for that Page Style etc.
;                  The different possible areas are: Text Body, Table Cell, Header, Footer, Footnote, Endnote, Frame.
;                  Returns an Array of objects for Direct Formatting paragraphs in a document, or for copying and inserting etc.
;                  Table Objects returned from this function can be used as a regular Table Object to modify the Table with.
; Related .......: _LOWriter_CursorParObjSectionsGet, _LOWriter_DocSelection, _LOWriter_CursorParObjDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorParObjCreateList(ByRef $oCursor, $bTableCheck = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oEnum, $oPar
	Local $iCount = 0
	Local $aoParagraphs[1]

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bTableCheck) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oEnum = $oCursor.Text.createEnumeration()
	If Not IsObj($oEnum) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bTableCheck = True) Then ReDim $aoParagraphs[1][2]

	While $oEnum.hasMoreElements()
		$oPar = $oEnum.nextElement()

		If ($bTableCheck = True) Then
			If UBound($aoParagraphs) <= ($iCount) Then ReDim $aoParagraphs[UBound($aoParagraphs) * 2][2]
			$aoParagraphs[$iCount][0] = $oPar
			$aoParagraphs[$iCount][1] = ($oPar.supportsService("com.sun.star.text.TextTable"))
			$iCount += 1

		Else
			If UBound($aoParagraphs) <= ($iCount) Then ReDim $aoParagraphs[UBound($aoParagraphs) * 2]
			$aoParagraphs[$iCount] = $oPar
			$iCount += 1
		EndIf

		Sleep((IsInt($iCount / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	WEnd

	If ($bTableCheck = True) Then
		ReDim $aoParagraphs[$iCount][2]

	Else
		ReDim $aoParagraphs[$iCount]
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoParagraphs)
EndFunc   ;==>_LOWriter_CursorParObjCreateList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorParObjDelete
; Description ...: Delete a Paragraph Object returned from _LOWriter_CursorParObjCreateList.
; Syntax ........: _LOWriter_CursorParObjDelete(ByRef $oParObj)
; Parameters ....: $oParObj             - A Paragraph Object returned by _LOWriter_CursorParObjCreateList.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Paragraph was successfully deleted.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oParObj not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You cannot delete the last paragraph contained in a Text area, it will cause a COM error.
; Related .......: _LOWriter_CursorParObjCreateList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorParObjDelete(ByRef $oParObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oParObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oParObj.supportsService("com.sun.star.text.TextTable")) Then
		$oParObj.dispose()

	Else
		$oParObj.Text.removeTextContent($oParObj)
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CursorParObjDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorParObjPaste
; Description ...: Inserts a ParObjCopy Object at the current ViewCursor location.
; Syntax ........: _LOWriter_CursorParObjPaste(ByRef $oDoc, ByRef $oParObj)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oParObj             - A Object returned from _LOWriter_CursorParObjCopy to insert.
; Return values .: Success: Integer
;                  @Error: 0, @Extended: 0, Return: 1 = Success. Data was successfully inserted at the ViewCursor location.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $oParObj not an Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_CursorParObjCopy
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorParObjPaste(ByRef $oDoc, ByRef $oParObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oParObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.insertTransferable($oParObj)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CursorParObjPaste

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorParObjSectionsGet
; Description ...: Break a Paragraph Object into individual Sections for Direct Formatting etc.
; Syntax ........: _LOWriter_CursorParObjSectionsGet(ByRef $oParagraph)
; Parameters ....: $oParagraph          - A Paragraph Object returned from _LOWriter_CursorParObjCreateList function. Make sure it's not a Table!
; Return values .: Success: Array
;                  @Error: 0, @Extended: 0, Return: Array = Success. A two column array. See remarks.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oParagraph is not an Object.
;                  @Error: 1, @Extended: 2 = $oParagraph does not support Paragraph service. Not a paragraph Object.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Error enumerating Paragraph sections.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Paragraph in a Document may have more than one section if it contains direct formatting, foot/endnote anchors etc.
;                  The Array returned is a two column array with array[0][0] containing the section Object.
;                  The second column, array[0][1] contains the section data type column being one of the following possible types:
;                  |- Text – String content.
;                  |- TextField – TextField content.
;                  |- TextContent – Indicates that text content is anchored as or to a character that is not really part of the paragraph — for example, a text frame or a graphic object.
;                  |- ControlCharacter – Control character.
;                  |- Footnote – Footnote or Endnote. (This is just the anchor character for the footnote/Endnote, not the actual foot/endnote content.
;                  |- ReferenceMark – Reference mark.
;                  |- DocumentIndexMark – Document index mark.
;                  |- Bookmark – Bookmark.
;                  |- Redline – Redline portion, which is a result of the change-tracking feature.
;                  |- Ruby – a ruby attribute which is used in Asian text.
;                  |- Frame — a frame.
;                  |- SoftPageBreak — a soft page break.
;                  |- InContentMetadata — a text range with attached metadata.
;                  For Reference marks, document index marks, etc., 2 text portions will be generated, one for the start position and one for the end position.
; Related .......: _LOWriter_CursorParObjCreateList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorParObjSectionsGet(ByRef $oParagraph)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSecEnum, $oParSection
	Local $aoSections[1][2]
	Local $iCount = 0

	If Not IsObj($oParagraph) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParagraph.supportsService("com.sun.star.text.Paragraph") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSecEnum = $oParagraph.createEnumeration()
	If Not IsObj($oSecEnum) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	While $oSecEnum.hasMoreElements()
		$oParSection = $oSecEnum.nextElement()

		If UBound($aoSections) <= ($iCount + 1) Then ReDim $aoSections[UBound($aoSections) * 10][2]
		$aoSections[$iCount][0] = $oParSection
		$aoSections[$iCount][1] = $oParSection.TextPortionType()
		$iCount += 1
		Sleep((IsInt($iCount / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	WEnd
	ReDim $aoSections[$iCount][2]

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoSections)
EndFunc   ;==>_LOWriter_CursorParObjSectionsGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorTextCursorCreate
; Description ...: Create a TextCursor Object for future Textcursor related functional use.
; Syntax ........: _LOWriter_CursorTextCursorCreate(ByRef $oDoc[, $bCreateAtEnd = True[, $bCreateAtViewCursor = False]])
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or_LOWriter_DocCreate function.
;                  $bCreateAtEnd        - [optional] Default is True. If True, creates the new cursor at the end of the Document. Else cursor is created at the beginning.
;                  $bCreateAtViewCursor - [optional] Default is False. If True, create the Text cursor at the document's View Cursor. See Remarks
; Return values .: Success: Object.
;                  @Error: 0, @Extended: ?, Return: Object = Success, Cursor object was returned. @Extended can be on of the constants, $LOW_CURDATA_* as defined in LibreOfficeWriter_Constants.au3 indicating the current created cursor is in that type of data.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  @Error: 1, @Extended: 2 = $bCreateAtEnd not a Boolean.
;                  @Error: 1, @Extended: 3 = $bCreateAtViewCursor not a Boolean.
;                  @Error: 1, @Extended: 4 = $bCreateAtEnd and $bCreateAtViewCursor both called with True, set either one to False.
;                  --Initialization Errors--
;                  @Error: 2, @Extended: 1 = Failed to create Text Cursor Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve ViewCursor Object.
;                  @Error: 3, @Extended: 2 = Failed to Retrieve Cursor Data Type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The cursor Created by this function in a text document, is used for inserting text, reading text, etc.
;                  If you call $bCreateAtEnd with False, the new cursor is created at the beginning of the document, True creates the cursor at the very end of the document.
;                  Calling $bCreateAtViewCursor with True will create a Textcursor at the current ViewCursor position.
;                  There are two types of cursors in Word documents. The one you see, called the "ViewCursor", and one you do not see, called a "TextCursor". A "ViewCursor" is the blinking cursor you see when you are editing a Word document, there is only one per document. A "TextCursor" on the other hand, is an invisible cursor used for inserting text etc., into a Writer document. You can have multiple "TextCursors" per document.
;                  The ViewCursor knows how the data is displayed, but doesn’t know about the data itself. TextCursors however, know a lot about the data but very little about how it is displayed. For example, ViewCursors do not know about words or paragraphs, and TextCursors do not know about lines, screens, or pages.
; Related .......: _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorTextCursorCreate(ByRef $oDoc, $bCreateAtEnd = True, $bCreateAtViewCursor = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCursor, $oViewCursor
	Local $iCursorType = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bCreateAtEnd) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bCreateAtViewCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($bCreateAtEnd = True) And ($bCreateAtViewCursor = True) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	If $bCreateAtViewCursor Then
		$oViewCursor = $oDoc.CurrentController.getViewCursor()
		If Not IsObj($oViewCursor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		$iCursorType = __LOWriter_Internal_CursorGetDataType($oDoc, $oViewCursor)
		If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		$oCursor = $oViewCursor.Text.createTextCursorByRange($oViewCursor)
		If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Else
		$oCursor = $oDoc.getText.createTextCursor()
		If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$iCursorType = $LOW_CURDATA_BODY_TEXT

		If ($bCreateAtEnd = True) Then
			$oCursor.gotoEnd(False)

		Else
			$oCursor.gotoStart(False)
		EndIf
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, $iCursorType, $oCursor)
EndFunc   ;==>_LOWriter_CursorTextCursorCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorViewCursorGetObj
; Description ...: Retrieve the ViewCursor Object for a Document.
; Syntax ........: _LOWriter_CursorViewCursorGetObj(ByRef $oDoc)
; Parameters ....: $oDoc                - A Document object returned by a previous _LOWriter_DocOpen, _LOWriter_DocConnect, or_LOWriter_DocCreate function.
; Return values .: Success: Object
;                  @Error: 0, @Extended: 0, Return: Object = Success. The Object for the Document's View Cursor is returned for use in other Cursor related functions.
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oDoc not an Object.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to retrieve ViewCursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The ViewCursor deals with the visible cursor of the Document. You can have only one ViewCursor at a time. ViewCursors support movement commands that are directly related to viewing, but can also be used like a TextCursor in most instances.
; Related .......: _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorViewCursorGetObj(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oViewCursor

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oViewCursor = $oDoc.CurrentController.getViewCursor()
	If Not IsObj($oViewCursor) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oViewCursor)
EndFunc   ;==>_LOWriter_CursorViewCursorGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorViewCursorGetPosition
; Description ...: Retrieve View Cursor position in Hundredths of a Millimeter (HMM).
; Syntax ........: _LOWriter_CursorViewCursorGetPosition(ByRef $oCursor)
; Parameters ....: $oCursor             - A View Cursor Object returned by _LOWriter_CursorViewCursorGetObj function.
; Return values .: Success: Array.
;                  @Error: 0, @Extended: 0, Return: Array = Success. Returning a two element array containing "X" coordinate in the first element ($aArray[0]), and the "Y" Coordinate in the second element ($aArray[1]). In Hundredths of a Millimeter (HMM).
;                  Failure: 0 and sets @Error and @Extended to non-zero.
;                  --Input Errors--
;                  @Error: 1, @Extended: 1 = $oCursor not an Object.
;                  @Error: 1, @Extended: 2 = $oCursor not a View Cursor.
;                  --Processing Errors--
;                  @Error: 3, @Extended: 1 = Failed to determine Cursor type.
;                  @Error: 3, @Extended: 2 = Failed to retrieve "X" coordinate.
;                  @Error: 3, @Extended: 3 = Failed to retrieve "Y" coordinate.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Cursor Coordinate position is relative to the top-left of the first page of the document. In Hundredths of a Millimeter (HMM).
; Related .......: _LOWriter_CursorViewCursorGetObj, _LOWriter_CursorMove, _LO_UnitConvert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_CursorViewCursorGetPosition(ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType
	Local $aiCurrPos[2]

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCursorType = __LOWriter_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType <> $LOW_CURTYPE_VIEW_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$aiCurrPos[0] = $oCursor.getPosition.X()
	If Not IsInt($aiCurrPos[0]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	$aiCurrPos[1] = $oCursor.getPosition.Y()
	If Not IsInt($aiCurrPos[1]) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $aiCurrPos)
EndFunc   ;==>_LOWriter_CursorViewCursorGetPosition

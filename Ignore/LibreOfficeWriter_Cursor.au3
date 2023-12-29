#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

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
; _LOWriter_CursorGetType
; _LOWriter_CursorGoToRange
; _LOWriter_CursorMove
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorGetDataType
; Description ...: Determines what type of Text data a Cursor is currently in.
; Syntax ........: _LOWriter_CursorGetType(ByRef $oDoc, ByRef $oCursor)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval functions.
; Return values .: Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Cursor Data Type.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success, Return value will be one of the constants, $LOW_CURDATA_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Returns what type of data a cursor is currently located in, such as a TextTable, Footnote etc.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					 _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					 _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor,
;					_LOWriter_FootnoteGetTextCursor
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
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $iFlag               - an integer value. The Requested status to return, see constants, $LOW_CURSOR_STAT_* as defined in LibreOfficeWriter_Constants.au3. See remarks.
; Return values .: Success: Variable.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFlag not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iFlag set to flag not available for "Text" cursor.
;				   @Error 1 @Extended 4 Return 0 = $iFlag set to flag not available for "Table" cursor.
;				   @Error 1 @Extended 5 Return 0 = $iFlag  set to flag not available for "View" cursor.
;				   @Error 1 @Extended 6 Return 0 = $oCursor unknown cursor type.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Status for Text Cursor.
;				   @Error 3 @Extended 2 Return 0 = Error retrieving Status for Table Cursor.
;				   @Error 3 @Extended 3 Return 0 = Error retrieving Status for View Cursor.
;				   --Success--
;				   @Error 0 @Extended 0 Return Variable. = Success. The requested status was successfully returned. See called flag for expected return type.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only certain flags work for certain types of cursors:
;				   |	Text And View Cursor Status Flag Constants:
;				   + $LOW_CURSOR_STAT_IS_COLLAPSED
;				   |	Text Cursor Status Flag Constants:
;				   + $LOW_CURSOR_STAT_IS_START_OF_WORD,
;				   + $LOW_CURSOR_STAT_IS_END_OF_WORD,
;				   + $LOW_CURSOR_STAT_IS_START_OF_SENTENCE,
;				   + $LOW_CURSOR_STAT_IS_END_OF_SENTENCE,
;				   + $LOW_CURSOR_STAT_IS_START_OF_PAR,
;				   + $LOW_CURSOR_STAT_IS_END_OF_PAR,
;				   |	View Cursor Status Flag Constants:
;				   + $LOW_CURSOR_STAT_IS_START_OF_LINE,
;				   + $LOW_CURSOR_STAT_IS_END_OF_LINE,
;				   + $LOW_CURSOR_STAT_GET_PAGE,
;				   |	Table Cursor Status Flag Constants:
;				   + $LOW_CURSOR_STAT_GET_RANGE_NAME
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					 _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					 _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor,
;					_LOWriter_FootnoteGetTextCursor, _LOWriter_CursorGetType
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
			If Not __LOWriter_IntIsBetween($iFlag, $LOW_CURSOR_STAT_IS_COLLAPSED, $LOW_CURSOR_STAT_IS_END_OF_PAR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
			$vReturn = Execute("$oCursor" & $aiCommands[$iFlag])
			Return (@error > 0) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, $vReturn))

		Case $LOW_CURTYPE_TABLE_CURSOR
			If Not ($iFlag = $LOW_CURSOR_STAT_GET_RANGE_NAME) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
			$vReturn = Execute("$oCursor" & $aiCommands[$iFlag])
			Return (@error > 0) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, $vReturn))

		Case $LOW_CURTYPE_VIEW_CURSOR
			If Not __LOWriter_IntIsBetween($iFlag, $LOW_CURSOR_STAT_IS_START_OF_LINE, $LOW_CURSOR_STAT_GET_PAGE, "", $LOW_CURSOR_STAT_IS_COLLAPSED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
			$vReturn = Execute("$oCursor" & $aiCommands[$iFlag])
			Return (@error > 0) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, $vReturn))

		Case Else
			Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0) ; unknown cursor data type.
	EndSwitch

EndFunc   ;==>_LOWriter_CursorGetStatus

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorGetType
; Description ...: Determine what type a Cursor Object is, such as a TableCursor, Text Cursor or a ViewCursor.
; Syntax ........: _LOWriter_CursorGetType(ByRef $oCursor)
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval functions.
; Return values .: Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Cursor Type.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success, Return value will be one of the Constants, $LOW_CURTYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Will also work for Paragraph object and paragraph section objects.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					 _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					 _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor,
;					_LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList, _LOWriter_ParObjSectionsGet
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
; Parameters ....: $oCursor             - [in/out] an object. an object. A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions.
;                  $oRange              - [in/out] an object. an object. A Text or View Cursor Object returned from any Cursor Object creation or retrieval functions to move $oCursor to.
;                  $bSelect             - [optional] a boolean value. Default is False. If True, the selection is expanded or created from original cursor location to Range location.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bSelect not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $oCursor not a Text or View Cursor.
;				   @Error 1 @Extended 5 Return 0 = $oRange not a Text or View Cursor.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error determining $oCursor cursor type.
;				   @Error 3 @Extended 2 Return 0 = Error determining $oRange cursor type.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Cursor successfully moved to $oRange position.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the Cursor being used as a range has anything selected, the selection will be selected in the Cursor called in $oCursor also.
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					 _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor,	_LOWriter_FootnoteGetTextCursor, _LOWriter_ParObjCreateList,
;					_LOWriter_ParObjSectionsGet, _LOWriter_CursorMove
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
	If ($iRangeType <> $LOW_CURTYPE_TEXT_CURSOR) And ($iRangeType <> $LOW_CURTYPE_VIEW_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oCursor.gotoRange($oRange, $bSelect)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_CursorGoToRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_CursorMove
; Description ...: Move a Cursor object in a document. Also for creating/Expanding selections.
; Syntax ........: _LOWriter_CursorMove(ByRef $oCursor, $iMove[, $iCount = 1[, $bSelect = False]])
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval functions.
;                  $iMove               - an integer value. The movement command. See remarks and Constants, $LOW_VIEWCUR_, $LOW_TEXTCUR_, $LOW_TABLECUR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iCount              - [optional] an integer value. Default is 1. Number of movements to make. See remarks.
;                  $bSelect             - [optional] a boolean value. Default is False. Whether to select data during this cursor movement. See remarks.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;				   @Error 1 @Extended 3 Return 0 = $iMove mismatch with Cursor type. See Cursor Type/Move Type Constants, $LOW_VIEWCUR_, $LOW_TEXTCUR_, $LOW_TABLECUR_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $iCount not an integer or is a negative.
;				   @Error 1 @Extended 5 Return 0 = $bSelect not a Boolean.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error determining cursor type.
;				   @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;				   @Error 3 @Extended 3 Return 0 = $oCursor Object unknown cursor type.
;				   --Success--
;				   @Error 0 @Extended ? Return Boolean = Success, Cursor object movement was processed successfully. Returns True if the full count of movements were successful, else false if none or only partially successful. @Extended set to number of successful movements. Or Page Number for "gotoPage" command. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iMove may be set to any of the following constants depending on the Cursor type you are intending to move.
;					 Only some movements accept movement amounts (such as "goRight" 2) etc. Also only some accept creating/ extending a selection of text/ data. They will be specified below.
;					 To Clear /Unselect a current selection, you can input a move such as "goRight", 0, False.
;					#Cursor Movement Constants which accept Number of Moves and Selecting:
;					-ViewCursor
;						$LOW_VIEWCUR_GO_DOWN,
;						$LOW_VIEWCUR_GO_UP,
;						$LOW_VIEWCUR_GO_LEFT,
;						$LOW_VIEWCUR_GO_RIGHT,
;					-TextCursor
;						$LOW_TEXTCUR_GO_LEFT,
;						$LOW_TEXTCUR_GO_RIGHT,
;						$LOW_TEXTCUR_GOTO_NEXT_WORD,
;						$LOW_TEXTCUR_GOTO_PREV_WORD,
;						$LOW_TEXTCUR_GOTO_NEXT_SENTENCE,
;						$LOW_TEXTCUR_GOTO_PREV_SENTENCE,
;						$LOW_TEXTCUR_GOTO_NEXT_PARAGRAPH,
;						$LOW_TEXTCUR_GOTO_PREV_PARAGRAPH,
;					-TableCursor
;						$LOW_TABLECUR_GO_LEFT,
;						$LOW_TABLECUR_GO_RIGHT,
;						$LOW_TABLECUR_GO_UP,
;						$LOW_TABLECUR_GO_DOWN,
;					#Cursor Movements which accept Number of Moves Only:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_NEXT_PAGE,
;						$LOW_VIEWCUR_JUMP_TO_PREV_PAGE,
;						$LOW_VIEWCUR_SCREEN_DOWN,
;						$LOW_VIEWCUR_SCREEN_UP,
;					#Cursor Movements which accept Selecting Only:
;					-ViewCursor
;						$LOW_VIEWCUR_GOTO_END_OF_LINE,
;						$LOW_VIEWCUR_GOTO_START_OF_LINE,
;						$LOW_VIEWCUR_GOTO_START,
;						$LOW_VIEWCUR_GOTO_END,
;					-TextCursor
;						$LOW_TEXTCUR_GOTO_START,
;						$LOW_TEXTCUR_GOTO_END,
;						$LOW_TEXTCUR_GOTO_END_OF_WORD,
;						$LOW_TEXTCUR_GOTO_START_OF_WORD,
;						$LOW_TEXTCUR_GOTO_END_OF_SENTENCE,
;						$LOW_TEXTCUR_GOTO_START_OF_SENTENCE,
;						$LOW_TEXTCUR_GOTO_END_OF_PARAGRAPH,
;						$LOW_TEXTCUR_GOTO_START_OF_PARAGRAPH,
;					-TableCursor
;						$LOW_TABLECUR_GOTO_START,
;						$LOW_TABLECUR_GOTO_END,
;					#Cursor Movements which accept nothing and are done once per call:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_FIRST_PAGE,
;						$LOW_VIEWCUR_JUMP_TO_LAST_PAGE,
;						$LOW_VIEWCUR_JUMP_TO_END_OF_PAGE,
;						$LOW_VIEWCUR_JUMP_TO_START_OF_PAGE,
;					-TextCursor
;						$LOW_TEXTCUR_COLLAPSE_TO_START,
;						$LOW_TEXTCUR_COLLAPSE_TO_END,
;					#Misc. Cursor Movements:
;					-ViewCursor
;						$LOW_VIEWCUR_JUMP_TO_PAGE
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					 _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					 _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor,
;					_LOWriter_FootnoteGetTextCursor, _LOWriter_TableCreateCursor, _LOWriter_CursorGoToRange
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
			Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; unknown cursor data type.
	EndSwitch
EndFunc   ;==>_LOWriter_CursorMove

#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#Tidy_Parameters=/sf /reel
#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"
#include "LibreOffice_Helper.au3"
#include "LibreOffice_Internal.au3"

; Common includes for Calc
#include "LibreOfficeCalc_Internal.au3"

; Other includes for Calc

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for performing various Cursor movements.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOCalc_SheetCursorMove
; _LOCalc_TextCursorCharPosition
; _LOCalc_TextCursorCharSpacing
; _LOCalc_TextCursorEffect
; _LOCalc_TextCursorFont
; _LOCalc_TextCursorFontColor
; _LOCalc_TextCursorGetString
; _LOCalc_TextCursorGoToRange
; _LOCalc_TextCursorInsertString
; _LOCalc_TextCursorIsCollapsed
; _LOCalc_TextCursorMove
; _LOCalc_TextCursorOverline
; _LOCalc_TextCursorParObjCreateList
; _LOCalc_TextCursorParObjSectionsGet
; _LOCalc_TextCursorStrikeOut
; _LOCalc_TextCursorUnderline
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_SheetCursorMove
; Description ...: Move a Sheet Cursor object in a document. Also for creating/Expanding selections.
; Syntax ........: _LOCalc_SheetCursorMove(ByRef $oCursor, $iMove[, $iColumns = 0[, $iRows = 0[, $iCount = 1[, $bSelect = False]]]])
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval functions.
;                  $iMove               - an integer value. The movement command. See remarks and Constants, $LOC_TEXTCUR_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iColumns            - [optional] an integer value. Default is 0. The Number of Columns either to contain in the Range, or to move, depending on the called move command.
;                  $iRows               - [optional] an integer value. Default is 0. The Number of Rows either to contain in the Range, or to move, depending on the called move command.
;                  $iCount              - [optional] an integer value. Default is 1. Number of movements to make. See remarks.
;                  $bSelect             - [optional] a boolean value. Default is False. Whether to select data during this cursor movement. See remarks.
; Return values .: Success: 1.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iMove less than 0 or greater than highest move Constant. See Move Constants, $LOC_TEXTCUR_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iColumns not an integer.
;                  @Error 1 @Extended 5 Return 0 = $iRows not an integer.
;                  @Error 1 @Extended 6 Return 0 = $iCount not an integer or is a negative.
;                  @Error 1 @Extended 7 Return 0 = $bSelect not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error determining cursor type.
;                  @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;                  @Error 3 @Extended 3 Return 0 = $oCursor Object unknown cursor type.
;                  --Success--
;                  @Error 0 @Extended ? Return 1 = Success, Cursor object movement was processed successfully. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only some movements accept Column and Row Values, creating/ extending a selection of cells, etc. They will be specified below.
;                  #Cursor Movement Constants which accept Column and Row values:
;                  $LOC_SHEETCUR_COLLAPSE_TO_SIZE, Call $iColumns with the number of columns to resize the range to contain, counting from the Top-Left hand cell of the current range. And call $iRows with the number of Rows to resize the range to contain, counting from the Top-Left hand cell of the current range..
;                  $LOC_SHEETCUR_GOTO_OFFSET Call $iColumns with the number of columns to move, either left (negative number) or right (positive number), and call $iRows with the number of Rows to move up (negative number) or down (positive number).
;                  #Cursor Movements which accept Selecting Only:
;                  $LOC_SHEETCUR_GOTO_USED_AREA_START, Call $bSelect with a Boolean whether to select cells while performing this move (True) or not.
;                  $LOC_SHEETCUR_GOTO_USED_AREA_END Call $bSelect with a Boolean whether to select cells while performing this move (True) or not.
;                  #Cursor Movements which accept nothing and are done once per call:
;                  $LOC_SHEETCUR_COLLAPSE_TO_CURRENT_ARRAY,
;                  $LOC_SHEETCUR_COLLAPSE_TO_CURRENT_REGION,
;                  $LOC_SHEETCUR_COLLAPSE_TO_MERGED_AREA,
;                  $LOC_SHEETCUR_EXPAND_TO_ENTIRE_COLUMN,
;                  $LOC_SHEETCUR_EXPAND_TO_ENTIRE_ROW,
;                  $LOC_SHEETCUR_GOTO_START,
;                  $LOC_SHEETCUR_GOTO_END
;                  #Cursor Movements which accept only number of moves ($iCount):
;                  $LOC_SHEETCUR_GOTO_NEXT, Call $iCount with the number of moves to perform.
;                  $LOC_SHEETCUR_GOTO_PREV Call $iCount with the number of moves to perform.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_SheetCursorMove(ByRef $oCursor, $iMove, $iColumns = 0, $iRows = 0, $iCount = 1, $bSelect = False)
	Local $iCursorType
	Local $bMoved = False

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCursorType = __LOCalc_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Switch $iCursorType
		Case $LOC_CURTYPE_SHEET_CURSOR
			$bMoved = __LOCalc_SheetCursorMove($oCursor, $iMove, $iColumns, $iRows, $iCount, $bSelect)

			Return SetError(@error, @extended, $bMoved)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; unknown or wrong cursor type.
	EndSwitch
EndFunc   ;==>_LOCalc_SheetCursorMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorCharPosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size for a Text Cursor.
; Syntax ........: _LOCalc_TextCursorCharPosition(ByRef $oTextCursor[, $bAutoSuper = Null[, $iSuperScript = Null[, $bAutoSub = Null[, $iSubScript = Null[, $iRelativeSize = Null]]]]])
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bAutoSuper          - [optional] a boolean value. Default is Null. If True, automatic sizing for Superscript is active.
;                  $iSuperScript        - [optional] an integer value (0-100,14000). Default is Null. The Superscript percentage value. See Remarks.
;                  $bAutoSub            - [optional] a boolean value. Default is Null. If True, automatic sizing for Subscript is active.
;                  $iSubScript          - [optional] an integer value (-100-100,14000,-14000). Default is Null. Subscript percentage value. See Remarks.
;                  $iRelativeSize       - [optional] an integer value (1-100). Default is Null. The size percentage relative to current font size.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor does not support Character properties.
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bAutoSuper not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bAutoSub not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iSuperScript not an integer, or less than 0, higher than 100 and Not 14000.
;                  @Error 1 @Extended 7 Return 0 = $iSubScript not an integer, or less than -100, higher than 100 and Not 14000 or -14000.
;                  @Error 1 @Extended 8 Return 0 = $iRelativeSize not an integer, or less than 1, higher than 100.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Text Cursor for Paragraph Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $iSuperScript
;                  |                               2 = Error setting $iSubScript
;                  |                               4 = Error setting $iRelativeSize.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving current settings in any Text Cursor formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of inaccurate values.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Set either $iSubScript or $iSuperScript to 0 to return it to Normal setting.
;                  The way LibreOffice is set up Super/Subscript are set in the same setting, Superscript is a positive number from 1 to 100 (percentage), Subscript is a negative number set to -1 to -100 percentage. For the user's convenience this function accepts both positive and negative numbers for Subscript, if a positive number is called for Subscript, it is automatically set to a negative.
;                  Automatic Superscript has a integer value of 14000, Auto Subscript has a integer value of -14000. There is no settable setting of Automatic Super/Sub Script, though one exists, it is read-only in LibreOffice, consequently I have made two separate parameters to be able to determine if the user wants to automatically set Superscript or Subscript.
;                  If you set both Auto Superscript to True and Auto Subscript to True, or $iSuperScript to an integer and $iSubScript to an integer, Subscript will be set as it is the last in the line to be set in this function, and thus will over-write any Superscript settings.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorCharPosition(ByRef $oTextCursor, $bAutoSuper = Null, $iSuperScript = Null, $bAutoSub = Null, $iSubScript = Null, $iRelativeSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oCursor

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oTextCursor.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch __LOCalc_Internal_CursorGetType($oTextCursor)
		Case $LOC_CURTYPE_PARAGRAPH
			; Paragraph Objects normally should behave the same as Text Cursors, like they do in Writer, but they don't in Calc. So I create a Text Cursor temporarily
			; to use in this function that has the Paragraph selected.
			$oCursor = $oTextCursor.Text.createTextCursorByRange($oTextCursor)
			If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case Else
			$oCursor = $oTextCursor
	EndSwitch

	$vReturn = __LOCalc_CharPosition($oCursor, $bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_TextCursorCharPosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorCharSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning) for a Text Cursor.
; Syntax ........: _LOCalc_TextCursorCharSpacing(ByRef $oTextCursor[, $bAutoKerning = Null[, $nKerning = Null]])
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bAutoKerning        - [optional] a boolean value. Default is Null. If True, applies a spacing in between certain pairs of characters.
;                  $nKerning            - [optional] a general number value (-2-928.8). Default is Null. The kerning value of the characters. See Remarks. Values are in Printer's Points as set in the Libre Office UI.
; Return values .: Success: Integer or Array.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor does not support Character properties.
;                  @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bAutoKerning not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $nKerning not a number, or less than -2 or greater than 928.8 Points.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Text Cursor for Paragraph Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;                  |                               1 = Error setting $bAutoKerning
;                  |                               2 = Error setting $nKerning.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving current settings in any Text Cursor formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of inaccurate values.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User Display, however the internal setting is measured in Micrometers. They will be automatically converted from Points to Micrometers and back for retrieval of settings.
;                  The acceptable values are from -2 Pt to 928.8 Pt. The values can be directly converted easily, however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative Micrometers internally from 928.9 up to 1000 Pt (Max setting).
;                  For example, 928.8Pt is the last correct value, which equals 32766 uM (Micrometers), after this LibreOffice reports the following: ;928.9 Pt = -32766 uM; 929 Pt = -32763 uM; 929.1 = -32759; 1000 pt = -30258. Attempting to set Libre's kerning value to anything over 32768 uM causes a COM exception, and attempting to set the kerning to any of these negative numbers sets the User viewable kerning value to -2.0 Pt. For these reasons the max settable kerning is -2.0 Pt to 928.8 Pt.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorCharSpacing(ByRef $oTextCursor, $bAutoKerning = Null, $nKerning = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oCursor

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oTextCursor.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch __LOCalc_Internal_CursorGetType($oTextCursor)
		Case $LOC_CURTYPE_PARAGRAPH
			; Paragraph Objects normally should behave the same as Text Cursors, like they do in Writer, but they don't in Calc. So I create a Text Cursor temporarily
			; to use in this function that has the Paragraph selected.
			$oCursor = $oTextCursor.Text.createTextCursorByRange($oTextCursor)
			If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case Else
			$oCursor = $oTextCursor
	EndSwitch

	$vReturn = __LOCalc_CharSpacing($oCursor, $bAutoKerning, $nKerning)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_TextCursorCharSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorEffect
; Description ...: Set or Retrieve the Font Effect settings for a Text Cursor.
; Syntax ........: _LOCalc_TextCursorEffect(ByRef $oTextCursor[, $iRelief = Null[, $bOutline = Null[, $bShadow = Null]]])
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $iRelief             - [optional] an integer value (0-2). Default is Null. The Character Relief style. See Constants, $LOC_RELIEF_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bOutline            - [optional] a boolean value. Default is Null. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] a boolean value. Default is Null. If True, the characters have a shadow.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor does not support Character properties.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iRelief not an Integer, less than 0 or greater than 2. See Constants, $LOC_RELIEF_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 5 Return 0 = $bOutline not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $bShadow not a Boolean.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Text Cursor for Paragraph Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iRelief
;                  |                               2 = Error setting $bOutline
;                  |                               4 = Error setting $bShadow
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving current settings in any Text Cursor formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of inaccurate values.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorEffect(ByRef $oTextCursor, $iRelief = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oCursor

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oTextCursor.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch __LOCalc_Internal_CursorGetType($oTextCursor)
		Case $LOC_CURTYPE_PARAGRAPH
			; Paragraph Objects normally should behave the same as Text Cursors, like they do in Writer, but they don't in Calc. So I create a Text Cursor temporarily
			; to use in this function that has the Paragraph selected.
			$oCursor = $oTextCursor.Text.createTextCursorByRange($oTextCursor)
			If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case Else
			$oCursor = $oTextCursor
	EndSwitch

	$vReturn = __LOCalc_CellEffect($oCursor, $iRelief, $bOutline, $bShadow)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_TextCursorEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorFont
; Description ...: Set and Retrieve the Font Settings for a Text Cursor.
; Syntax ........: _LOCalc_TextCursorFont(ByRef $oTextCursor[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $sFontName           - [optional] a string value. Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] a general number value. Default is Null. The new Font size.
;                  $iPosture            - [optional] an integer value (0-5). Default is Null. The Font Italic setting. See Constants, $LOC_POSTURE_* as defined in LibreOfficeCalc_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] an integer value(0,50-200). Default is Null. The Font Bold settings see Constants, $LOC_WEIGHT_* as defined in LibreOfficeCalc_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor does not support Character properties.
;                  @Error 1 @Extended 3 Return 0 = Font called in $sFontName not available.
;                  @Error 1 @Extended 4 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 5 Return 0 = $sFontName not a String.
;                  @Error 1 @Extended 6 Return 0 = $nFontSize not a number.
;                  @Error 1 @Extended 7 Return 0 = $iPosture not an Integer, less than 0, or greater than 5. See Constants, $LOC_POSTURE_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 8 Return 0 = $iWeight not an Integer, less than 50 but not equal to 0, or greater than 200. See Constants, $LOC_WEIGHT_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Text Cursor for Paragraph Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $sFontName
;                  |                               2 = Error setting $nFontSize
;                  |                               4 = Error setting $iPosture
;                  |                               8 = Error setting $iWeight
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
;                  Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted, such as oblique, ultra Bold etc.
;                  Libre Calc accepts only the predefined weight values, any other values are changed automatically to an acceptable value, which could trigger a settings error.
; Related .......: _LOCalc_FontsGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorFont(ByRef $oTextCursor, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oCursor

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oTextCursor.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If ($sFontName <> Null) And Not _LOCalc_FontExists($sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	Switch __LOCalc_Internal_CursorGetType($oTextCursor)
		Case $LOC_CURTYPE_PARAGRAPH
			; Paragraph Objects normally should behave the same as Text Cursors, like they do in Writer, but they don't in Calc. So I create a Text Cursor temporarily
			; to use in this function that has the Paragraph selected.
			$oCursor = $oTextCursor.Text.createTextCursorByRange($oTextCursor)
			If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case Else
			$oCursor = $oTextCursor
	EndSwitch

	$vReturn = __LOCalc_CellFont($oCursor, $sFontName, $nFontSize, $iPosture, $iWeight)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_TextCursorFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorFontColor
; Description ...: Set or Retrieve the Font Color for a Text Cursor.
; Syntax ........: _LOCalc_TextCursorFontColor(ByRef $oTextCursor[, $iFontColor = Null])
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $iFontColor          - [optional] an integer value (-1-16777215). Default is Null. The Color value in Long Integer format to make the font, can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for Auto color.
; Return values .: Success: 1 or Integer.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor does not support Character properties.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $iFontColor not an Integer, less than 0, or greater than 16777215.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Text Cursor for Paragraph Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $iFontColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current Font Color as an Integer.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving current settings in any Text Cursor formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of inaccurate values.
;                  Though Transparency is present on the Font Effects page in the UI, there is (as best as I can find) no setting for it available to read and modify. And further, it seems even in L.O. the setting does not affect the font's transparency, though it may change the color value.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
; Related .......: _LO_ConvertColorToLong, _LO_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorFontColor(ByRef $oTextCursor, $iFontColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oCursor

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oTextCursor.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch __LOCalc_Internal_CursorGetType($oTextCursor)
		Case $LOC_CURTYPE_PARAGRAPH
			; Paragraph Objects normally should behave the same as Text Cursors, like they do in Writer, but they don't in Calc. So I create a Text Cursor temporarily
			; to use in this function that has the Paragraph selected.
			$oCursor = $oTextCursor.Text.createTextCursorByRange($oTextCursor)
			If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case Else
			$oCursor = $oTextCursor
	EndSwitch

	$vReturn = __LOCalc_CellFontColor($oCursor, $iFontColor)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_TextCursorFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorGetString
; Description ...: Retrieve the string of text currently selected by a cursor or contained in a paragraph object.
; Syntax ........: _LOCalc_TextCursorGetString(ByRef $oObj)
; Parameters ....: $oObj                - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
; Return values .: Success: String
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oObj not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oObj doesn't support Character Properties service.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Cursor type.
;                  --Success--
;                  @Error 0 @Extended 0 Return String = Success. The selected text in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Libre Office documentation states that when used in Libre Basic, GetString is limited to 64kb's in size. I do not know if the same limitation applies to any outside use of GetString (such as through Autoit).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorGetString(ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oObj.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oObj.getString())
EndFunc   ;==>_LOCalc_TextCursorGetString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorGoToRange
; Description ...: Moves a Text cursor to another Text Cursor or Paragraph portion Position or Range.
; Syntax ........: _LOCalc_TextCursorGoToRange(ByRef $oCursor, ByRef $oRange[, $bSelect = False])
; Parameters ....: $oCursor             - [in/out] an object. an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $oRange              - [in/out] an object. an object. The Cursor or paragraph portion to move cursor called in $oCursor to. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bSelect             - [optional] a boolean value. Default is False. If True, the selection is expanded or created from original cursor location to Range location.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oRange not an Object.
;                  @Error 1 @Extended 3 Return 0 = $bSelect not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $oCursor not a Text Cursor.
;                  @Error 1 @Extended 5 Return 0 = $oRange is a Sheet Cursor and is not supported.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error determining $oCursor cursor type.
;                  @Error 3 @Extended 2 Return 0 = Error determining $oRange cursor type.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Cursor successfully moved to $oRange position.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If the Cursor being used as a range has anything selected, the selection will be selected in the Cursor called in $oCursor also.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorGoToRange(ByRef $oCursor, ByRef $oRange, $bSelect = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType, $iRangeType

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRange) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bSelect) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iCursorType = __LOCalc_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$iRangeType = __LOCalc_Internal_CursorGetType($oRange)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)
	If ($iCursorType <> $LOC_CURTYPE_TEXT_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iRangeType = $LOC_CURTYPE_SHEET_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oCursor.gotoRange($oRange, $bSelect)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_TextCursorGoToRange

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorInsertString
; Description ...: Insert a string using a Text cursor.
; Syntax ........: _LOCalc_TextCursorInsertString(ByRef $oCursor, $sString[, $bOverwrite = False])
; Parameters ....: $oCursor             - [in/out] an object. A Text Cursor Object returned from any Cursor Object creation function.
;                  $sString             - a string value. A String to insert.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, and the cursor object has text selected, the selection is overwritten, else if False, the string is inserted to the left of the selection.
; Return values .: Success: 1
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $sString not a string..
;                  @Error 1 @Extended 3 Return 0 = $bOverwrite not a Boolean.
;                  @Error 1 @Extended 4 Return 0 = $oCursor is not a Text Cursor.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error retrieving Cursor type.
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. String was successfully inserted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_PageStyleHeaderCreateTextCursor, _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_CellCreateTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorInsertString(ByRef $oCursor, $sString, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sString) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$iCursorType = __LOCalc_Internal_CursorGetType($oCursor)
	If @error > 0 Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If ($iCursorType <> $LOC_CURTYPE_TEXT_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oCursor.Text.insertString($oCursor, $sString, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOCalc_TextCursorInsertString

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorIsCollapsed
; Description ...: Retrieve the current status of a Text cursor, whether it has any data selected or not.
; Syntax ........: _LOCalc_TextCursorIsCollapsed(ByRef $oCursor)
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oCursor unknown cursor type.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Failed to retrieve Cursor Type.
;                  --Success--
;                  @Error 0 @Extended 0 Return Boolean = Success. Successfully queried whether cursor's selection is collapsed. Returning Boolean result, True = Cursor has no data selected, False = cursor has data selected.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOCalc_TextCursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorIsCollapsed(ByRef $oCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iCursorType
	Local $bReturn

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCursorType = __LOCalc_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Switch $iCursorType
		Case $LOC_CURTYPE_TEXT_CURSOR
			$bReturn = $oCursor.isCollapsed()

			Return SetError($__LO_STATUS_SUCCESS, 0, $bReturn)

		Case Else

			Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0) ; unknown or wrong cursor data type.
	EndSwitch
EndFunc   ;==>_LOCalc_TextCursorIsCollapsed

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorMove
; Description ...: Move a Text Cursor object in a document. Also for creating/Expanding selections.
; Syntax ........: _LOCalc_TextCursorMove(ByRef $oCursor, $iMove[, $iCount = 1[, $bSelect = False]])
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval functions.
;                  $iMove               - an integer value. The movement command. See remarks and Constants, $LOC_TEXTCUR_* as defined in LibreOfficeCalc_Constants.au3.
;                  $iCount              - [optional] an integer value. Default is 1. Number of movements to make. See remarks.
;                  $bSelect             - [optional] a boolean value. Default is False. Whether to select data during this cursor movement. See remarks.
; Return values .: Success: Boolean.
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $iMove not an Integer.
;                  @Error 1 @Extended 3 Return 0 = $iMove less than 0 or greater than highest move Constant. See Move Constants, $LOC_TEXTCUR_* as defined in LibreOfficeCalc_Constants.au3.
;                  @Error 1 @Extended 4 Return 0 = $iCount not an integer or is a negative.
;                  @Error 1 @Extended 5 Return 0 = $bSelect not a Boolean.
;                  --Processing Errors--
;                  @Error 3 @Extended 1 Return 0 = Error determining cursor type.
;                  @Error 3 @Extended 2 Return 0 = Error processing cursor move.
;                  @Error 3 @Extended 3 Return 0 = $oCursor Object unknown cursor type.
;                  --Success--
;                  @Error 0 @Extended ? Return Boolean = Success, Cursor object movement was processed successfully. Returns True if the full count of movements were successful, else false if none or only partially successful. @Extended set to number of successful movements. See Remarks
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Only some movements accept movement amounts and selecting (such as $LOC_TEXTCUR_GO_RIGHT 2, True) etc. Also only some accept creating/ extending a selection of text/ data. They will be specified below.
;                  To Clear /Unselect a current selection, you can input a move such as $LOC_TEXTCUR_GO_RIGHT, 0, False.
;                  #Cursor Movement Constants which accept Number of Moves and Selecting:
;                  $LOC_TEXTCUR_GO_LEFT,
;                  $LOC_TEXTCUR_GO_RIGHT,
;                  #Cursor Movements which accept Selecting Only:
;                  $LOC_TEXTCUR_GOTO_START,
;                  $LOC_TEXTCUR_GOTO_END,
;                  #Cursor Movements which accept nothing and are done once per call:
;                  $LOC_TEXTCUR_COLLAPSE_TO_START,
;                  $LOC_TEXTCUR_COLLAPSE_TO_END
; Related .......: _LOCalc_TextCursorIsCollapsed
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorMove(ByRef $oCursor, $iMove, $iCount = 1, $bSelect = False)
	Local $iCursorType
	Local $bMoved = False

	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$iCursorType = __LOCalc_Internal_CursorGetType($oCursor)
	If @error Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Switch $iCursorType
		Case $LOC_CURTYPE_TEXT_CURSOR
			$bMoved = __LOCalc_TextCursorMove($oCursor, $iMove, $iCount, $bSelect)

			Return SetError(@error, @extended, $bMoved)

		Case Else

			Return SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0) ; unknown or wrong cursor type.
	EndSwitch
EndFunc   ;==>_LOCalc_TextCursorMove

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorOverline
; Description ...: Set and retrieve the Overline settings for a Text Cursor.
; Syntax ........: _LOCalc_TextCursorOverline(ByRef $oTextCursor[, $bWordOnly = Null[, $iOverLineStyle = Null[, $bOLHasColor = Null[, $iOLColor = Null]]]])
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not Overlined.
;                  $iOverLineStyle      - [optional] an integer value (0-18). Default is Null. The style of the Overline line, see constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3. See Remarks.
;                  $bOLHasColor         - [optional] a boolean value. Default is Null. If True, the Overline is colored, must be set to true in order to set the Overline color.
;                  $iOLColor            - [optional] an integer value (-1-16777215). Default is Null. The Overline color, set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor does not support Character properties.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iOverLineStyle not an Integer, less than 0, or greater than 18. See constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3. See Remarks.
;                  @Error 1 @Extended 6 Return 0 = $bOLHasColor not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iOLColor not an Integer, less than -1, or greater than 16777215.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Text Cursor for Paragraph Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $iOverLineStyle
;                  |                               4 = Error setting $bOLHasColor
;                  |                               8 = Error setting $iOLColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving current settings in any Text Cursor formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of inaccurate values.
;                  Overline line style uses the same constants as underline style.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorToLong, _LO_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorOverline(ByRef $oTextCursor, $bWordOnly = Null, $iOverLineStyle = Null, $bOLHasColor = Null, $iOLColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oCursor

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oTextCursor.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch __LOCalc_Internal_CursorGetType($oTextCursor)
		Case $LOC_CURTYPE_PARAGRAPH
			; Paragraph Objects normally should behave the same as Text Cursors, like they do in Writer, but they don't in Calc. So I create a Text Cursor temporarily
			; to use in this function that has the Paragraph selected.
			$oCursor = $oTextCursor.Text.createTextCursorByRange($oTextCursor)
			If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case Else
			$oCursor = $oTextCursor
	EndSwitch

	$vReturn = __LOCalc_CellOverLine($oCursor, $bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_TextCursorOverline

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorParObjCreateList
; Description ...: Return Objects for every paragraph contained in a specific section of a document.
; Syntax ........: _LOCalc_TextCursorParObjCreateList(ByRef $oTextCursor)
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
; Return values .: Success: 1D Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Enumeration of Paragraphs.
;                  --Success--
;                  @Error 0 @Extended ? Return Array = Success. Returns an Array of Paragraph Objects, @Extended is set to the number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The paragraphs are enumerated for the area the cursor is currently within, for example, the Text Cursor is currently in a Cell, the enumeration of paragraphs would be for the Cell the cursor was presently in.
;                  Returns an Array of objects for Directly Formatting paragraphs in a document, or for deleting or inserting in other areas, etc.
; Related .......: _LOCalc_TextCursorParObjSectionsGet
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorParObjCreateList(ByRef $oTextCursor)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oEnum, $oPar
	Local $iCount = 0
	Local $aoParagraphs[1]

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oEnum = $oTextCursor.Text.createEnumeration()
	If Not IsObj($oEnum) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	While $oEnum.hasMoreElements()
		$oPar = $oEnum.nextElement()

		If UBound($aoParagraphs) <= ($iCount) Then ReDim $aoParagraphs[UBound($aoParagraphs) * 2]
		$aoParagraphs[$iCount] = $oPar
		$iCount += 1

		Sleep((IsInt($iCount / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	WEnd

	ReDim $aoParagraphs[$iCount]

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoParagraphs)
EndFunc   ;==>_LOCalc_TextCursorParObjCreateList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorParObjSectionsGet
; Description ...: Break a Paragraph Object into individual Sections for Formatting etc. See Remarks.
; Syntax ........: _LOCalc_TextCursorParObjSectionsGet(ByRef $oParObj)
; Parameters ....: $oParObj             - [in/out] an object. A Paragraph Object returned from _LOCalc_TextCursorParObjCreateList function.
; Return values .: Success: Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oParObj is not an Object.
;                  @Error 1 @Extended 2 Return 0 = Object called in $oParObj not a paragraph Object.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Error enumerating Paragraph sections.
;                  --Success--
;                  @Error 0 @Extended 0 Return Array = Success. A two column array. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Paragraph in a Document may have more than one section if it contains different formatting, etc.
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
; Related .......: _LOCalc_TextCursorParObjCreateList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorParObjSectionsGet(ByRef $oParObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSecEnum, $oParSection
	Local $aoSections[1][2]
	Local $iCount = 0

	If Not IsObj($oParObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If ($oParObj.ImplementationName() <> "SvxUnoTextContent") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oSecEnum = $oParObj.createEnumeration()
	If Not IsObj($oSecEnum) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	While $oSecEnum.hasMoreElements()
		$oParSection = $oSecEnum.nextElement()

		If UBound($aoSections) <= ($iCount + 1) Then ReDim $aoSections[UBound($aoSections) * 10][2]
		$aoSections[$iCount][0] = $oParSection
		$aoSections[$iCount][1] = $oParSection.TextPortionType
		$iCount += 1
		Sleep((IsInt($iCount / $__LOCCONST_SLEEP_DIV) ? (10) : (0)))
	WEnd
	ReDim $aoSections[$iCount][2]

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoSections)
EndFunc   ;==>_LOCalc_TextCursorParObjSectionsGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorStrikeOut
; Description ...: Set or Retrieve the Strikeout settings for a Text Cursor.
; Syntax ........: _LOCalc_TextCursorStrikeOut(ByRef $oTextCursor[, $bWordOnly = Null[, $bStrikeOut = Null[, $iStrikeLineStyle = Null]]])
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If True, strike out is applied to words only, skipping whitespaces.
;                  $bStrikeOut          - [optional] a boolean value. Default is Null. If True, strikeout is applied to characters.
;                  $iStrikeLineStyle    - [optional] an integer value (0-6). Default is Null. The Strikeout Line Style, see constants, $LOC_STRIKEOUT_* as defined in LibreOfficeCalc_Constants.au3.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor does not support Character properties.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $bStrikeOut not a Boolean.
;                  @Error 1 @Extended 6 Return 0 = $iStrikeLineStyle not an Integer, less than 0 or greater than 6. See constants, $LOC_STRIKEOUT_* as defined in LibreOfficeCalc_Constants.au3.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Text Cursor for Paragraph Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $bStrikeOut
;                  |                               4 = Error setting $iStrikeLineStyle
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving current settings in any Text Cursor formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of inaccurate values.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorStrikeOut(ByRef $oTextCursor, $bWordOnly = Null, $bStrikeOut = Null, $iStrikeLineStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oCursor

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oTextCursor.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch __LOCalc_Internal_CursorGetType($oTextCursor)
		Case $LOC_CURTYPE_PARAGRAPH
			; Paragraph Objects normally should behave the same as Text Cursors, like they do in Writer, but they don't in Calc. So I create a Text Cursor temporarily
			; to use in this function that has the Paragraph selected.
			$oCursor = $oTextCursor.Text.createTextCursorByRange($oTextCursor)
			If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case Else
			$oCursor = $oTextCursor
	EndSwitch

	$vReturn = __LOCalc_CellStrikeOut($oCursor, $bWordOnly, $bStrikeOut, $iStrikeLineStyle)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_TextCursorStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOCalc_TextCursorUnderline
; Description ...: Set and retrieve the Underline settings for a Text Cursor.
; Syntax ........: _LOCalc_TextCursorUnderline(ByRef $oTextCursor[, $bWordOnly = Null[, $iUnderLineStyle = Null[, $bULHasColor = Null[, $iULColor = Null]]]])
; Parameters ....: $oTextCursor         - [in/out] an object. A Text Cursor Object returned by a previous _LOCalc_PageStyleFooterCreateTextCursor, _LOCalc_PageStyleHeaderCreateTextCursor, or _LOCalc_CellCreateTextCursor function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not underlined.
;                  $iUnderLineStyle     - [optional] an integer value (0-18). Default is Null. The Underline line style, see constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3.
;                  $bULHasColor         - [optional] a boolean value. Default is Null. If True, the underline is colored, must be set to true in order to set the underline color.
;                  $iULColor            - [optional] an integer value (-1-16777215). Default is Null. The color of the underline, set in Long integer format. Can be a custom value, or one of the constants, $LO_COLOR_* as defined in LibreOffice_Constants.au3. Set to $LO_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array
;                  Failure: 0 and sets the @Error and @Extended flags to non-zero.
;                  --Input Errors--
;                  @Error 1 @Extended 1 Return 0 = $oTextCursor not an Object.
;                  @Error 1 @Extended 2 Return 0 = $oTextCursor does not support Character properties.
;                  @Error 1 @Extended 3 Return 0 = Variable passed to internal function not an Object.
;                  @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;                  @Error 1 @Extended 5 Return 0 = $iUnderLineStyle not an Integer, less than 0, or greater than 18. See constants, $LOC_UNDERLINE_* as defined in LibreOfficeCalc_Constants.au3. See Remarks.
;                  @Error 1 @Extended 6 Return 0 = $bULHasColor not a Boolean.
;                  @Error 1 @Extended 7 Return 0 = $iULColor not an Integer, less than -1, or greater than 16777215.
;                  --Initialization Errors--
;                  @Error 2 @Extended 1 Return 0 = Failed to create Text Cursor for Paragraph Object.
;                  --Property Setting Errors--
;                  @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for following values:
;                  |                               1 = Error setting $bWordOnly
;                  |                               2 = Error setting $iUnderLineStyle
;                  |                               4 = Error setting $bULHasColor
;                  |                               8 = Error setting $iULColor
;                  --Success--
;                  @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;                  @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Retrieving current settings in any Text Cursor formatting functions may be inaccurate as multiple different settings could be selected at once, which would result in a return of inaccurate values.
;                  Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;                  Call any optional parameter with Null keyword to skip it.
; Related .......: _LO_ConvertColorToLong, _LO_ConvertColorFromLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOCalc_TextCursorUnderline(ByRef $oTextCursor, $bWordOnly = Null, $iUnderLineStyle = Null, $bULHasColor = Null, $iULColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOCalc_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn
	Local $oCursor

	If Not IsObj($oTextCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oTextCursor.supportsService("com.sun.star.style.CharacterProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	Switch __LOCalc_Internal_CursorGetType($oTextCursor)
		Case $LOC_CURTYPE_PARAGRAPH
			; Paragraph Objects normally should behave the same as Text Cursors, like they do in Writer, but they don't in Calc. So I create a Text Cursor temporarily
			; to use in this function that has the Paragraph selected.
			$oCursor = $oTextCursor.Text.createTextCursorByRange($oTextCursor)
			If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		Case Else
			$oCursor = $oTextCursor
	EndSwitch

	$vReturn = __LOCalc_CellUnderLine($oCursor, $bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOCalc_TextCursorUnderline

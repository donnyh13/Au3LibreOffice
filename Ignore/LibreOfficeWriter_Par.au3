#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Writer
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; Other includes for Writer
#include "LibreOfficeWriter_Cursor.au3"
#include "LibreOfficeWriter_Doc.au3"
#include "LibreOfficeWriter_Num.au3"
#include "LibreOfficeWriter_Page.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Creating, Modifying, and Applying L.O. Writer Paragraph Styles.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_ParObjCopy
; _LOWriter_ParObjCreateList
; _LOWriter_ParObjDelete
; _LOWriter_ParObjPaste
; _LOWriter_ParObjSectionsGet
; _LOWriter_ParObjSelect
; _LOWriter_ParStyleAlignment
; _LOWriter_ParStyleBackColor
; _LOWriter_ParStyleBorderColor
; _LOWriter_ParStyleBorderPadding
; _LOWriter_ParStyleBorderStyle
; _LOWriter_ParStyleBorderWidth
; _LOWriter_ParStyleCreate
; _LOWriter_ParStyleDelete
; _LOWriter_ParStyleDropCaps
; _LOWriter_ParStyleEffect
; _LOWriter_ParStyleExists
; _LOWriter_ParStyleFont
; _LOWriter_ParStyleFontColor
; _LOWriter_ParStyleGetObj
; _LOWriter_ParStyleHyphenation
; _LOWriter_ParStyleIndent
; _LOWriter_ParStyleOrganizer
; _LOWriter_ParStyleOutLineAndList
; _LOWriter_ParStyleOverLine
; _LOWriter_ParStylePageBreak
; _LOWriter_ParStylePosition
; _LOWriter_ParStyleRotateScale
; _LOWriter_ParStyleSet
; _LOWriter_ParStylesGetNames
; _LOWriter_ParStyleShadow
; _LOWriter_ParStyleSpace
; _LOWriter_ParStyleSpacing
; _LOWriter_ParStyleStrikeOut
; _LOWriter_ParStyleTabStopCreate
; _LOWriter_ParStyleTabStopDelete
; _LOWriter_ParStyleTabStopList
; _LOWriter_ParStyleTabStopMod
; _LOWriter_ParStyleTxtFlowOpt
; _LOWriter_ParStyleUnderLine
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParObjCopy
; Description ...: "Copies" data selected by the ViewCursor, returning an Object for use in inserting later.
; Syntax ........: _LOWriter_ParObjCopy(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to Copy Selected Data, make sure Data is selected using the ViewCursor.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object  = Success. Data was successfully copied, returning an Object for use in _LOWriter_ParObjPaste.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Data you desire to be copied MUST be selected with the ViewCursor, see _LOWriter_ParObjSelect.
;				   This function works essentially the same as Copy/ Ctrl+C, except it doesn't use your clipboard.
;				   The Object returned is used in _LOWriter_ParObjPaste to insert the data again.
;				   Copying data this way works for Tables, Images, frames and Text, including with direct formatting, etc.
;				   Data copied can be inserted into the same or another document.
; Related .......: _LOWriter_ParObjPaste, _LOWriter_ParObjSelect, _LOWriter_DocGetViewCursor, _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParObjCopy(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oObj

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oObj = $oDoc.CurrentController.getTransferable() ; Copy
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oObj)
EndFunc   ;==>_LOWriter_ParObjCopy

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParObjCreateList
; Description ...: Return Objects for every paragraph contained in a specific section of a document.
; Syntax ........: _LOWriter_ParObjCreateList(ByRef $oCursor[, $bTableCheck = False])
; Parameters ....: $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval functions. See Remarks
;                  $bTableCheck         - [optional] a boolean value. Default is False. If True, returned array will be 2 dimensional, with the second column indicating if the paragraph object is a Table (True) or not (False).
; Return values .: Success: 1D or 2D Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bTableCheck not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create Enumeration of Paragraphs.
;				   --Success--
;				   @Error 0 @Extended ? Return Array  = Success. Returns an Array of Paragraph Objects, @Extended is set to the number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: 	$oCursor can be either a ViewCursor or a TextCursor, the paragraphs are enumerated for the area the cursor is currently within,
;						for example, the ViewCursor is currently in a Table, the enumeration of paragraphs would be for the Cell the cursor was presently in.
;					In the main document the enumeration would be for the entire Text Body, in the Header, it would for the that Header for that Page Style etc.
;					The different possible areas are: Text Body, Table Cell, Header, Footer, Footnote, Endnote, Frame.
;					Returns an Array of objects for Direct Formatting paragraphs in a document, or for copying and inserting etc.
;					Table Objects returned from this function can be used as a regular Table Object to modify the Table with.
; Related .......: _LOWriter_ParObjSectionsGet, _LOWriter_ParObjSelect, _LOWriter_ParObjDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParObjCreateList(ByRef $oCursor, $bTableCheck = False)
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
EndFunc   ;==>_LOWriter_ParObjCreateList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParObjDelete
; Description ...: Delete a Paragraph Object returned from _LOWriter_ParObjCreateList. See Remarks.
; Syntax ........: _LOWriter_ParObjDelete(ByRef $oParObj)
; Parameters ....: $oParObj             - [in/out] an object. A Paragraph Object returned by _LOWriter_ParObjCreateList.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParObj not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1  = Success. Paragraph was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: You cannot delete the last paragraph contained in a Text area, it will cause a COM error.
; Related .......: _LOWriter_ParObjCreateList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParObjDelete(ByRef $oParObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oParObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oParObj.supportsService("com.sun.star.text.TextTable")) Then
		$oParObj.dispose()

	Else
		$oParObj.Text.removeTextContent($oParObj)

	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_ParObjDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParObjPaste
; Description ...: Inserts a ParObjCopy Object at the current ViewCursor location.
; Syntax ........: _LOWriter_ParObjPaste(ByRef $oDoc, ByRef $oParObj)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oParObj             - [in/out] an object. A Object returned from _LOWriter_ParObjCopy to insert.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParObj not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1  = Success. Data was successfully inserted at the ViewCursor location.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ParObjCopy
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParObjPaste(ByRef $oDoc, ByRef $oParObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oParObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.insertTransferable($oParObj)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_ParObjPaste

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParObjSectionsGet
; Description ...: Break a Paragraph Object into individual Sections for Direct Formatting etc. See Remarks.
; Syntax ........: _LOWriter_ParObjSectionsGet(ByRef $oParagraph)
; Parameters ....: $oParagraph          - [in/out] an object. A Paragraph Object returned from _LOWriter_ParObjCreateList function. Make sure it's not a Table!
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParagraph is not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParagraph does not support Paragraph service -- Not a paragraph Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error enumerating Paragraph sections.
;				   --Success--
;				   @Error 0 @Extended 0 Return Array = Success. A two column array. See remarks.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Paragraph in a Document may have more than one section if it contains direct formatting, foot/endnote anchors etc.
;				   The Array returned is a two column array with array[0][0] containing the section Object.
;				   The second column, array[0][1] contains the section data type column being one of the following possible types:
;				   |- Text – String content.
;				   |- TextField – TextField content.
;				   |- TextContent – Indicates that text content is anchored as or to a character that is not really part of the paragraph — for example, a text frame or a graphic object.
;				   |- ControlCharacter – Control character.
;				   |- Footnote – Footnote or Endnote. (Note this is just the anchor character for the footnote/Endnote, not the actual foot/endnote content.
;				   |- ReferenceMark – Reference mark.
;				   |- DocumentIndexMark – Document index mark.
;				   |- Bookmark – Bookmark.
;				   |- Redline – Redline portion, which is a result of the change-tracking feature.
;				   |- Ruby – a ruby attribute which is used in Asian text.
;				   |- Frame — a frame.
;				   |- SoftPageBreak — a soft page break.
;				   |- InContentMetadata — a text range with attached metadata.
;				   Note: For Reference marks, document index marks, etc., 2 text portions will be generated, one for the start position and one for the end position.
; Related .......: _LOWriter_ParObjCreateList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParObjSectionsGet(ByRef $oParagraph)
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
		$aoSections[$iCount][1] = $oParSection.TextPortionType
		$iCount += 1
		Sleep((IsInt($iCount / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	WEnd
	ReDim $aoSections[$iCount][2]
	Return SetError($__LO_STATUS_SUCCESS, $iCount, $aoSections)
EndFunc   ;==>_LOWriter_ParObjSectionsGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParObjSelect
; Description ...: Causes a Paragraph Object to be selected by the ViewCursor.
; Syntax ........: _LOWriter_ParObjSelect(ByRef $oDoc, ByRef $oObj)
; Parameters ....: $oDoc             - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oObj             - [in/out] an object. A Paragraph Object returned from _LOWriter_ParObjCreateList, a Table or Frame Object, or a Text Cursor with data selected, can be used.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve ViewCursor Object.
;				   @Error 3 @Extended 2 Return 0 = Failed to move ViewCursor.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1  = Success. Object was successfully selected.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function causes the ViewCursor to move and select a Paragraph, Table, Frame, TextCursor data, etc., usually in preparation for calling _LOWriter_ParObjCopy.
; Related .......: _LOWriter_ParObjCreateList, _LOWriter_ParObjCopy, _LOWriter_TableGetObjByName, _LOWriter_TableGetObjByCursor,
;					_LOWriter_TableInsert, _LOWriter_FrameGetObjByCursor, _LOWriter_FrameGetObjByName,
;					_LOWriter_DocGetViewCursor,	_LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor,	_LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParObjSelect(ByRef $oDoc, ByRef $oObj)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oViewCursor

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oDoc.CurrentController.Select($oObj)

	If ($oObj.supportsService("com.sun.star.text.TextTable")) Then

		$oViewCursor = _LOWriter_DocGetViewCursor($oDoc)
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_END, 1, True) ; Move and select to End of cell
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

		_LOWriter_CursorMove($oViewCursor, $LOW_VIEWCUR_GOTO_END, 1, True) ; Move and select to End of Table
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0)

	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_ParObjSelect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleAlignment
; Description ...: Set and Retrieve Alignment settings for a paragraph style.
; Syntax ........: _LOWriter_ParStyleAlignment(ByRef $oParStyle[, $iHorAlign = Null[, $iVertAlign = Null[, $iLastLineAlign = Null[, $bExpandSingleWord = Null[, $bSnapToGrid = Null[, $iTxtDirection = Null]]]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iHorAlign           - [optional] an integer value (0-3). Default is Null. The Horizontal alignment of the paragraph. See Constants, $LOW_PAR_ALIGN_HOR_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $iVertAlign          - [optional] an integer value (0-4). Default is Null. The Vertical alignment of the paragraph. See Constants, $LOW_PAR_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLastLineAlign      - [optional] an integer value (0-3). Default is Null. The last line alignment for the paragraph. See Constants, $LOW_PAR_LAST_LINE_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $bExpandSingleWord   - [optional] a boolean value. Default is Null. If True, and the last line of a justified paragraph consists of one word, the word is stretched to the width of the paragraph.
;                  $bSnapToGrid         - [optional] a boolean value. Default is Null. If True, aligns the paragraph to a text grid (if one is active).
;                  $iTxtDirection       - [optional] an integer value (0-5). Default is Null. The Text Writing Direction. See Constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3. [Libre Office Default is 4]
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iHorAlign not an integer, less than 0, or greater than 3. See constants, $LOW_PAR_ALIGN_HOR_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 5 Return 0 = $iVertAlign not an integer, less than 0, or greater than 4. See constants, $LOW_PAR_ALIGN_VERT_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $iLastLineAlign not an integer, less than 0, or greater than 3. See constants, $LOW_PAR_LAST_LINE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 7 Return 0 = $bExpandSingleWord not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $bSnapToGrid not a Boolean.
;				   @Error 1 @Extended 9 Return 0 = $iTxtDirection not an Integer, less than 0, or greater than 5, see constants, $LOW_TXT_DIR_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iHorAlign
;				   |								2 = Error setting $iVertAlign
;				   |								4 = Error setting $iLastLineALign
;				   |								8 = Error setting $bExpandSIngleWord
;				   |								16 = Error setting $bSnapToGrid
;				   |								32 = Error setting $iTxtDirection
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iHorAlign must be set to $LOW_PAR_ALIGN_HOR_JUSTIFIED(2) before you can set $iLastLineAlign, and
;					$iLastLineAlign must be set to $LOW_PAR_LAST_LINE_JUSTIFIED(2) before $bExpandSingleWord can be set.
;					Note: $iTxtDirection constants 2,3, and 5 may not be available depending on your language settings.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleAlignment(ByRef $oParStyle, $iHorAlign = Null, $iVertAlign = Null, $iLastLineAlign = Null, $bExpandSingleWord = Null, $bSnapToGrid = Null, $iTxtDirection = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_ParAlignment($oParStyle, $iHorAlign, $iVertAlign, $iLastLineAlign, $bExpandSingleWord, $bSnapToGrid, $iTxtDirection)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleAlignment

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleBackColor
; Description ...: Set or Retrieve background color settings for a Paragraph style.
; Syntax ........: _LOWriter_ParStyleBackColor(ByRef $oDoc, $sParStyle[, $iBackColor = Null[, $bBackTransparent = Null]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iBackColor          - [optional] an integer value (-1-16777215). Default is Null. The background color. Set in Long integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1), to turn Background color off.
;                  $bBackTransparent    - [optional] a boolean value. Default is Null. If True, the background color is transparent.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iBackColor not an integer, less than -1, or greater than 16777215.
;				   @Error 1 @Extended 5 Return 0 = $bBackTransparent not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iBackColor
;				   |								2 = Error setting $bBackTransparent
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertColorFromLong,
;					_LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleBackColor(ByRef $oParStyle, $iBackColor = Null, $bBackTransparent = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_ParBackColor($oParStyle, $iBackColor, $bBackTransparent)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleBackColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleBorderColor
; Description ...: Set and Retrieve the Paragraph Style Border Line Color. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_ParStyleBorderColor(ByRef $oParStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iTop                - [optional] an integer value (0-16777215). Default is Null. Set the Top Border Line Color of the Paragraph Style in Long Color code format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value (0-16777215). Default is Null. Set the Bottom Border Line Color of the Paragraph Style in Long Color code format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value (0-16777215). Default is Null. Set the Left Border Line Color of the Paragraph Style in Long Color code format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value (0-16777215). Default is Null. Set the Right Border Line Color of the Paragraph Style in Long Color code format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
; Internal Remark: Certain Error values are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0, or greater than 16,777,215.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0, or greater than 16,777,215.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0, or greater than 16,777,215.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0, or greater than 16,777,215.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Cannot set Top Border Color when Top Border width not set.
;				   @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Color when Bottom Border width not set.
;				   @Error 4 @Extended 3 Return 0 = Cannot set Left Border Color when Left Border width not set.
;				   @Error 4 @Extended 4 Return 0 = Cannot set Right Border Color when Right Border width not set.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1  = Success.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertColorFromLong,
;					_LOWriter_ConvertColorToLong, _LOWriter_ParStyleBorderWidth, _LOWriter_ParStyleBorderStyle,
;					_LOWriter_ParStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleBorderColor(ByRef $oParStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_COLOR_BLACK, $LOW_COLOR_WHITE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oParStyle, False, False, True, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleBorderColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleBorderPadding
; Description ...: Set or retrieve the Border Padding (spacing between the Paragraph and border) settings.
; Syntax ........: _LOWriter_ParStyleBorderPadding(ByRef $oParStyle[, $iAll = Null[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iAll                - [optional] an integer value. Default is Null. Set all four padding distances to one distance in Micrometers (uM).
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Distance between the Border and Paragraph in Micrometers(uM).
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Distance between the Border and Paragraph in Micrometers(uM).
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Distance between the Border and Paragraph in Micrometers(uM).
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Distance between the Border and Paragraph in Micrometers(uM).
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iAll not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iTop not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $iBottom not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $Left not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iRight not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iAll border distance
;				   |								2 = Error setting $iTop border distance
;				   |								4 = Error setting $iBottom border distance
;				   |								8 = Error setting $iLeft border distance
;				   |								16 = Error setting $iRight border distance
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer, _LOWriter_ParStyleBorderWidth, _LOWriter_ParStyleBorderStyle,
;					_LOWriter_ParStyleBorderColor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleBorderPadding(ByRef $oParStyle, $iAll = Null, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_ParBorderPadding($oParStyle, $iAll, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleBorderPadding

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleBorderStyle
; Description ...: Set and retrieve the Paragraph Style Border Line style. Libre Office Version 3.4 and Up.
; Syntax ........: _LOWriter_ParStyleBorderStyle(ByRef $oParStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iTop                - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Top Border Line Style of the Paragraph Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iBottom             - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Bottom Border Line Style of the Paragraph Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLeft               - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Left Border Line Style of the Paragraph Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iRight              - [optional] an integer value (0x7FFF,0-17). Default is Null. Set the Right Border Line Style of the Paragraph Style using one of the line style constants, $LOW_BORDERSTYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Internal Remark: Certain Error values are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to higher than 17, and not equal to 0x7FFF, or less than 0.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Cannot set Top Border Style when Top Border width not set.
;				   @Error 4 @Extended 2 Return 0 = Cannot set Bottom Border Style when Bottom Border width not set.
;				   @Error 4 @Extended 3 Return 0 = Cannot set Left Border Style when Left Border width not set.
;				   @Error 4 @Extended 4 Return 0 = Cannot set Right Border Style when Right Border width not set.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1  = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Border Width must be set first to be able to set Border Style and Color.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ParStyleBorderWidth,
;					_LOWriter_ParStyleBorderColor, _LOWriter_ParStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleBorderStyle(ByRef $oParStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, $LOW_BORDERSTYLE_SOLID, $LOW_BORDERSTYLE_DASH_DOT_DOT, "", $LOW_BORDERSTYLE_NONE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$vReturn = __LOWriter_Border($oParStyle, False, True, False, $iTop, $iBottom, $iLeft, $iRight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleBorderStyle

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleBorderWidth
; Description ...: Set and retrieve the Paragraph Style Border Line Width, or the Paragraph Style Connect Border option.
; Syntax ........: _LOWriter_ParStyleBorderWidth(ByRef $oParStyle[, $iTop = Null[, $iBottom = Null[, $iLeft = Null[, $iRight = Null[, $bConnectBorder = Null]]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iTop                - [optional] an integer value. Default is Null. Set the Top Border Line width of the Paragraph Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3. Libre Office Version 3.4 and Up.
;                  $iBottom             - [optional] an integer value. Default is Null. Set the Bottom Border Line Width of the Paragraph Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3. Libre Office Version 3.4 and Up.
;                  $iLeft               - [optional] an integer value. Default is Null. Set the Left Border Line width of the Paragraph Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3. Libre Office Version 3.4 and Up.
;                  $iRight              - [optional] an integer value. Default is Null. Set the Right Border Line Width of the Paragraph Style in Micrometers. Can be a custom value, or one of the constants, $LOW_BORDERWIDTH_* as defined in LibreOfficeWriter_Constants.au3. Libre Office Version 3.4 and Up.
;                  $bConnectBorder      - [optional] a boolean value. Default is Null. If True, borders set for a paragraph are merged with the next paragraph. Note: Borders are only merged if they are identical. Libre Office Version 3.4 and Up.
; Internal Remark: Certain Error values are passed from the internal border setting function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = $iTop not an integer, or set to less than 0.
;				   @Error 1 @Extended 4 Return 0 = $iBottom not an integer, or set to less than 0.
;				   @Error 1 @Extended 5 Return 0 = $iLeft not an integer, or set to less than 0.
;				   @Error 1 @Extended 6 Return 0 = $iRight not an integer, or set to less than 0.
;				   @Error 1 @Extended 7 Return 0 = $bConnectBorder not a Boolean and Not set to Null.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Creating Object "com.sun.star.table.BorderLine2"
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Internal command error. More than one set to True. UDF Must be fixed.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1  = Success.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: To "Turn Off" Borders, set them to 0
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer, _LOWriter_ParStyleBorderStyle, _LOWriter_ParStyleBorderColor,
;					_LOWriter_ParStyleBorderPadding
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleBorderWidth(ByRef $oParStyle, $iTop = Null, $iBottom = Null, $iLeft = Null, $iRight = Null, $bConnectBorder = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If ($iTop <> Null) And Not __LOWriter_IntIsBetween($iTop, 0, $iTop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($iBottom <> Null) And Not __LOWriter_IntIsBetween($iBottom, 0, $iBottom) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If ($iLeft <> Null) And Not __LOWriter_IntIsBetween($iLeft, 0, $iLeft) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($iRight <> Null) And Not __LOWriter_IntIsBetween($iRight, 0, $iRight) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	If ($bConnectBorder <> Null) And Not IsBool($bConnectBorder) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)

	If __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight, $bConnectBorder) Then
		$vReturn = __LOWriter_Border($oParStyle, True, False, False, $iTop, $iBottom, $iLeft, $iRight)
		__LOWriter_AddTo1DArray($vReturn, $oParStyle.ParaIsConnectBorder())
		Return SetError($__LO_STATUS_SUCCESS, 1, $vReturn)
	ElseIf Not __LOWriter_VarsAreNull($iTop, $iBottom, $iLeft, $iRight) Then
		$vReturn = __LOWriter_Border($oParStyle, True, False, False, $iTop, $iBottom, $iLeft, $iRight)
		If @error Then Return SetError(@error, @extended, $vReturn)
	EndIf
	If ($bConnectBorder <> Null) Then $oParStyle.ParaIsConnectBorder = $bConnectBorder

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_ParStyleBorderWidth

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleCreate
; Description ...: Create a new Paragraph Style in a Document.
; Syntax ........: _LOWriter_ParStyleCreate(ByRef $oDoc, $sParStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sParStyle           - a string value. The Name of the new Paragraph Style to Create.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sParStyle not a String.
;				   @Error 1 @Extended 3 Return 0 = Paragraph Style name called in $sParStyle already exists in document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error Retrieving "ParagraphStyle" Object.
;				   @Error 2 @Extended 2 Return 0 = Error Creating "com.sun.star.style.ParagraphStyle" Object.
;				   @Error 2 @Extended 3 Return 0 = Error Retrieving Created Paragraph Style Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error creating new Paragraph Style by name.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. New paragraph Style successfully created. Returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ParStyleDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleCreate(ByRef $oDoc, $sParStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oParStyles, $oStyle, $oParStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$oParStyles = $oDoc.StyleFamilies().getByName("ParagraphStyles")
	If Not IsObj($oParStyles) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	If _LOWriter_ParStyleExists($oDoc, $sParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	$oStyle = $oDoc.createInstance("com.sun.star.style.ParagraphStyle")
	If Not IsObj($oStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oParStyles.insertByName($sParStyle, $oStyle)

	If Not $oParStyles.hasByName($sParStyle) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oParStyle = $oParStyles.getByName($sParStyle)
	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 3, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oParStyle)
EndFunc   ;==>_LOWriter_ParStyleCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleDelete
; Description ...: Delete a User-Created Paragraph Style from a Document.
; Syntax ........: _LOWriter_ParStyleDelete(ByRef $oDoc, $oParStyle[, $bForceDelete = False[, $sReplacementStyle = "Default Paragraph Style"]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function. Must be a User-Created Style, not a built-in Style native to LibreOffice.
;                  $bForceDelete        - [optional] a boolean value. Default is False. If True, Paragraph style will be deleted regardless of whether it is in use or not.
;                  $sReplacementStyle   - [optional] a string value. Default is "Default Paragraph Style". The Paragraph style to use instead of the one being deleted if the paragraph style being deleted is applied to text in the document.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 4 Return 0 = $bForceDelete not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sReplacementStyle not a String.
;				   @Error 1 @Extended 6 Return 0 = Paragraph Style called in $sReplacementStyle doesn't exist in Document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving "ParagraphStyles" Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Paragraph Style Name.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = $sParStyle is not a User-Created Paragraph Style and cannot be deleted.
;				   @Error 3 @Extended 2 Return 0 = $sParStyle is in use and $bForceDelete is false.
;				   @Error 3 @Extended 3 Return 0 = $sParStyle still exists after deletion attempt.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Paragraph Style called in $sParStyle was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ParStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleDelete(ByRef $oDoc, ByRef $oParStyle, $bForceDelete = False, $sReplacementStyle = "Default Paragraph Style")
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oParStyles
	Local $sParStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bForceDelete) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sReplacementStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If ($sReplacementStyle <> "") And Not _LOWriter_ParStyleExists($oDoc, $sReplacementStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oParStyles = $oDoc.StyleFamilies().getByName("ParagraphStyles")
	If Not IsObj($oParStyles) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	$sParStyle = $oParStyle.Name()
	If Not IsString($sParStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	If Not $oParStyle.isUserDefined() Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
	If $oParStyle.isInUse() And Not ($bForceDelete) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 2, 0) ; If Style is in use return an error unless force delete is true.

	If ($oParStyle.getParentStyle() = Null) Or ($sReplacementStyle <> "Default Paragraph Style") Then $oParStyle.setParentStyle($sReplacementStyle)
	; If Parent style is blank set it to "Default Paragraph Style", Or if not but User has called a specific style set it to that.

	$oParStyles.removeByName($sParStyle)
	Return ($oParStyles.hasByName($sParStyle)) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 3, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ParStyleDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleDropCaps
; Description ...: Set or Retrieve DropCaps settings for a Paragraph style.
; Syntax ........: _LOWriter_ParStyleDropCaps(ByRef $oDoc, $oParStyle[, $iNumChar = Null[, $iLines = Null[, $iSpcTxt = Null[, $bWholeWord = Null[, $sCharStyle = Null]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iNumChar            - [optional] an integer value (0-9). Default is Null. The number of characters to make into DropCaps.
;                  $iLines              - [optional] an integer value (0,2-9). Default is Null. The number of lines to drop down.
;                  $iSpcTxt             - [optional] an integer value. Default is Null. The distance between the drop cap and the following text. in Micrometers.
;                  $bWholeWord          - [optional] a boolean value. Default is Null. If True, DropCap the whole first word. (Nullifys $iNumChars.)
;                  $sCharStyle          - [optional] a string value. Default is Null. The character style to use for the DropCaps. See Remarks.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 4 Return 0 = Character Style called in $sCharStyle not found in document.
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iNumChar not an integer, less than 0, or greater than 9.
;				   @Error 1 @Extended 7 Return 0 = $iLines not an Integer, less than 0, equal to 1, or greater than 9
;				   @Error 1 @Extended 8 Return 0 = $iSpaceTxt not an Integer, or less than 0.
;				   @Error 1 @Extended 9 Return 0 = $bWholeWord not a Boolean.
;				   @Error 1 @Extended 10 Return 0 = $sCharStyle not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving DropCap Format Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iNumChar
;				   |								2 = Error setting $iLines
;				   |								4 = Error setting $iSpcTxt
;				   |								8 = Error setting $bWholeWord
;				   |								16 = Error setting $sCharStyle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Set $iNumChars, $iLines, $iSpcTxt to 0 to disable DropCaps.
;					I am unable to find a way to set Drop Caps character style to "None" as is available in the User Interface.
;					When it is set to "None" Libre returns a blank string ("") but setting it to a blank string throws a COM
;					error/Exception, even when attempting to set it to Libre's own return value without any in-between
;					variables, in case I was mistaken as to it being a blank string, but this still caused a COM error. So
;					consequently, you cannot set Character Style to "None", but you can still disable Drop Caps as noted above.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleDropCaps(ByRef $oDoc, ByRef $oParStyle, $iNumChar = Null, $iLines = Null, $iSpcTxt = Null, $bWholeWord = Null, $sCharStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sCharStyle <> Null) And Not _LOWriter_CharStyleExists($oDoc, $sCharStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOWriter_ParDropCaps($oParStyle, $iNumChar, $iLines, $iSpcTxt, $bWholeWord, $sCharStyle)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleDropCaps

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleEffect
; Description ...: Set or Retrieve the Font Effect settings for a Paragraph Style.
; Syntax ........: _LOWriter_ParStyleEffect(ByRef $oParStyle[, $iRelief = Null[, $iCase = Null[, $bHidden = Null[, $bOutline = Null[, $bShadow = Null]]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iRelief             - [optional] an integer value (0-2). Default is Null. The Character Relief style. See Constants, $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iCase               - [optional] an integer value (0-4). Default is Null. The Character Case Style. See Constants, $LOW_CASEMAP_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bHidden             - [optional] a boolean value. Default is Null. If True, the Characters are hidden.
;                  $bOutline            - [optional] a boolean value. Default is Null. If True, the characters have an outline around the outside.
;                  $bShadow             - [optional] a boolean value. Default is Null. If True, the characters have a shadow.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iRelief not an integer or less than 0, or greater than 2. See Constants, $LOW_RELIEF_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 5 Return 0 = $iCase not an integer or less than 0, or greater than 4. See Constants, $LOW_CASEMAP_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $bHidden not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $bOutline not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $bShadow not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iRelief
;				   |								2 = Error setting $iCase
;				   |								4 = Error setting $bHidden
;				   |								8 = Error setting $bOutline
;				   |								16 = Error setting $bShadow
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleEffect(ByRef $oParStyle, $iRelief = Null, $iCase = Null, $bHidden = Null, $bOutline = Null, $bShadow = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharEffect($oParStyle, $iRelief, $iCase, $bHidden, $bOutline, $bShadow)
	Return SetError(@error, @extended, $vReturn)

EndFunc   ;==>_LOWriter_ParStyleEffect

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleExists
; Description ...: Check whether a Document contains a specific Paragraph Style by name.
; Syntax ........: _LOWriter_ParStyleExists(ByRef $oDoc, $sParStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sParStyle           - a string value. The Paragraph Style Name to search for.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sParStyle not a String.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. If the Document contains the Paragraph style calling in $sParStyle, True is returned, else False.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleExists(ByRef $oDoc, $sParStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If $oDoc.StyleFamilies.getByName("ParagraphStyles").hasByName($sParStyle) Then Return SetError($__LO_STATUS_SUCCESS, 0, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOWriter_ParStyleExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleFont
; Description ...: Set and Retrieve the Font Settings for a Paragraph Style.
; Syntax ........: _LOWriter_ParStyleFont(ByRef $oDoc, ByRef $oParStyle[, $sFontName = Null[, $nFontSize = Null[, $iPosture = Null[, $iWeight = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $sFontName           - [optional] a string value. Default is Null. The Font Name to use.
;                  $nFontSize           - [optional] a general number value. Default is Null. The new Font size.
;                  $iPosture            - [optional] an integer value (0-5). Default is Null. The Font Italic setting. See Constants, $LOW_POSTURE_* as defined in LibreOfficeWriter_Constants.au3. Also see remarks.
;                  $iWeight             - [optional] an integer value (0,50-200). Default is Null. The Font Bold settings see Constants, $LOW_WEIGHT_* as defined in LibreOfficeWriter_Constants.au3. Also see remarks.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 4 Return 0 = Font called in $sFontName not available in current document.
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $sFontName not a String.
;				   @Error 1 @Extended 7 Return 0 = $nFontSize not a Number.
;				   @Error 1 @Extended 8 Return 0 = $iPosture not an Integer, less than 0, or greater than 5. See Constants.
;				   @Error 1 @Extended 9 Return 0 = $iWeight less than 50 and not 0, or more than 200. See Constants.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sFontName
;				   |								2 = Error setting $nFontSize
;				   |								4 = Error setting $iPosture
;				   |								8 = Error setting $iWeight
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					Not every font accepts Bold and Italic settings, and not all settings for bold and Italic are accepted,
;					such as oblique, ultra Bold etc. Libre Writer accepts only the predefined weight values, any other values
;					are changed automatically to an acceptable value, which could trigger a settings error.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_FontsList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleFont(ByRef $oDoc, ByRef $oParStyle, $sFontName = Null, $nFontSize = Null, $iPosture = Null, $iWeight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($sFontName <> Null) And Not _LOWriter_FontExists($oDoc, $sFontName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOWriter_CharFont($oParStyle, $sFontName, $nFontSize, $iPosture, $iWeight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleFont

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleFontColor
; Description ...: Set or retrieve the font color, transparency and highlighting of a paragraph style.
; Syntax ........: _LOWriter_ParStyleFontColor(ByRef $oParStyle[, $iFontColor = Null[, $iTransparency = Null[, $iHighlight = Null]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iFontColor          - [optional] an integer value (-1-16777215). Default is Null. the desired Color value in Long Integer format, to make the font, can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for Auto color.
;                  $iTransparency       - [optional] an integer value (0-100). Default is Null. Transparency percentage. 0 is visible, 100 is invisible. Available for Libre Office 7.0 and up.
;                  $iHighlight          - [optional] an integer value (-1-16777215). Default is Null. A Color value in Long Integer format, to highlight the text in, can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for No color.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iFontColor not an integer, less than -1, or greater than 16777215.
;				   @Error 1 @Extended 5 Return 0 = $iTransparency not an Integer, or less than 0, or greater than 100%.
;				   @Error 1 @Extended 6 Return 0 = $iHighlight not an integer, less than -1, or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $FontColor
;				   |								2 = Error setting $iTransparency.
;				   |								4 = Error setting $iHighlight
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 7.0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 or 3 Element Array with values in order of function parameters. If The current Libre Office version is below 7.0 the returned array will contain 2 elements, because $iTransparency is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Note: When setting transparency, the value of font color value is changed.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertColorFromLong,
;					_LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleFontColor(ByRef $oParStyle, $iFontColor = Null, $iTransparency = Null, $iHighlight = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharFontColor($oParStyle, $iFontColor, $iTransparency, $iHighlight)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleFontColor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleGetObj
; Description ...: Retrieve a Paragraph Style Object for use with other ParStyle functions.
; Syntax ........: _LOWriter_ParStyleGetObj(ByRef $oDoc, $sParStyle)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sParStyle           - a string value. The Paragraph Style name to retrieve the Object for.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sParStyle not a String.
;				   @Error 1 @Extended 3 Return 0 = Paragraph Style called in $sParStyle not found in Document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Paragraph Style Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Paragraph Style successfully retrieved, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ParStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleGetObj(ByRef $oDoc, $sParStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oParStyle

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not _LOWriter_ParStyleExists($oDoc, $sParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	$oParStyle = $oDoc.StyleFamilies().getByName("ParagraphStyles").getByName($sParStyle)
	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oParStyle)
EndFunc   ;==>_LOWriter_ParStyleGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleHyphenation
; Description ...: Set or Retrieve Hyphenation settings for a paragraph Style.
; Syntax ........: _LOWriter_ParStyleHyphenation(ByRef $oParStyle[, $bAutoHyphen = Null[, $bHyphenNoCaps = Null[, $iMaxHyphens = Null[, $iMinLeadingChar = Null[, $iMinTrailingChar = Null]]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $bAutoHyphen         - [optional] a boolean value. Default is Null. If True, automatic hyphenation is applied.
;                  $bHyphenNoCaps       - [optional] a boolean value. Default is Null. If True, hyphenation of words written in CAPS is disabled for this paragraph. Libre 6.4 and up.
;                  $iMaxHyphens         - [optional] an integer value (0-99). Default is Null. The maximum number of consecutive hyphens.
;                  $iMinLeadingChar     - [optional] an integer value (2-9). Default is Null. The minimum number of characters to remain before the hyphen character (when hyphenation is applied).
;                  $iMinTrailingChar    - [optional] an integer value (2-9). Default is Null. The minimum number of characters to remain after the hyphen character (when hyphenation is applied).
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bAutoHyphen not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bHyphenNoCaps not  a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iMaxHyphens not an Integer, less than 0, or greater than 99.
;				   @Error 1 @Extended 7 Return 0 = $iMinLeadingChar not an Integer, less than 2, or greater than 9.
;				   @Error 1 @Extended 8 Return 0 = $iMinTrailingChar not an Integer, less than 2, or greater than 9.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bAutoHyphen
;				   |								2 = Error setting $bHyphenNoCaps
;				   |								4 = Error setting $iMaxHyphens
;				   |								8 = Error setting $iMinLeadingChar
;				   |								16 = Error setting $iMinTrailingChar
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 6.4.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 or 5 Element Array with values in order of function parameters. If the current Libre Office Version is below 6.4, then the Array returned will contain 4 elements because $bHyphenNoCaps is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note: $bAutoHyphen needs to be set to True for the rest of the settings to be activated, but they will be still successfully be set regardless.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleHyphenation(ByRef $oParStyle, $bAutoHyphen = Null, $bHyphenNoCaps = Null, $iMaxHyphens = Null, $iMinLeadingChar = Null, $iMinTrailingChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_ParHyphenation($oParStyle, $bAutoHyphen, $bHyphenNoCaps, $iMaxHyphens, $iMinLeadingChar, $iMinTrailingChar)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleHyphenation

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleIndent
; Description ...: Set or Retrieve Indent settings for a Paragraph Style.
; Syntax ........: _LOWriter_ParStyleIndent(ByRef $oParStyle[, $iBeforeTxt = Null[, $iAfterTxt = Null[, $iFirstLine = Null[, $bAutoFirstLine = Null]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iBeforeTxt          - [optional] an integer value (-9998989-17094). Default is Null. The amount of space to indent the paragraph from the page margin. If you want the paragraph to extend into the page margin, enter a negative number. Set in Micrometers(uM).
;                  $iAfterTxt           - [optional] an integer value (-9998989-17094). Default is Null. The amount of space to indent the paragraph from the page margin. If you want the paragraph to extend into the page margin, enter a negative number. Set in Micrometers(uM).
;                  $iFirstLine          - [optional] an integer value (-57785-17094). Default is Null. The amount to Indent the first line of a paragraph. Set in Micrometers(uM).
;                  $bAutoFirstLine      - [optional] a boolean value. Default is Null. If True, the first line will be indented automatically.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iBeforeText not an integer, less than -9998989, or greater than 17094 uM.
;				   @Error 1 @Extended 5 Return 0 = $iAfterText not an integer, less than -9998989, or greater than 17094 uM.
;				   @Error 1 @Extended 6 Return 0 = $iFirstLine not an integer, less than -57785, or greater than 17094 uM.
;				   @Error 1 @Extended 7 Return 0 = $bAutoFirstLine not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iBeforeTxt
;				   |								2 = Error setting $iAfterTxt
;				   |								4 = Error setting $iFirstLine
;				   |								8 = Error setting $bAutoFirstLine
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iFirstLine Indent cannot be set if $bAutoFirstLine is set to True.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleIndent(ByRef $oParStyle, $iBeforeTxt = Null, $iAfterTxt = Null, $iFirstLine = Null, $bAutoFirstLine = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_ParIndent($oParStyle, $iBeforeTxt, $iAfterTxt, $iFirstLine, $bAutoFirstLine)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleIndent

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleOrganizer
; Description ...: Set or retrieve the Organizer settings of a Paragraph Style.
; Syntax ........: _LOWriter_ParStyleOrganizer(ByRef $oDoc, $oParStyle[, $sNewParStyleName = Null[, $sFollowStyle = Null[, $sParentStyle = Null[, $bAutoUpdate = Null[, $bHidden = Null]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $sNewParStyleName    - [optional] a string value. Default is Null. The new name to set the paragraph style called in $oParStyle to.
;                  $sFollowStyle        - [optional] a string value. Default is Null. The Paragraph Style name to apply to the following paragraph.
;                  $sParentStyle        - [optional] a string value. Default is Null. Set an existing  paragraph style (or an Empty String ("") = - None -) to apply its settings to the current style. Use the other settings to modify the inherited style settings.
;                  $bAutoUpdate         - [optional] a boolean value. Default is Null. If True, Updates the style when you apply direct formatting to a paragraph using this style in your document. The formatting of all paragraphs using this style is automatically updated.
;                  $bHidden             - [optional] a boolean value. Default is Null. If True, this style is hidden in the L.O. UI. Libre 4.0 and up only.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 4 Return 0 = $sNewParStyleName not a String.
;				   @Error 1 @Extended 5 Return 0 = Paragraph Style name called in $sNewParStyleName already exists in document.
;				   @Error 1 @Extended 6 Return 0 = $sFollowStyle not a String.
;				   @Error 1 @Extended 7 Return 0 = Paragraph Style called in $sFollowStyle doesn't exist in this document.
;				   @Error 1 @Extended 8 Return 0 = $sParentStyle not a String.
;				   @Error 1 @Extended 9 Return 0 = Paragraph Style called in $sParentStyle doesn't exist in this Document.
;				   @Error 1 @Extended 10 Return 0 = $bAutoUpdate not a Boolean.
;				   @Error 1 @Extended 11 Return 0 = $bHidden not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sNewParStyleName
;				   |								2 = Error setting $sFollowStyle
;				   |								4 = Error setting $sParentStyle
;				   |								8 = Error setting $bAutoUpdate
;				   |								16 = Error setting $bHidden
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 4.0.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 or 5 Element Array with values in order of function parameters.
;				   +								If the Libre Office version is below 4.0, the Array will contain 4 elements because $bHidden is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ParStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleOrganizer(ByRef $oDoc, ByRef $oParStyle, $sNewParStyleName = Null, $sFollowStyle = Null, $sParentStyle = Null, $bAutoUpdate = Null, $bHidden = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avOrganizer[4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If __LOWriter_VarsAreNull($sNewParStyleName, $sParentStyle, $sFollowStyle, $bAutoUpdate, $bHidden) Then
		If __LOWriter_VersionCheck(4.0) Then
			__LOWriter_ArrayFill($avOrganizer, $oParStyle.Name(), __LOWriter_ParStyleNameToggle($oParStyle.getPropertyValue("FollowStyle"), True), _
					__LOWriter_ParStyleNameToggle($oParStyle.ParentStyle(), True), _
					$oParStyle.IsAutoUpdate(), $oParStyle.Hidden())
		Else
			__LOWriter_ArrayFill($avOrganizer, $oParStyle.Name(), __LOWriter_ParStyleNameToggle($oParStyle.getPropertyValue("FollowStyle"), True), _
					__LOWriter_ParStyleNameToggle($oParStyle.ParentStyle(), True), $oParStyle.IsAutoUpdate())
		EndIf
		Return SetError($__LO_STATUS_SUCCESS, 1, $avOrganizer)
	EndIf

	If ($sNewParStyleName <> Null) Then
		If Not IsString($sNewParStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If _LOWriter_ParStyleExists($oDoc, $sNewParStyleName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oParStyle.Name = $sNewParStyleName
		$iError = ($oParStyle.Name() = $sNewParStyleName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sFollowStyle <> Null) Then
		If Not IsString($sFollowStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_ParStyleExists($oDoc, $sFollowStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$sFollowStyle = __LOWriter_ParStyleNameToggle($sFollowStyle)
		$oParStyle.setPropertyValue("FollowStyle", $sFollowStyle)
		$iError = ($oParStyle.getPropertyValue("FollowStyle") = $sFollowStyle) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sParentStyle <> Null) Then
		If Not IsString($sParentStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		If ($sParentStyle <> "") Then
			If Not _LOWriter_ParStyleExists($oDoc, $sParentStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
			$sParentStyle = __LOWriter_ParStyleNameToggle($sParentStyle)
		EndIf
		$oParStyle.ParentStyle = $sParentStyle
		$iError = ($oParStyle.ParentStyle() = $sParentStyle) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($bAutoUpdate <> Null) Then
		If Not IsBool($bAutoUpdate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
		$oParStyle.IsAutoUpdate = $bAutoUpdate
		$iError = ($oParStyle.IsAutoUpdate() = $bAutoUpdate) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($bHidden <> Null) Then
		If Not IsBool($bHidden) Then Return SetError($__LO_STATUS_INPUT_ERROR, 11, 0)
		If Not __LOWriter_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
		$oParStyle.Hidden = $bHidden
		$iError = ($oParStyle.Hidden() = $bHidden) ? ($iError) : (BitOR($iError, 16))
	EndIf

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_ParStyleOrganizer

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleOutLineAndList
; Description ...: Set and Retrieve the Outline and List settings for a paragraph style.
; Syntax ........: _LOWriter_ParStyleOutLineAndList(ByRef $oDoc, $oParStyle[, $iOutline = Null[, $sNumStyle = Null[, $bParLineCount = Null[, $iLineCountVal = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iOutline            - [optional] an integer value (0-10). Default is Null. The Outline Level, see Constants, $LOW_OUTLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sNumStyle           - [optional] a string value. Default is Null. Specifies the Numbering Style name for Paragraph numbering. Set to "" for None.
;                  $bParLineCount       - [optional] a boolean value. Default is Null. If True, the paragraph is included in the line numbering.
;                  $iLineCountVal       - [optional] an integer value. Default is Null. The start value for numbering if a new numbering starts at this paragraph. Set to 0 for no line numbering restart.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 4 Return 0 = Numbering Style called in $sNumStyle not found in current document.
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iOutline not an integer, less than 0, or greater than 10. See constants, $LOW_OUTLINE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 7 Return 0 = $sNumStyle not a String.
;				   @Error 1 @Extended 8 Return 0 = $bParLineCount not a Boolean.
;				   @Error 1 @Extended 9 Return 0 = $iLineCountVal not an Integer, or less than 0.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iOutline
;				   |								2 = Error setting $sNumStyle
;				   |								4 = Error setting $bParLineCount
;				   |								8 = Error setting $iLineCountVal
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_NumStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleOutLineAndList(ByRef $oDoc, ByRef $oParStyle, $iOutline = Null, $sNumStyle = Null, $bParLineCount = Null, $iLineCountVal = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sNumStyle <> Null) And ($sNumStyle <> "") And Not _LOWriter_NumStyleExists($oDoc, $sNumStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOWriter_ParOutLineAndList($oParStyle, $iOutline, $sNumStyle, $bParLineCount, $iLineCountVal)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleOutLineAndList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleOverLine
; Description ...: Set and retrieve the OverLine settings for a paragraph style.
; Syntax ........: _LOWriter_ParStyleOverLine(ByRef $oParStyle[, $bWordOnly = Null[, $iOverLineStyle = Null[, $bOLHasColor = Null[, $iOLColor = Null]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not Overlined.
;                  $iOverLineStyle      - [optional] an integer value (0-18). Default is Null. The style of the Overline line, see constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3. See Remarks.
;                  $bOLHasColor         - [optional] a boolean value. Default is Null. If True, the Overline is colored, must be set to true in order to set the Overline color.
;                  $iOLColor            - [optional] an integer value (-1-16777215). Default is Null. The Overline color, set in Long integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iOverLineStyle not an Integer, or less than 0, or greater than 18. See constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $bOLHasColor not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iOLColor not an Integer, or less than -1, or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bWordOnly
;				   |								2 = Error setting $iOverLineStyle
;				   |								4 = Error setting $OLHasColor
;				   |								8 = Error setting $iOLColor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: OverLine line style uses the same constants as underline style.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Note: $bOLHasColor must be set to true in order to set the Overline color.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertColorFromLong,
;					_LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleOverLine(ByRef $oParStyle, $bWordOnly = Null, $iOverLineStyle = Null, $bOLHasColor = Null, $iOLColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharOverLine($oParStyle, $bWordOnly, $iOverLineStyle, $bOLHasColor, $iOLColor)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleOverLine

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStylePageBreak
; Description ...: Set or Retrieve Page Break Settings for a Paragraph Style.
; Syntax ........: _LOWriter_ParStylePageBreak(ByRef $oDoc, $oParStyle[, $iBreakType = Null[, $sPageStyle = Null[, $iPgNumOffSet = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iBreakType          - [optional] an integer value (0-6). Default is Null. The Page Break Type. See Constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sPageStyle          - [optional] a string value. Default is Null. Creates a page break before the paragraph it belongs to and assigns the value as the name of the new page style to use. See remarks.
;                  $iPgNumOffSet        - [optional] an integer value. Default is Null. If a page break property is set at a paragraph, this property contains the new value for the page number.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 4 Return 0 = Page Style called in $sPageStyle not found in current document.
;				   @Error 1 @Extended 5 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iBreakType not an integer, less than 0, or greater than 6. See constants, $LOW_BREAK_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 7 Return 0 = $sPageStyle not a String.
;				   @Error 1 @Extended 8 Return 0 = $iPgNumOffSet not an Integer, or less than 0.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iBreakType
;				   |								2 = Error setting $sPageStyle
;				   |								4 = Error setting $iPgNumOffSet
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Break Type must be set before Page Style will be able to be set, and page style needs set before $iPgNumOffSet can be set.
;					Libre doesn't directly show in its User interface options for Break type constants #3 and #6 (Column both) and (Page both),
;						but doesn't throw an error when being set to either one, so they are included here, though I'm not sure if they will work correctly.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Note: If you set the $sPageStyle parameter, to remove the page break setting you must set this to "".
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStylePageBreak(ByRef $oDoc, ByRef $oParStyle, $iBreakType = Null, $sPageStyle = Null, $iPgNumOffSet = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If ($sPageStyle <> Null) And ($sPageStyle <> "") And Not _LOWriter_PageStyleExists($oDoc, $sPageStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOWriter_ParPageBreak($oParStyle, $iBreakType, $sPageStyle, $iPgNumOffSet)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStylePageBreak

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStylePosition
; Description ...: Set and retrieve settings related to Sub/Super Script and relative size.
; Syntax ........: _LOWriter_ParStylePosition(ByRef $oParStyle[, $bAutoSuper = Null[, $iSuperScript = Null[, $bAutoSub = Null[, $iSubScript = Null[, $iRelativeSize = Null]]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $bAutoSuper          - [optional] a boolean value. Default is Null. If True, automatic sizing for Superscript is active.
;                  $iSuperScript        - [optional] an integer value (0-100,14000) Default is Null. The Superscript percentage value. See Remarks.
;                  $bAutoSub            - [optional] a boolean value. Default is Null. If True, automatic sizing for Subscript is active.
;                  $iSubScript          - [optional] an integer value (-100-100,-14000,14000) Default is Null. The Subscript percentage value. See Remarks.
;                  $iRelativeSize       - [optional] an integer value (1-100). Default is Null. The size percentage relative to current font size.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bAutoSuper not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bAutoSub not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iSuperScript not an integer, or less than 0, greater than 100 and not 14000.
;				   @Error 1 @Extended 7 Return 0 = $iSubScript not an integer, or less than -100, greater than 100, and not 14000 or -14000.
;				   @Error 1 @Extended 8 Return 0 = $iRelativeSize not an integer, or less than 1, greater than 100.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iSuperScript
;				   |								2 = Error setting $iSubScript
;				   |								4 = Error setting $iRelativeSize.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;					Set either $iSubScript or $iSuperScript to 0 to return it to Normal setting.
;					The way LibreOffice is set up Super/Subscript are set in the same setting, Super is a positive number from 1 to 100 (percentage),Subscript is a negative number set to 1 to 100 percentage.
;					For the user's convenience this function accepts both positive and negative numbers for Subscript, if a positive number is called for Subscript, it is automatically set to a negative.
;					Automatic Superscript has a integer value of 14000, Auto Subscript has a integer value of -14000.
;					There is no settable setting of Automatic Super/Sub Script, though one exists, it is read-only in LibreOffice, consequently I have made two
;						separate parameters to be able to determine if the user wants to automatically set Superscript or Subscript.
;					If you set both Auto Superscript to True and Auto Subscript to True, or $iSuperScript to an integer and $iSubScript to an integer,
;						Subscript will be set as it is the last in the line to be set in this function, and thus will over-write any Superscript settings.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStylePosition(ByRef $oParStyle, $bAutoSuper = Null, $iSuperScript = Null, $bAutoSub = Null, $iSubScript = Null, $iRelativeSize = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharPosition($oParStyle, $bAutoSuper, $iSuperScript, $bAutoSub, $iSubScript, $iRelativeSize)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStylePosition

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleRotateScale
; Description ...: Set or retrieve the character rotational and Scale settings for a paragraph Style.
; Syntax ........: _LOWriter_ParStyleRotateScale(ByRef $oParStyle[, $iRotation = Null[, $iScaleWidth = Null]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iRotation           - [optional] an integer value (0,90,270). Default is Null. Degrees to rotate the text.
;                  $iScaleWidth         - [optional] an integer value (1-100). Default is Null. The percentage to  horizontally stretch or compress the text. 100 is normal sizing.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iRotation not an Integer or not equal to 0, 90, or 270 degrees.
;				   @Error 1 @Extended 5 Return 0 = $iScaleWidth not an Integer or less than 1%, or greater than 100%.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iRotation
;				   |								2 = Error setting $iScaleWidth
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleRotateScale(ByRef $oParStyle, $iRotation = Null, $iScaleWidth = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharRotateScale($oParStyle, $iRotation, $iScaleWidth)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleRotateScale

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleSet
; Description ...: Set a Paragraph style for a paragraph by Cursor or paragraph Object.
; Syntax ........: _LOWriter_ParStyleSet(ByRef $oDoc, ByRef $oObj, $sParStyle)
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oObj           - [in/out] an object. A Cursor Object returned from any Cursor Object creation or retrieval functions, Or A Paragraph Object returned from _LOWriter_ParObjCreateList function.
;                  $sParStyle      - a string value. The Paragraph Style name to set the paragraph to.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oObj not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oObj does not support Paragraph Properties Service.
;				   @Error 1 @Extended 4 Return 0 = $sParStyle not a String.
;				   @Error 1 @Extended 5 Return 0 = Paragraph Style called in $sParStyle not found in Document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended 1 Return 0 = Error setting Paragraph Style.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Paragraph Style successfully set.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor, _LOWriter_ParStylesGetNames
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleSet(ByRef $oDoc, ByRef $oObj, $sParStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oObj) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oObj.supportsService("com.sun.star.style.ParagraphProperties") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not _LOWriter_ParStyleExists($oDoc, $sParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	$sParStyle = __LOWriter_ParStyleNameToggle($sParStyle)
	$oObj.ParaStyleName = $sParStyle
	Return ($oObj.ParaStyleName() = $sParStyle) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_PROP_SETTING_ERROR, 1, 0))
EndFunc   ;==>_LOWriter_ParStyleSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStylesGetNames
; Description ...: Retrieve a list of all Paragraph Style names available for a document.
; Syntax ........: _LOWriter_ParStylesGetNames(ByRef $oDoc[, $bUserOnly = False[, $bAppliedOnly = False]])
; Parameters ....: $oDoc           - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $bUserOnly      - [optional] a boolean value. Default is False. If True only User-Created Paragraph Styles are returned.
;                  $bAppliedOnly   - [optional] a boolean value. Default is False. If True only Applied paragraph Styles are returned.
; Return values .: Success: Integer or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bUserOnly not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $bAppliedOnly not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Paragraph Styles Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 0 = Success. No Paragraph Styles found according to parameters.
;				   @Error 0 @Extended ? Return Array = Success. An Array containing all Paragraph Styles matching the input parameters.
;				   +		@Extended contains the count of results returned.
;				   +		If Only a Document object is input, all available Paragraph styles will be returned.
;				   +		Else if $bUserOnly is set to True, only User-Created Paragraph Styles are returned.
;				   +		Else if $bAppliedOnly is set to True, only Applied paragraph Styles are returned.
;				   +		If Both are true then only User-Created paragraph styles that are applied are returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Two paragraph styles have two separate names, Default Paragraph Style is also internally called "Standard" and Complimentary Close, which is internally called "Salutation".
;					Either name works when setting a Paragraph Style, but on certain functions that return a Paragraph Style Name, you may see one of these alternative names.
; Related .......: _LOWriter_ParStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStylesGetNames(ByRef $oDoc, $bUserOnly = False, $bAppliedOnly = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oStyles
	Local $aStyles[0]
	Local $iCount = 0
	Local $sExecute = ""

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsBool($bUserOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bAppliedOnly) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	$oStyles = $oDoc.StyleFamilies.getByName("ParagraphStyles")
	If Not IsObj($oStyles) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	ReDim $aStyles[$oStyles.getCount()]

	If Not $bUserOnly And Not $bAppliedOnly Then
		For $i = 0 To $oStyles.getCount() - 1
			$aStyles[$i] = $oStyles.getByIndex($i).DisplayName()
			Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
		Next
		Return SetError($__LO_STATUS_SUCCESS, $i, $aStyles)
	EndIf

	$sExecute = ($bUserOnly) ? ("($oStyles.getByIndex($i).isUserDefined())") : ($sExecute)
	$sExecute = ($bUserOnly And $bAppliedOnly) ? ($sExecute & " And ") : ($sExecute)
	$sExecute = ($bAppliedOnly) ? ($sExecute & "($oStyles.getByIndex($i).isInUse())") : ($sExecute)

	For $i = 0 To $oStyles.getCount() - 1
		If Execute($sExecute) Then
			$aStyles[$iCount] = $oStyles.getByIndex($i).DisplayName
			$iCount += 1
		EndIf
		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next
	ReDim $aStyles[$iCount]

	Return ($iCount = 0) ? (SetError($__LO_STATUS_SUCCESS, 0, 1)) : (SetError($__LO_STATUS_SUCCESS, $iCount, $aStyles))
EndFunc   ;==>_LOWriter_ParStylesGetNames

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleShadow
; Description ...: Set or Retrieve the Shadow settings for a Paragraph style.
; Syntax ........: _LOWriter_ParStyleShadow(ByRef $oParStyle[, $iWidth = Null[, $iColor = Null[, $bTransparent = Null[, $iLocation = Null]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iWidth              - [optional] an integer value. Default is Null. The shadow width, set in Micrometers.
;                  $iColor              - [optional] an integer value (0-16777215). Default is Null. The color of the shadow, set in Long Integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bTransparent        - [optional] a boolean value. Default is Null. If True, the shadow is transparent.
;                  $iLocation           - [optional] an integer value (0-4). Default is Null. The location of the shadow compared to the paragraph. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iWidth not an integer or less than 0.
;				   @Error 1 @Extended 5 Return 0 = $iColor not an integer, less than 0, or greater than 16777215.
;				   @Error 1 @Extended 6 Return 0 = $bTransparent not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iLocation not an Integer, less than 0, or greater than 4. See Constants, $LOW_SHADOW_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Shadow Format Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Shadow Format Object for Error checking.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iWidth
;				   |								2 = Error setting $iColor
;				   |								4 = Error setting $bTransparent
;				   |								8 = Error setting $iLocation
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Note: LibreOffice may change the shadow width +/- a Micrometer.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertColorFromLong,
;					_LOWriter_ConvertColorToLong,  _LOWriter_ConvertFromMicrometer, _LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleShadow(ByRef $oParStyle, $iWidth = Null, $iColor = Null, $bTransparent = Null, $iLocation = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_ParShadow($oParStyle, $iWidth, $iColor, $bTransparent, $iLocation)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleShadow

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleSpace
; Description ...: Set and Retrieve Line Spacing settings for a paragraph style.
; Syntax ........: _LOWriter_ParStyleSpace(ByRef $oParStyle[, $iAbovePar = Null[, $iBelowPar = Null[, $bAddSpace = Null[, $iLineSpcMode = Null[, $iLineSpcHeight = Null[, $bPageLineSpc = Null]]]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iAbovePar           - [optional] an integer value (0-10008). Default is Null. The Space above a paragraph, in Micrometers.
;                  $iBelowPar           - [optional] an integer value (0-10008). Default is Null. The Space Below a paragraph, in Micrometers.
;                  $bAddSpace           - [optional] a boolean value. Default is Null. If true, the top and bottom margins of the paragraph should not be applied when the previous and next paragraphs have the same style. Libre Office Version 3.6 and Up.
;                  $iLineSpcMode        - [optional] an integer value (0-3). Default is Null. The line spacing type for the paragraph. See Constants, $LOW_LINE_SPC_MODE_* as defined in LibreOfficeWriter_Constants.au3, also notice min and max values for each.
;                  $iLineSpcHeight      - [optional] an integer value. Default is Null. This value specifies the spacing of the lines. See Remarks for Minimum and Max values.
;                  $bPageLineSpc        - [optional] a boolean value. Default is Null. If True, register mode is applied to the paragraph. See Remarks.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $iAbovePar not an integer, less than 0, or greater than 10008 uM.
;				   @Error 1 @Extended 5 Return 0 = $iBelowPar not an integer, less than 0, or greater than 10008 uM.
;				   @Error 1 @Extended 6 Return 0 = $bAddSpc not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iLineSpcMode not an integer, less than 0, or greater than 3. See Constants, $LOW_LINE_SPC_MODE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 8 Return 0 = $iLineSpcHeight not an Integer.
;				   @Error 1 @Extended 9 Return 0 = $iLineSpcMode set to 0(Proportional) and $iLineSpcHeight less than 6(%) or greater than 65535(%).
;				   @Error 1 @Extended 10 Return 0 = $iLineSpcMode set to 1, or 2 (Minimum, or Leading) and $iLineSpcHeight less than 0, uM or greater than 10008 uM
;				   @Error 1 @Extended 11 Return 0 = $iLineSpcMode set to 3 (Fixed) and $iLineSpcHeight less than 51, uM or greater than 10008 uM.
;				   @Error 1 @Extended 12 Return 0 = $bPageLineSpc not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaLineSpacing Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iAbovePar
;				   |								2 = Error setting $iBelowPar
;				   |								4 = Error setting $bAddSpace
;				   |								8 = Error setting $iLineSpcMode
;				   |								16 = Error setting $iLineSpcHeight
;				   |								32 = Error setting bPageLineSpc
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office version lower than 3.6.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 or 6 Element Array with values in order of function parameters.
;				   +								If the current Libre Office version is less than 3.6, the returned Array will contain 5 elements, because $bAddSpace is not available.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $bPageLineSpc(Register mode) is only used if the register mode property of the page style is switched on.
;					$bPageLineSpc(Register Mode) Aligns the baseline of each line of text to a vertical document grid, so that each line is the same height.
;					Note: The settings in Libre Office, (Single,1.15, 1.5, Double,) Use the Proportional mode, and are just varying percentages. e.g Single = 100, 1.15 = 115%, 1.5 = 150%, Double = 200%.
;					$iLineSpcHeight depends on the $iLineSpcMode used, see constants for accepted Input values.
;					Note: $iAbovePar, $iBelowPar, $iLineSpcHeight may change +/- 1 Micrometer once set.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleSpace(ByRef $oParStyle, $iAbovePar = Null, $iBelowPar = Null, $bAddSpace = Null, $iLineSpcMode = Null, $iLineSpcHeight = Null, $bPageLineSpc = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_ParSpace($oParStyle, $iAbovePar, $iBelowPar, $bAddSpace, $iLineSpcMode, $iLineSpcHeight, $bPageLineSpc)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleSpace

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleSpacing
; Description ...: Set and retrieve the spacing between characters (Kerning) for a Paragraph style.
; Syntax ........: _LOWriter_ParStyleSpacing(ByRef $oParStyle[, $bAutoKerning = Null[, $nKerning = Null]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $bAutoKerning        - [optional] a boolean value. Default is Null. If True, applies a spacing in between certain pairs of characters.
;                  $nKerning            - [optional] a general number value (-2-928.8). Default is Null. The kerning value of the characters. See Remarks. Values are in Printer's Points as set in the Libre Office UI.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bAutoKerning not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $nKerning not a number, or less than -2, or greater than 928.8 Points.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bAutoKerning
;				   |								2 = Error setting $nKerning.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					When setting Kerning values in LibreOffice, the measurement is listed in Pt (Printer's Points) in the User Display, however the internal setting is measured in Micrometers.
;					They will be automatically converted from Points to Micrometers and back for retrieval of settings.
;					The acceptable values are from -2 Pt to  928.8 Pt. the figures can be directly converted easily,
;						however, for an unknown reason to myself, LibreOffice begins counting backwards and in negative Micrometers internally from 928.9 up to 1000 Pt (Max setting).
;					For example, 928.8Pt is the last correct value, which equals 32766 uM (Micrometers), after this LibreOffice reports the following: ;928.9 Pt = -32766 uM; 929 Pt = -32763 uM; 929.1 = -32759; 1000 pt = -30258.
;					Attempting to set Libre's kerning value to anything over 32768 uM causes a COM exception, and attempting to set the kerning to any of these negative numbers sets the User viewable kerning value to -2.0 Pt.
;					For these reasons the max settable kerning is -2.0 Pt  to 928.8 Pt.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleSpacing(ByRef $oParStyle, $bAutoKerning = Null, $nKerning = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharSpacing($oParStyle, $bAutoKerning, $nKerning)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleSpacing

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleStrikeOut
; Description ...: Set or Retrieve the StrikeOut settings for a Paragraph style.
; Syntax ........: _LOWriter_ParStyleStrikeOut(ByRef $oParStyle[, $bWordOnly = Null[, $bStrikeOut = Null[, $iStrikeLineStyle = Null]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If True, strike out is applied to words only, skipping whitespaces.
;                  $bStrikeOut          - [optional] a boolean value. Default is Null. If True, strikeout is applied to characters.
;                  $iStrikeLineStyle    - [optional] an integer value (0-6). Default is Null. The Strikeout Line Style, see constants, $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bStrikeOut not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iStrikeLineStyle not an Integer, or less than 0, or greater than 6. See constants, $LOW_STRIKEOUT_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bWordOnly
;				   |								2 = Error setting $bStrikeOut
;				   |								4 = Error setting $iStrikeLineStyle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Note Strikeout line style is converted to a single line in Ms word document format.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleStrikeOut(ByRef $oParStyle, $bWordOnly = Null, $bStrikeOut = Null, $iStrikeLineStyle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharStrikeOut($oParStyle, $bWordOnly, $bStrikeOut, $iStrikeLineStyle)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleStrikeOut

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleTabStopCreate
; Description ...: Create a new TabStop for a Paragraph Style.
; Syntax ........: _LOWriter_ParStyleTabStopCreate(ByRef $oParStyle, $iPosition[, $iAlignment = Null[, $iFillChar = Null[, $iDecChar = Null]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iPosition           - an integer value. The TabStop position to set the new TabStop to. Set in Micrometers (uM). See Remarks.
;                  $iFillChar           - [optional] an integer value. Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
;                  $iAlignment          - [optional] an integer value (0-4). Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iDecChar            - [optional] an integer value. Default is Null. A character (in Asc Value(See AutoIt Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOW_TAB_ALIGN_DECIMAL.
; Return values .: Success: Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = $iPosition not an Integer.
;				   @Error 1 @Extended 4 Return 0 = Position called in $iPosition already exists in this Paragraph Style.
;				   @Error 1 @Extended 5 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iFillChar not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iAlignment not an Integer, less than 0, or greater than 4. See Constants , $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 8 Return 0 = $iDecChar not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Array Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.style.TabStop" Object.
;				   @Error 2 @Extended 3 Return 0 = Error retrieving list of TabStop Positions.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to identify the new Tabstop position once inserted.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return Integer = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iPosition
;				   |								2 = Error setting $iFillChar
;				   |								4 = Error setting $iAlignment
;				   |								8 = Error setting $iDecChar
;				   |								Note: $iNewTabStop position is still returned as even though some settings weren't successfully set, the new TabStop was still created.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Settings were successfully set. New TabStop position is returned.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iPosition once set can vary +/- 1 uM. To ensure you can identify the tabstop to modify it again,
;					This function returns the new TabStop position in @Extended when $iPosition is set, return value will be set to 2. See Return Values.
;					Note: Since $iPosition can fluctuate +/- 1 uM when it is inserted into LibreOffice, it is possible to accidentally overwrite an already existing TabStop.
;					Note: $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32.
;					The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95).
;					You can also enter a custom ASC value. See ASC AutoIt Func. and "ASCII Character Codes" in the AutoIt help file.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertFromMicrometer,
;					_LOWriter_ConvertToMicrometer, _LOWriter_ParStyleTabStopDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleTabStopCreate(ByRef $oParStyle, $iPosition, $iFillChar = Null, $iAlignment = Null, $iDecChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iPosition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If __LOWriter_ParHasTabStop($oParStyle, $iPosition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$iPosition = __LOWriter_ParTabStopCreate($oParStyle, $iPosition, $iAlignment, $iFillChar, $iDecChar)
	Return SetError(@error, @extended, $iPosition)
EndFunc   ;==>_LOWriter_ParStyleTabStopCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleTabStopDelete
; Description ...: Delete a TabStop from a Paragraph Style.
; Syntax ........: _LOWriter_ParStyleTabStopDelete(ByRef $oDoc, $oParStyle, $iTabStop)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iTabStop            - an integer value. The Tab position of the TabStop to modify. See Remarks.
; Return values .: Success: Boolean.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 4 Return 0 = $iTabStop not an Integer.
;				   @Error 1 @Extended 5 Return 0 = Tab Stop position called in $iTabStop not found in this ParStyle.
;				   @Error 1 @Extended 6 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 7 Return 0 = Passed Document Object to internal function not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to identify and delete TabStop in Paragraph.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Returns true if TabStop was successfully deleted.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin.
;					This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be one of a certain length per document.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ParStyleTabStopList,
;					_LOWriter_ParStyleTabStopCreate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleTabStopDelete(ByRef $oDoc, ByRef $oParStyle, $iTabStop)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsInt($iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not __LOWriter_ParHasTabStop($oParStyle, $iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$vReturn = __LOWriter_ParTabStopDelete($oParStyle, $oDoc, $iTabStop)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleTabStopDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleTabStopList
; Description ...: Retrieve a List of TabStops available in a Paragraph Style.
; Syntax ........: _LOWriter_ParStyleTabStopList(ByRef $oParStyle)
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. An Array of TabStops. @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ParStyleTabStopMod,
;					_LOWriter_ParStyleTabStopDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleTabStopList(ByRef $oParStyle)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aiTabList

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$aiTabList = __LOWriter_ParTabStopList($oParStyle)

	Return SetError(@error, @extended, $aiTabList)
EndFunc   ;==>_LOWriter_ParStyleTabStopList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleTabStopMod
; Description ...: Modify or retrieve the properties of an existing TabStop in a Paragraph Style.
; Syntax ........: _LOWriter_ParStyleTabStopMod(ByRef $oParStyle, $iTabStop[, $iPosition = Null[, $iFillChar = Null[, $iAlignment = Null[, $iDecChar = Null]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $iTabStop            - an integer value. The Tab position of the TabStop to modify. See Remarks.
;                  $iPosition           - [optional] an integer value. Default is Null. The New position to set the called Tab Stop position to. Set in Micrometers (uM). See Remarks.
;                  $iFillChar           - [optional] an integer value. Default is Null. The Asc (see AutoIt function) value of any character (except 0/Null) you want to act as a Tab Fill character. See remarks.
;                  $iAlignment          - [optional] an integer value (0-4). Default is Null. The position of where the end of a Tab is aligned to compared to the text. See Constants, $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iDecChar            - [optional] an integer value. Default is Null. A character (in Asc Value(See AutoIt Function)) that you want the decimal tab to use as a decimal separator. Can only be set if $iAlignment is set to $LOW_TAB_ALIGN_DECIMAL.
; Return values .: Success: Integer or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = $iTabStop not an Integer.
;				   @Error 1 @Extended 4 Return 0 = Tab Stop position called in $iTabStop not found in this ParStyle.
;				   @Error 1 @Extended 5 Return 0 = Passed Object to internal function not an Object.
;				   @Error 1 @Extended 6 Return 0 = $iPosition not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iFillChar not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iAlignment not an Integer, less than 0, or greater than 4. See Constants, $LOW_TAB_ALIGN_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 9 Return 0 = $iDecChar not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving ParaTabStops Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Requested TabStop Object.
;				   @Error 2 @Extended 3 Return 0 = Error retrieving list of TabStop Positions.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Paragraph style already contains a TabStop at the length/Position specified in $iPosition.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iPosition
;				   |								2 = Error setting $iFillChar
;				   |								4 = Error setting $iAlignment
;				   |								8 = Error setting $iDecChar
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
;				   @Error 0 @Extended ? Return 2 = Success. Settings were successfully set. New TabStop position is returned in @Extended.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: $iTabStop refers to the position, or essential the "length" of a TabStop from the edge of a page margin.
;						This is the only reliable way to identify a Tabstop to be able to interact with it, as there can only be
;						one of a certain length per document.
;					Note: $iPosition once set can vary +/- 1 uM. To ensure you can identify the tabstop to modify it again,
;						This function returns the new TabStop position in @Extended when $iPosition is set, return value will
;						be set to 2. See Return Values.
;					Note: Since $iPosition can fluctuate +/- 1 uM when it is inserted into LibreOffice, it is possible to
;						accidentally overwrite an already existing TabStop.
;					Note: $iFillChar, Libre's Default value, "None" is in reality a space character which is Asc value 32.
;						The other values offered by Libre are: Period (ASC 46), Dash (ASC 45) and Underscore (ASC 95). You can
;						also enter a custom ASC value. See ASC AutoIt Func. and "ASCII Character Codes" in the AutoIt help file.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ParStyleTabStopCreate,
;					_LOWriter_ParStyleTabStopList, _LOWriter_ConvertFromMicrometer,	_LOWriter_ConvertToMicrometer
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleTabStopMod(ByRef $oParStyle, $iTabStop, $iPosition = Null, $iFillChar = Null, $iAlignment = Null, $iDecChar = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsInt($iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LOWriter_ParHasTabStop($oParStyle, $iTabStop) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$vReturn = __LOWriter_ParTabStopMod($oParStyle, $iTabStop, $iPosition, $iFillChar, $iAlignment, $iDecChar)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleTabStopMod

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleTxtFlowOpt
; Description ...: Set and Retrieve Text Flow settings for a Paragraph Style.
; Syntax ........: _LOWriter_ParStyleTxtFlowOpt(ByRef $oParStyle[, $bParSplit = Null[, $bKeepTogether = Null[, $iParOrphans = Null[, $iParWidows = Null]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $bParSplit           - [optional] a boolean value. Default is Null. If False, prevents the paragraph from splitting between two pages or columns
;                  $bKeepTogether       - [optional] a boolean value. Default is Null. If True, prevents page or column breaks between this and the following paragraph
;                  $iParOrphans         - [optional] an integer value (0,2-9). Default is Null. The minimum number of lines of the paragraph that have to be at bottom of a page if the paragraph is spread over more than one page. 0 = (disabled).
;                  $iParWidows          - [optional] an integer value (0,2-9). Default is Null. The minimum number of lines of the paragraph that have to be at top of a page if the paragraph is spread over more than one page. 0 = (disabled).
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bParSplit not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bKeepTogether not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iParOrphans not an Integer, less than 0, equal to 1, or greater than 9.
;				   @Error 1 @Extended 7 Return 0 = $iParWidows not an Integer, less than 0, equal to 1, or greater than 9.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bParSplit
;				   |								2 = Error setting $bKeepTogether
;				   |								4 = Error setting $iParOrphans
;				   |								8 = Error setting $iParWidows
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: If you do not set ParSplit to True, the rest of the settings will still show to have been set but will not become active until $bParSplit is set to true.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleTxtFlowOpt(ByRef $oParStyle, $bParSplit = Null, $bKeepTogether = Null, $iParOrphans = Null, $iParWidows = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_ParTxtFlowOpt($oParStyle, $bParSplit, $bKeepTogether, $iParOrphans, $iParWidows)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleTxtFlowOpt

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_ParStyleUnderLine
; Description ...: Set and retrieve the Underline settings for a paragraph style.
; Syntax ........: _LOWriter_ParStyleUnderLine(ByRef $oParStyle[, $bWordOnly = Null[, $iUnderLineStyle = Null[, $bULHasColor = Null[, $iULColor = Null]]]])
; Parameters ....: $oParStyle           - [in/out] an object. A Paragraph Style object returned by previous _LOWriter_ParStyleCreate, or _LOWriter_ParStyleGetObj function.
;                  $bWordOnly           - [optional] a boolean value. Default is Null. If true, white spaces are not underlined.
;                  $iUnderLineStyle     - [optional] an integer value (0-18). Default is Null. The Underline line style, see constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bULHasColor         - [optional] a boolean value. Default is Null. If True, the underline is colored, must be set to true in order to set the underline color.
;                  $iULColor            - [optional] an integer value (-1-16777215). Default is Null. The color of the underline, set in Long integer format. Can be a custom value, or one of the constants, $LOW_COLOR_* as defined in LibreOfficeWriter_Constants.au3. Set to $LOW_COLOR_OFF(-1) for automatic color mode.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oParStyle not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oParStyle not a Paragraph Object.
;				   @Error 1 @Extended 3 Return 0 = Passed Object for internal function not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bWordOnly not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iUnderLineStyle not an Integer, or less than 0, or greater than 18. See constants, $LOW_UNDERLINE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $bULHasColor not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iULColor not an Integer, or less than -1, or greater than 16777215.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bWordOnly
;				   |								2 = Error setting $iUnderLineStyle
;				   |								4 = Error setting $ULHasColor
;				   |								8 = Error setting $iULColor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Note: $bULHasColor must be set to true in order to set the underline color.
; Related .......: _LOWriter_ParStyleCreate, _LOWriter_ParStyleGetObj, _LOWriter_ConvertColorFromLong, _LOWriter_ConvertColorToLong
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_ParStyleUnderLine(ByRef $oParStyle, $bWordOnly = Null, $iUnderLineStyle = Null, $bULHasColor = Null, $iULColor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $vReturn

	If Not IsObj($oParStyle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not $oParStyle.supportsService("com.sun.star.style.ParagraphStyle") Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$vReturn = __LOWriter_CharUnderLine($oParStyle, $bWordOnly, $iUnderLineStyle, $bULHasColor, $iULColor)
	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_ParStyleUnderLine

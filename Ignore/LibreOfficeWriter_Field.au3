#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7

#include-once

; Main LibreOffice Includes
#include "LibreOffice_Constants.au3"

; Common includes for Writer
#include "LibreOfficeWriter_Constants.au3"
#include "LibreOfficeWriter_Helper.au3"
#include "LibreOfficeWriter_Internal.au3"

; Other includes for Writer
#include "LibreOfficeWriter_Doc.au3"

; #INDEX# =======================================================================================================================
; Title .........: LibreOffice UDF
; AutoIt Version : v3.3.16.1
; Description ...: Provides basic functionality through AutoIt for Inserting and Modifying L.O. Writer Fields.
; Author(s) .....: donnyh13, mLipok
; Dll ...........:
;
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _LOWriter_FieldAuthorInsert
; _LOWriter_FieldAuthorModify
; _LOWriter_FieldChapterInsert
; _LOWriter_FieldChapterModify
; _LOWriter_FieldCombCharInsert
; _LOWriter_FieldCombCharModify
; _LOWriter_FieldCommentInsert
; _LOWriter_FieldCommentModify
; _LOWriter_FieldCondTextInsert
; _LOWriter_FieldCondTextModify
; _LOWriter_FieldCurrentDisplayGet
; _LOWriter_FieldDateTimeInsert
; _LOWriter_FieldDateTimeModify
; _LOWriter_FieldDelete
; _LOWriter_FieldDocInfoCommentsInsert
; _LOWriter_FieldDocInfoCommentsModify
; _LOWriter_FieldDocInfoCreateAuthInsert
; _LOWriter_FieldDocInfoCreateAuthModify
; _LOWriter_FieldDocInfoCreateDateTimeInsert
; _LOWriter_FieldDocInfoCreateDateTimeModify
; _LOWriter_FieldDocInfoEditTimeInsert
; _LOWriter_FieldDocInfoEditTimeModify
; _LOWriter_FieldDocInfoKeywordsInsert
; _LOWriter_FieldDocInfoKeywordsModify
; _LOWriter_FieldDocInfoModAuthInsert
; _LOWriter_FieldDocInfoModAuthModify
; _LOWriter_FieldDocInfoModDateTimeInsert
; _LOWriter_FieldDocInfoModDateTimeModify
; _LOWriter_FieldDocInfoPrintAuthInsert
; _LOWriter_FieldDocInfoPrintAuthModify
; _LOWriter_FieldDocInfoPrintDateTimeInsert
; _LOWriter_FieldDocInfoPrintDateTimeModify
; _LOWriter_FieldDocInfoRevNumInsert
; _LOWriter_FieldDocInfoRevNumModify
; _LOWriter_FieldDocInfoSubjectInsert
; _LOWriter_FieldDocInfoSubjectModify
; _LOWriter_FieldDocInfoTitleInsert
; _LOWriter_FieldDocInfoTitleModify
; _LOWriter_FieldFileNameInsert
; _LOWriter_FieldFileNameModify
; _LOWriter_FieldFuncHiddenParInsert
; _LOWriter_FieldFuncHiddenParModify
; _LOWriter_FieldFuncHiddenTextInsert
; _LOWriter_FieldFuncHiddenTextModify
; _LOWriter_FieldFuncInputInsert
; _LOWriter_FieldFuncInputModify
; _LOWriter_FieldFuncPlaceholderInsert
; _LOWriter_FieldFuncPlaceholderModify
; _LOWriter_FieldGetAnchor
; _LOWriter_FieldInputListInsert
; _LOWriter_FieldInputListModify
; _LOWriter_FieldPageNumberInsert
; _LOWriter_FieldPageNumberModify
; _LOWriter_FieldRefBookMarkInsert
; _LOWriter_FieldRefBookMarkModify
; _LOWriter_FieldRefEndnoteInsert
; _LOWriter_FieldRefEndnoteModify
; _LOWriter_FieldRefFootnoteInsert
; _LOWriter_FieldRefFootnoteModify
; _LOWriter_FieldRefGetType
; _LOWriter_FieldRefInsert
; _LOWriter_FieldRefMarkDelete
; _LOWriter_FieldRefMarkGetAnchor
; _LOWriter_FieldRefMarkList
; _LOWriter_FieldRefMarkSet
; _LOWriter_FieldRefModify
; _LOWriter_FieldsAdvGetList
; _LOWriter_FieldsDocInfoGetList
; _LOWriter_FieldSenderInsert
; _LOWriter_FieldSenderModify
; _LOWriter_FieldSetVarInsert
; _LOWriter_FieldSetVarMasterCreate
; _LOWriter_FieldSetVarMasterDelete
; _LOWriter_FieldSetVarMasterExists
; _LOWriter_FieldSetVarMasterGetObj
; _LOWriter_FieldSetVarMasterList
; _LOWriter_FieldSetVarMasterListFields
; _LOWriter_FieldSetVarModify
; _LOWriter_FieldsGetList
; _LOWriter_FieldShowVarInsert
; _LOWriter_FieldShowVarModify
; _LOWriter_FieldStatCountInsert
; _LOWriter_FieldStatCountModify
; _LOWriter_FieldStatTemplateInsert
; _LOWriter_FieldStatTemplateModify
; _LOWriter_FieldUpdate
; _LOWriter_FieldVarSetPageInsert
; _LOWriter_FieldVarSetPageModify
; _LOWriter_FieldVarShowPageInsert
; _LOWriter_FieldVarShowPageModify
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldAuthorInsert
; Description ...: Insert a Author Field.
; Syntax ........: _LOWriter_FieldAuthorInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sAuthor = Null[, $bFullName = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author Name to insert. Note, $bIsFixed must be set to True for this value to stay the same as set.
;                  $bFullName           - [optional] a boolean value. Default is Null. If True, displays the full name. Else Initials. For a Fixed custom name, this does nothing.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sAuthor not a String.
;				   @Error 1 @Extended 7 Return 0 = $bFullName not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.Author" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Author field, returning Author Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldAuthorModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldAuthorInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sAuthor = Null, $bFullName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oAuthField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oAuthField = $oDoc.createInstance("com.sun.star.text.TextField.Author")
	If Not IsObj($oAuthField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oAuthField.IsFixed = $bIsFixed
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oAuthField.Content = $sAuthor
	EndIf

	If ($bFullName <> Null) Then
		If Not IsBool($bFullName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oAuthField.FullName = $bFullName
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oAuthField, $bOverwrite)

	If ($sAuthor <> Null) Then ; Sometimes Author Disappears upon Insertion, make a check to re-set the Author value.
		If $oAuthField.Content <> $sAuthor And ($oAuthField.IsFixed() = True) Then $oAuthField.Content = $sAuthor
	EndIf

	If ($oAuthField.IsFixed() = False) Then $oAuthField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oAuthField)
EndFunc   ;==>_LOWriter_FieldAuthorInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldAuthorModify
; Description ...: Set or Retrieve a Author Field's settings.
; Syntax ........: _LOWriter_FieldAuthorModify(ByRef $oAuthField[, $bIsFixed = Null[, $sAuthor = Null[, $bFullName = Null]]])
; Parameters ....: $oAuthField          - [in/out] an object. A Author field Object from _LOWriter_FieldAuthorInsert, or _LOWriter_FieldsGetList function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author Name to insert. Note, $bIsFixed must be set to True for this value to stay the same as set.
;                  $bFullName           - [optional] a boolean value. Default is Null. If True, displays the full name. Else Initials.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oAuthField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sAuthor not a String.
;				   @Error 1 @Extended 4 Return 0 = $bFullName not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sAuthor
;				   |								4 = Error setting $bFullName
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldAuthorInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldAuthorModify(ByRef $oAuthField, $bIsFixed = Null, $sAuthor = Null, $bFullName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avAuth[3]

	If Not IsObj($oAuthField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sAuthor, $bFullName) Then
		__LOWriter_ArrayFill($avAuth, $oAuthField.IsFIxed(), $oAuthField.Content(), $oAuthField.FullName())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avAuth)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oAuthField.IsFIxed = $bIsFixed
		$iError = ($oAuthField.IsFIxed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oAuthField.Content = $sAuthor
		$iError = ($oAuthField.Content() = $sAuthor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bFullName <> Null) Then
		If Not IsBool($bFullName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oAuthField.FullName = $bFullName
		$iError = ($oAuthField.FullName() = $bFullName) ? ($iError) : (BitOR($iError, 4))
	EndIf

	$oAuthField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldAuthorModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldChapterInsert
; Description ...: Insert a Chapter Field.
; Syntax ........: _LOWriter_FieldChapterInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $iChapFrmt = Null[, $iLevel = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iChapFrmt           - [optional] an integer value (0-4). Default is Null. The Display format for the Chapter Field. See Constants, $LOW_FIELD_CHAP_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLevel              - [optional] an integer value (1-10). Default is Null. The Chapter level to display.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iChapFrmt not an integer, less than 0, or greater than 4. See Constants, $LOW_FIELD_CHAP_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $iLevel not an Integer, less than 1, or greater than 10.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.Chapter" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Chapter field, returning Chapter Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldChapterModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldChapterInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $iChapFrmt = Null, $iLevel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oChapField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oChapField = $oDoc.createInstance("com.sun.star.text.TextField.Chapter")
	If Not IsObj($oChapField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($iChapFrmt <> Null) Then
		If Not __LOWriter_IntIsBetween($iChapFrmt, $LOW_FIELD_CHAP_FRMT_NAME, $LOW_FIELD_CHAP_FRMT_DIGIT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oChapField.ChapterFormat = $iChapFrmt
	EndIf

	If ($iLevel <> Null) Then
		If Not __LOWriter_IntIsBetween($iLevel, 1, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oChapField.Level = ($iLevel - 1) ; Level is 0 Based
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oChapField, $bOverwrite)

	$oChapField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oChapField)
EndFunc   ;==>_LOWriter_FieldChapterInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldChapterModify
; Description ...: Set or Retrieve a Chapter Field's settings.
; Syntax ........: _LOWriter_FieldChapterModify(ByRef $oChapField[, $iChapFrmt = Null[, $iLevel = Null]])
; Parameters ....: $oChapField          - [in/out] an object. A Chapter field Object from a previous_LOWriter_FieldChapterInsert,  or _LOWriter_FieldsGetList function.
;                  $iChapFrmt           - [optional] an integer value (0-4). Default is Null. The Display format for the Chapter Field. See Constants, $LOW_FIELD_CHAP_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iLevel              - [optional] an integer value (1-10). Default is Null. The Chapter level to display.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oChapField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iChapFrmt not an integer, less than 0, or greater than 4. See Constants, $LOW_FIELD_CHAP_FRMT_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $iLevel not an Integer, less than 1, or greater than 10.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iChapFrmt
;				   |								2 = Error setting $iLevel
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldChapterInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldChapterModify(ByRef $oChapField, $iChapFrmt = Null, $iLevel = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $aiChap[2]

	If Not IsObj($oChapField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iChapFrmt, $iLevel) Then
		__LOWriter_ArrayFill($aiChap, $oChapField.ChapterFormat(), ($oChapField.Level() + 1)) ; Level is 0 Based -- Add 1 to make it like L.O. UI
		Return SetError($__LO_STATUS_SUCCESS, 1, $aiChap)
	EndIf

	If ($iChapFrmt <> Null) Then
		If Not __LOWriter_IntIsBetween($iChapFrmt, $LOW_FIELD_CHAP_FRMT_NAME, $LOW_FIELD_CHAP_FRMT_DIGIT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oChapField.ChapterFormat = $iChapFrmt
		$iError = ($oChapField.ChapterFormat() = $iChapFrmt) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iLevel <> Null) Then
		If Not __LOWriter_IntIsBetween($iLevel, 1, 10) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oChapField.Level = ($iLevel - 1) ; Level is 0 Based
		$iError = ($oChapField.Level() = ($iLevel - 1)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	$oChapField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldChapterModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCombCharInsert
; Description ...: Insert a Combined Character Field.
; Syntax ........: _LOWriter_FieldCombCharInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $sCharacters = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $sCharacters         - [optional] a string value. Default is Null. The Characters to insert in a combined character field. Max 6 character String Length.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sCharacters not a String.
;				   @Error 1 @Extended 6 Return 0 = $sCharacters longer than 6 characters.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.CombinedCharacters" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Combined Character field, returning Combined Character Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldCombCharModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCombCharInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sCharacters = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCombCharField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oCombCharField = $oDoc.createInstance("com.sun.star.text.TextField.CombinedCharacters")
	If Not IsObj($oCombCharField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sCharacters <> Null) Then
		If Not IsString($sCharacters) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If (StringLen($sCharacters) > 6) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oCombCharField.Content = $sCharacters
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oCombCharField, $bOverwrite)

	$oCombCharField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCombCharField)
EndFunc   ;==>_LOWriter_FieldCombCharInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCombCharModify
; Description ...: Set or Retrieve a Combined Character Field's settings.
; Syntax ........: _LOWriter_FieldCombCharModify(ByRef $oCombCharField[, $sCharacters = Null])
; Parameters ....: $oCombCharField      - [in/out] an object. A Combined Character field Object from a previous _LOWriter_FieldCombCharInsert, or _LOWriter_FieldsGetList( function.
;                  $sCharacters         - [optional] a string value. Default is Null. The Characters to insert in a combined character field. Max 6 character String Length.
; Return values .: Success: 1 or String.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sCharacters not a String.
;				   @Error 1 @Extended 3 Return 0 = String called in $sCharacters longer than 6 characters.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sCharacters
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return String = Success. All optional parameters were set to Null, returning current Combined Characters value.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldCombCharInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCombCharModify(ByRef $oCombCharField, $sCharacters = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oCombCharField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sCharacters) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oCombCharField.Content())

	If Not IsString($sCharacters) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (StringLen($sCharacters) > 6) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	$oCombCharField.Content = $sCharacters
	$iError = ($oCombCharField.Content() = $sCharacters) ? ($iError) : (BitOR($iError, 1))

	$oCombCharField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldCombCharModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCommentInsert
; Description ...: Insert a Comment field into a document at a cursor's position.
; Syntax ........: _LOWriter_FieldCommentInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $sContent = Null[, $sAuthor = Null[, $tDateStruct = Null[, $sInitials = Null[, $sName = Null[, $bResolved = Null]]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $sContent            - [optional] a string value. Default is Null. The content of the comment.
;                  $sAuthor             - [optional] a string value. Default is Null. The author of the comment.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display for the comment, created previously by _LOWriter_DateStructCreate. If left as Null, the current date is used.
;                  $sInitials           - [optional] a string value. Default is Null. The Initials of the creator. Libre Office version 4.0 and up only.
;                  $sName               - [optional] a string value. Default is Null. The name of the creator. Libre Office version 4.0 and up only.
;                  $bResolved           - [optional] a boolean value. Default is Null. If True, the comment is marked as resolved.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sContent not a String.
;				   @Error 1 @Extended 6 Return 0 = $sAuthor not a String.
;				   @Error 1 @Extended 7 Return 0 = $tDateStruct not an Object.
;				   @Error 1 @Extended 8 Return 0 = $sInitials not a String.
;				   @Error 1 @Extended 9 Return 0 = $sName not a String.
;				   @Error 1 @Extended 10 Return 0 = $bResolved not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.Annotation" Object.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office Version lower than 4.0.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted comment field, returning Comment Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldCommentModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DateStructCreate _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCommentInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sContent = Null, $sAuthor = Null, $tDateStruct = Null, $sInitials = Null, $sName = Null, $bResolved = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCommentField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oCommentField = $oDoc.createInstance("com.sun.star.text.TextField.Annotation")
	If Not IsObj($oCommentField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sContent <> Null) Then
		If Not IsString($sContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oCommentField.Content = $sContent
	Else
		$oCommentField.Content = " " ;If Content is Blank, Comment/Annotation will disappear.
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oCommentField.Author = $sAuthor
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oCommentField.DateTimeValue = $tDateStruct
	Else
		$oCommentField.DateTimeValue = _LOWriter_DateStructCreate()
	EndIf

	If ($sInitials <> Null) Then
		If Not IsString($sInitials) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		If Not __LOWriter_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
		$oCommentField.Initials = $sInitials
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If Not __LOWriter_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
		$oCommentField.Name = $sName
	EndIf

	If ($bResolved <> Null) Then
		If Not IsBool($bResolved) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
		$oCommentField.Resolved = $bResolved
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oCommentField, $bOverwrite)

	$oCommentField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCommentField)
EndFunc   ;==>_LOWriter_FieldCommentInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCommentModify
; Description ...: Set or retrieve Comment settings.
; Syntax ........: _LOWriter_FieldCommentModify(ByRef $oDoc, ByRef $oCommentField[, $sContent = Null[, $sAuthor = Null[, $tDateStruct = Null[, $sInitials = Null[, $sName = Null[, $bResolved = Null]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCommentField       - [in/out] an object. A Comment field Object from a previous _LOWriter_FieldCommentInsert, or _LOWriter_FieldsGetList function.
;                  $sContent            - [optional] a string value. Default is Null. The content of the comment.
;                  $sAuthor             - [optional] a string value. Default is Null. The author of the comment.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display for the comment, created previously by _LOWriter_DateStructCreate.
;                  $sInitials           - [optional] a string value. Default is Null. The Initials of the creator. Libre Office version 4.0 and up only.
;                  $sName               - [optional] a string value. Default is Null. The name of the creator. Libre Office version 4.0 and up only.
;                  $bResolved           - [optional] a boolean value. Default is Null. If True, the comment is marked as resolved.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCommentField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sContent not a String.
;				   @Error 1 @Extended 4 Return 0 = $sAuthor not a String.
;				   @Error 1 @Extended 5 Return 0 = $tDateStruct not an Object.
;				   @Error 1 @Extended 6 Return 0 = $sInitials not a String.
;				   @Error 1 @Extended 7 Return 0 = $sName not a String.
;				   @Error 1 @Extended 8 Return 0 = $bResolved not a Boolean.
;				   --Version Related Errors--
;				   @Error 7 @Extended 1 Return 0 = Current Libre Office Version lower than 4.0.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sContent
;				   |								2 = Error setting $sAuthor
;				   |								4 = Error setting $tDateStruct
;				   |								8 = Error setting $sInitials
;				   |								16 = Error setting $sName
;				   |								32 = Error setting $bResolved
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array If L.O. version is less than 4.0, else a 6 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldCommentInsert, _LOWriter_FieldsGetList, _LOWriter_DateStructCreate _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCommentModify(ByRef $oDoc, ByRef $oCommentField, $sContent = Null, $sAuthor = Null, $tDateStruct = Null, $sInitials = Null, $sName = Null, $bResolved = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avAnnot[4]
	Local $bRefresh = False

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCommentField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($sContent, $sAuthor, $tDateStruct, $sInitials, $sName, $bResolved) Then
		If __LOWriter_VersionCheck(4.0) Then
			__LOWriter_ArrayFill($avAnnot, $oCommentField.Content(), $oCommentField.Author(), $oCommentField.DateTimeValue(), $oCommentField.Initials(), _
					$oCommentField.Name(), $oCommentField.Resolved())
		Else
			__LOWriter_ArrayFill($avAnnot, $oCommentField.Content(), $oCommentField.Author(), $oCommentField.DateTimeValue(), $oCommentField.Resolved())
		EndIf
		Return SetError($__LO_STATUS_SUCCESS, 1, $avAnnot)
	EndIf

	If ($sContent <> Null) Then
		If Not IsString($sContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oCommentField.Content = $sContent
		$iError = ($oCommentField.Content() = $sContent) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oCommentField.Author = $sAuthor
		$iError = ($oCommentField.Author() = $sAuthor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oCommentField.DateTimeValue = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oCommentField.DateTimeValue(), $tDateStruct)) ? ($iError) : (BitOR($iError, 4))
		$bRefresh = True
	EndIf

	If ($sInitials <> Null) Then
		If Not IsString($sInitials) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not __LOWriter_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
		$oCommentField.Initials = $sInitials
		$iError = ($oCommentField.Initials() = $sInitials) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		If Not __LOWriter_VersionCheck(4.0) Then Return SetError($__LO_STATUS_VER_ERROR, 1, 0)
		$oCommentField.Name = $sName
		$iError = ($oCommentField.Name = $sName) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($bResolved <> Null) Then
		If Not IsBool($bResolved) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$oCommentField.Resolved = $bResolved
		$iError = ($oCommentField.Resolved() = $bResolved) ? ($iError) : (BitOR($iError, 32))
		$bRefresh = True
	EndIf

	If ($bRefresh = True) Then _
		$oDoc.Text.createTextCursorByRange($oCommentField.Anchor()).Text.insertTextContent($oCommentField.Anchor(), $oCommentField, True)

	$oCommentField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldCommentModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCondTextInsert
; Description ...: Insert a Conditional Text Field.
; Syntax ........: _LOWriter_FieldCondTextInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $sCondition = Null[, $sThen = Null[, $sElse = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $sCondition          - [optional] a string value. Default is Null. The condition to test.
;                  $sThen               - [optional] a string value. Default is Null. The text to display if the condition is True.
;                  $sElse               - [optional] a string value. Default is Null. The text to display if the condition is False.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sCondition not a String.
;				   @Error 1 @Extended 6 Return 0 = $sThen not a String.
;				   @Error 1 @Extended 7 Return 0 = $sElse not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.ConditionalText" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted a Conditional Text field, returning the Conditional Text Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldCondTextModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCondTextInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sCondition = Null, $sThen = Null, $sElse = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCondTextField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oCondTextField = $oDoc.createInstance("com.sun.star.text.TextField.ConditionalText")
	If Not IsObj($oCondTextField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oCondTextField.Condition = $sCondition
	EndIf

	If ($sThen <> Null) Then
		If Not IsString($sThen) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oCondTextField.TrueContent = $sThen
	EndIf

	If ($sElse <> Null) Then
		If Not IsString($sElse) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oCondTextField.FalseContent = $sElse
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oCondTextField, $bOverwrite)

	$oCondTextField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCondTextField)
EndFunc   ;==>_LOWriter_FieldCondTextInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCondTextModify
; Description ...: Set or Retrieve a Conditional Text Field's settings.
; Syntax ........: _LOWriter_FieldCondTextModify(ByRef $oCondTextField[, $sCondition = Null[, $sThen = Null[, $sElse = Null]]])
; Parameters ....: $oCondTextField      - [in/out] an object. A Conditional Text field Object from a previous _LOWriter_FieldCondTextInsert, or _LOWriter_FieldsGetList function.
;                  $sCondition          - [optional] a string value. Default is Null. The condition to test.
;                  $sThen               - [optional] a string value. Default is Null. The text to display if the condition is True.
;                  $sElse               - [optional] a string value. Default is Null. The text to display if the condition is False.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sCondition not a String.
;				   @Error 1 @Extended 3 Return 0 = $sThen not a String.
;				   @Error 1 @Extended 4 Return 0 = $sElse not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sCondition
;				   |								2 = Error setting $sThen
;				   |								4 = Error setting $sElse
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters, with an additional parameter in the last element to indicate if the condition is evaluated as True or not.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldCondTextInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCondTextModify(ByRef $oCondTextField, $sCondition = Null, $sThen = Null, $sElse = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avCond[4]

	If Not IsObj($oCondTextField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sCondition, $sThen, $sElse) Then
		__LOWriter_ArrayFill($avCond, $oCondTextField.Condition(), $oCondTextField.TrueContent(), $oCondTextField.FalseContent(), _
				($oCondTextField.IsConditionTrue()) ? (False) : (True)) ; IsConditionTrue is Backwards.
		Return SetError($__LO_STATUS_SUCCESS, 1, $avCond)
	EndIf

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oCondTextField.Condition = $sCondition
		$iError = ($oCondTextField.Condition() = $sCondition) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sThen <> Null) Then
		If Not IsString($sThen) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oCondTextField.TrueContent = $sThen
		$iError = ($oCondTextField.TrueContent() = $sThen) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sElse <> Null) Then
		If Not IsString($sElse) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oCondTextField.FalseContent = $sElse
		$iError = ($oCondTextField.FalseContent() = $sElse) ? ($iError) : (BitOR($iError, 4))
	EndIf

	$oCondTextField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldCondTextModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldCurrentDisplayGet
; Description ...: Retrieve the current data displayed by a field.
; Syntax ........: _LOWriter_FieldCurrentDisplayGet(ByRef $oField)
; Parameters ....: $oField              - [in/out] an object. A Field Object returned from a previous insert, _LOWriter_FieldsGetList, _LOWriter_FieldsDocInfoGetList function.
; Return values .: Success: String
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oField not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return String = Success. Returning current Field display content in String format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Note, a Comment Field will return an empty string, use the Comment Field function to retrieve the current
;					comment content. A DocInfoComments field will work with this function however.
;					Note: This will work for most Fields, but not all. Check and see which will work and which wont.
; Related .......: _LOWriter_FieldsGetList, _LOWriter_FieldsAdvGetList, _LOWriter_FieldsDocInfoGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldCurrentDisplayGet(ByRef $oField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $sPresentation

	If Not IsObj($oField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If ($oField.supportsService("com.sun.star.text.textfield.ConditionalText")) Then ; Conditional Text Fields don't update "CurrentPresentation" setting,
		; so acquire the current display based on whether the condition is true or not.
		$sPresentation = ($oField.IsConditionTrue() = False) ? ($oField.TrueContent()) : ($oField.FalseContent())
	Else
		$sPresentation = $oField.CurrentPresentation()
	EndIf

	Return SetError($__LO_STATUS_SUCCESS, 0, $sPresentation)
EndFunc   ;==>_LOWriter_FieldCurrentDisplayGet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDateTimeInsert
; Description ...: Insert a Date and/or Time Field.
; Syntax ........: _LOWriter_FieldDateTimeInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $tDateStruct = Null[, $bIsDate = Null[, $iOffset = Null[, $iDateFormatKey = Null]]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display for the comment, created previously by _LOWriter_DateStructCreate.
;                  $bIsDate             - [optional] a boolean value. Default is Null. If True, the field is considered as containing a Date, $iOffset will be evaluated in Days. Else if False, Field will be considered as containing a Time, $iOffset will be evaluated in minutes.
;                  $iOffset             - [optional] an integer value. Default is Null. The offset to apply to the date, either in Minutes or Days, depending on the current $bIsDate setting.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList function.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $tDateStruct not an Object.
;				   @Error 1 @Extended 7 Return 0 = $bIsDate not a Boolean.
;				   @Error 1 @Extended 8 Return 0 = $iOffset not an Integer.
;				   @Error 1 @Extended 9 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 10 Return 0 = $iDateFormatKey not found in current Document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.DateTime" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Date/Time field, returning Date/Time Field Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDateTimeModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DateStructCreate, _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDateTimeInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $tDateStruct = Null, $bIsDate = Null, $iOffset = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDateTimeField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDateTimeField = $oDoc.createInstance("com.sun.star.text.TextField.DateTime")
	If Not IsObj($oDateTimeField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDateTimeField.IsFixed = $bIsFixed
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oDateTimeField.DateTimeValue = $tDateStruct
	EndIf

	If ($bIsDate <> Null) Then
		If Not IsBool($bIsDate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oDateTimeField.IsDate = $bIsDate
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$oDateTimeField.Adjust = ($oDateTimeField.IsDate() = True) ? (Int((1440 * $iOffset))) : ($iOffset)
		; If IsDate = True, Then Calculate number of minutes in a day (1440) times number of days to off set the Date/ Value,
		; else, just set it to Number of minutes called.
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 10, 0)
		$oDateTimeField.NumberFormat = $iDateFormatKey
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDateTimeField, $bOverwrite)

	If ($tDateStruct <> Null) Then ; Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If (__LOWriter_DateStructCompare($oDateTimeField.DateTimeValue(), $tDateStruct) = False) And ($oDateTimeField.IsFixed() = True) Then $oDateTimeField.DateTimeValue = $tDateStruct
	EndIf

	If ($oDateTimeField.IsFixed() = False) Then $oDateTimeField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDateTimeField)
EndFunc   ;==>_LOWriter_FieldDateTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDateTimeModify
; Description ...: Set or Retrieve a Date/Time Field's settings.
; Syntax ........: _LOWriter_FieldDateTimeModify(ByRef $oDoc, ByRef $oDateTimeField[, $bIsFixed = Null[, $tDateStruct = Null[, $bIsDate = Null[, $iOffset = Null[, $iDateFormatKey = Null]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oDateTimeField      - [in/out] an object. A Date/Time field Object from a previous _LOWriter_FieldDateTimeInsert, or _LOWriter_FieldsGetList function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $tDateStruct         - [optional] a dll struct value. Default is Null. The date to display for the comment, created previously by _LOWriter_DateStructCreate.
;                  $bIsDate             - [optional] a boolean value. Default is Null. If True, the field is considered as containing a Date, $iOffset will be evaluated in Days. Else False, Field will be considered as containing a Time, $iOffset will be evaluated in minutes.
;                  $iOffset             - [optional] an integer value. Default is Null. The offset to apply to the date, either in Minutes or Days, depending on the current $bIsDate setting.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $tDateStruct not an Object.
;				   @Error 1 @Extended 4 Return 0 = $bIsDate not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iOffset not an Integer.
;				   @Error 1 @Extended 6 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iDateFormatKey not found in current Document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $tDateStruct
;				   |								4 = Error setting $bIsDate
;				   |								8 = Error setting $iOffset
;				   |								16 = Error setting $iDateFormatKey
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 5 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDateTimeInsert, _LOWriter_FieldsGetList, _LOWriter_DateStructCreate, _LOWriter_DateStructModify
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDateTimeModify(ByRef $oDoc, ByRef $oDateTimeField, $bIsFixed = Null, $tDateStruct = Null, $bIsDate = Null, $iOffset = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avDateTime[5]

	If Not IsObj($oDateTimeField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $tDateStruct, $bIsDate, $iOffset, $iDateFormatKey) Then
		; Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		; fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		; from the value.
		$iNumberFormat = $oDateTimeField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : ($iNumberFormat)

		__LOWriter_ArrayFill($avDateTime, $oDateTimeField.IsFixed(), $oDateTimeField.DateTimeValue(), $oDateTimeField.IsDate(), _
				($oDateTimeField.IsDate() = True) ? (Int(($oDateTimeField.Adjust() / 1440))) : ($oDateTimeField.Adjust()), $iNumberFormat)
		; If IsDate = True, Then Calculate number of minutes in a day (1440) divided by number of days of off set. Otherwise
		; return Number of minutes.
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDateTime)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oDateTimeField.IsFixed = $bIsFixed
		$iError = ($oDateTimeField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($tDateStruct <> Null) Then
		If Not IsObj($tDateStruct) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDateTimeField.DateTimeValue = $tDateStruct
		$iError = (__LOWriter_DateStructCompare($oDateTimeField.DateTimeValue(), $tDateStruct)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bIsDate <> Null) Then
		If Not IsBool($bIsDate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oDateTimeField.IsDate = $bIsDate
		$iError = ($oDateTimeField.IsDate() = $bIsDate) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$iOffset = ($oDateTimeField.IsDate() = True) ? Int((1440 * $iOffset)) : $iOffset
		; If IsDate = True, Then Calculate number of minutes in a day (1440) times number of days to off set the Date/ Value,
		; else, just set it to Number of minutes called.

		$oDateTimeField.Adjust = $iOffset
		$iError = ($oDateTimeField.Adjust() = $iOffset) ? ($iError) : (BitOR($iError, 8))
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oDateTimeField.NumberFormat = $iDateFormatKey
		$iError = ($oDateTimeField.NumberFormat() = ($iDateFormatKey)) ? ($iError) : (BitOR($iError, 16))
	EndIf

	If ($oDateTimeField.IsFixed() = False) Then $oDateTimeField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDateTimeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDelete
; Description ...: Delete a Field from a Document.
; Syntax ........: _LOWriter_FieldDelete(ByRef $oDoc, ByRef $oField[, $bDeleteMaster = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oField              - [in/out] an object. A Field Object from a previous Insert, _LOWriter_FieldsGetList, _LOWriter_FieldsAdvGetList, or _LOWriter_FieldsDocInfoGetList function.
;                  $bDeleteMaster       - [optional] a boolean value. Default is False. If True, and the field has a Master Field, the MasterField (With any other dependent fields) will be deleted.
; Return values .: Success: 1.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bDeleteMaster not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving TextFieldMaster Object.
;				   @Error 2 @Extended 2 Return 0 = Error retrieving Field Master Array of dependent fields.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully deleted the field with the Text Master Field.
;				   @Error 0 @Extended 1 Return 1 = Success. Successfully deleted the field.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldsGetList, _LOWriter_FieldsAdvGetList, _LOWriter_FieldsDocInfoGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDelete(ByRef $oDoc, ByRef $oField, $bDeleteMaster = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFieldMaster
	Local $aoDependents[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bDeleteMaster) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($bDeleteMaster = True) And ($oField.TextFieldMaster.Name() <> "") Then
		$oFieldMaster = $oField.TextFieldMaster()
		If Not IsObj($oFieldMaster) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		$aoDependents = $oFieldMaster.DependentTextFields()
		If Not IsArray($aoDependents) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

		If (UBound($aoDependents) > 0) Then
			For $i = 0 To UBound($aoDependents) - 1
				$aoDependents[$i].Anchor.Text.removeTextContent($aoDependents[$i])
				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
			Next
		EndIf

		$oFieldMaster.dispose()
		Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	EndIf

	$oField.Anchor.Text.removeTextContent($oField)

	Return SetError($__LO_STATUS_SUCCESS, 1, 1)
EndFunc   ;==>_LOWriter_FieldDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCommentsInsert
; Description ...: Insert a Document Information Comments Field.
; Syntax ........: _LOWriter_FieldDocInfoCommentsInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sComments = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sComments           - [optional] a string value. Default is Null. The Comments text to display, note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sComments not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.Description" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Comments Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoCommentsModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCommentsInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sComments = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoCommentField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoCommentField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.Description")
	If Not IsObj($oDocInfoCommentField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoCommentField.IsFixed = $bIsFixed
	EndIf

	If ($sComments <> Null) Then
		If Not IsString($sComments) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoCommentField.Content = $sComments
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoCommentField, $bOverwrite)

	If ($sComments <> Null) Then ; Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oDocInfoCommentField.Content <> $sComments And ($oDocInfoCommentField.IsFixed() = True) Then $oDocInfoCommentField.Content = $sComments
	EndIf

	If ($oDocInfoCommentField.IsFixed() = False) Then $oDocInfoCommentField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoCommentField)
EndFunc   ;==>_LOWriter_FieldDocInfoCommentsInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCommentsModify
; Description ...: Set or Retrieve a Document Information Comments Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoCommentsModify(ByRef $oDocInfoCommentField[, $bIsFixed = Null[, $sComments = Null]])
; Parameters ....: $oDocInfoCommentField  - [in/out] an object. A Doc Info Comments field Object from a previous _LOWriter_FieldDocInfoCommentsInsert, or _LOWriter_FieldsGetList function.
;                  $bIsFixed              - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sComments             - [optional] a string value. Default is Null. The Comments text to display, note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoCommentField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sComments not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sComments
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoCommentsInsert, _LOWriter_FieldsGetList, _LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCommentsModify(ByRef $oDocInfoCommentField, $bIsFixed = Null, $sComments = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoCom[2]

	If Not IsObj($oDocInfoCommentField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sComments) Then
		__LOWriter_ArrayFill($avDocInfoCom, $oDocInfoCommentField.IsFixed(), $oDocInfoCommentField.Content())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoCom)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoCommentField.IsFixed = $bIsFixed
		$iError = ($oDocInfoCommentField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sComments <> Null) Then
		If Not IsString($sComments) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoCommentField.Content = $sComments
		$iError = ($oDocInfoCommentField.Content() = $sComments) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoCommentField.IsFixed() = False) Then $oDocInfoCommentField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoCommentsModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCreateAuthInsert
; Description ...: Insert a Document Information Create Author Field.
; Syntax ........: _LOWriter_FieldDocInfoCreateAuthInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sAuthor = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed must be set to True in order for this to remain as set.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sAuthor not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.CreateAuthor" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Created By Author Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoCreateAuthModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocGenPropCreation
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCreateAuthInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoCreateAuthField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoCreateAuthField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.CreateAuthor")
	If Not IsObj($oDocInfoCreateAuthField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoCreateAuthField.IsFixed = $bIsFixed
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoCreateAuthField.Author = $sAuthor
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoCreateAuthField, $bOverwrite)

	If ($sAuthor <> Null) Then ; Sometimes Author Disappears upon Insertion, make a check to re-set the Author value.
		If $oDocInfoCreateAuthField.Author <> $sAuthor And ($oDocInfoCreateAuthField.IsFixed() = True) Then $oDocInfoCreateAuthField.Author = $sAuthor
	EndIf

	If ($oDocInfoCreateAuthField.IsFixed() = False) Then $oDocInfoCreateAuthField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoCreateAuthField)
EndFunc   ;==>_LOWriter_FieldDocInfoCreateAuthInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCreateAuthModify
; Description ...: Set or Retrieve a Document Information Create Author Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoCreateAuthModify(ByRef $oDocInfoCreateAuthField[, $bIsFixed = Null[, $sAuthor = Null]])
; Parameters ....: $oDocInfoCreateAuthField  - [in/out] an object. A Created By Author field Object from a previous _LOWriter_FieldDocInfoCreateAuthInsert or _LOWriter_FieldsDocInfoGetList function.
;                  $bIsFixed                 - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sAuthor                  - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed must be set to True in order for this to remain as set.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoCreateAuthField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sAuthor not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sAuthor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoCreateAuthInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocGenPropCreation
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCreateAuthModify(ByRef $oDocInfoCreateAuthField, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoModAuth[2]

	If Not IsObj($oDocInfoCreateAuthField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sAuthor) Then
		__LOWriter_ArrayFill($avDocInfoModAuth, $oDocInfoCreateAuthField.IsFixed(), $oDocInfoCreateAuthField.Author())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoModAuth)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoCreateAuthField.IsFixed = $bIsFixed
		$iError = ($oDocInfoCreateAuthField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoCreateAuthField.Author = $sAuthor
		$iError = ($oDocInfoCreateAuthField.Author() = $sAuthor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoCreateAuthField.IsFixed() = False) Then $oDocInfoCreateAuthField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoCreateAuthModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCreateDateTimeInsert
; Description ...: Insert a Document Information Create Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoCreateDateTimeInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iDateFormatKey = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iDateFormatKey not found in document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.CreateDateTime" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Created Date/Time Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoCreateDateTimeModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyList, _LOWriter_DocGenPropCreation
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCreateDateTimeInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoCreateDtTmField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoCreateDtTmField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.CreateDateTime")
	If Not IsObj($oDocInfoCreateDtTmField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoCreateDtTmField.IsFixed = $bIsFixed
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oDocInfoCreateDtTmField.NumberFormat = $iDateFormatKey
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoCreateDtTmField, $bOverwrite)

	If ($oDocInfoCreateDtTmField.IsFixed() = False) Then $oDocInfoCreateDtTmField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoCreateDtTmField)
EndFunc   ;==>_LOWriter_FieldDocInfoCreateDateTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoCreateDateTimeModify
; Description ...: Set or Retrieve a Document Information Create Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoCreateDateTimeModify(ByRef $oDoc, ByRef $oDocInfoCreateDtTmField[, $bIsFixed = Null[, $iDateFormatKey = Null]])
; Parameters ....: $oDoc                     - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oDocInfoCreateDtTmField  - [in/out] an object. A Created at Date/Time field Object from a previous _LOWriter_FieldDocInfoCreateDateTimeInsert, or _LOWriter_FieldsDocInfoGetList function.
;                  $bIsFixed                 - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey           - [optional] an integer value. Default is Null. A Date or Time Format Key returned from a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oDocInfoCreateDtTmField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iDateFormatKey not found in document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iDateFormatKey
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoCreateDateTimeInsert, _LOWriter_FieldsDocInfoGetList,
;					_LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyList, _LOWriter_DocGenPropCreation
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoCreateDateTimeModify(ByRef $oDoc, ByRef $oDocInfoCreateDtTmField, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avDocInfoCrtDate[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDocInfoCreateDtTmField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $iDateFormatKey) Then
		; Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		; fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000 from
		; the value.
		$iNumberFormat = $oDocInfoCreateDtTmField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : ($iNumberFormat)

		__LOWriter_ArrayFill($avDocInfoCrtDate, $oDocInfoCreateDtTmField.IsFixed(), $iNumberFormat)
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoCrtDate)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoCreateDtTmField.IsFixed = $bIsFixed
		$iError = ($oDocInfoCreateDtTmField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoCreateDtTmField.NumberFormat = $iDateFormatKey
		$iError = ($oDocInfoCreateDtTmField.NumberFormat() = $iDateFormatKey) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoCreateDtTmField.IsFixed() = False) Then $oDocInfoCreateDtTmField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoCreateDateTimeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoEditTimeInsert
; Description ...: Insert a Document Information Total Editing Time Field.
; Syntax ........: _LOWriter_FieldDocInfoEditTimeInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iTimeFormatKey = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $iTimeFormatKey      - [optional] an integer value. Default is Null. A Time Format Key returned from a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iTimeFormatKey not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iTimeFormatKey not found in document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.EditTime" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Total Editing Time Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: _LOWriter_FieldDocInfoEditTimeModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyList, _LOWriter_DocGenProp
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoEditTimeInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iTimeFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoEditTimeField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoEditTimeField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.EditTime")
	If Not IsObj($oDocInfoEditTimeField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoEditTimeField.IsFixed = $bIsFixed
	EndIf

	If ($iTimeFormatKey <> Null) Then
		If Not IsInt($iTimeFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iTimeFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oDocInfoEditTimeField.NumberFormat = $iTimeFormatKey
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoEditTimeField, $bOverwrite)

	If ($oDocInfoEditTimeField.IsFixed() = False) Then $oDocInfoEditTimeField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoEditTimeField)
EndFunc   ;==>_LOWriter_FieldDocInfoEditTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoEditTimeModify
; Description ...: Set or Retrieve a Document Information Total Editing Time Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoEditTimeModify(ByRef $oDoc, ByRef $oDocInfoEditTimeField[, $bIsFixed = Null[, $iTimeFormatKey = Null]])
; Parameters ....: $oDoc                  - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oDocInfoEditTimeField - [in/out] an object. A Doc Info Total Editing Time field Object from a previous _LOWriter_FieldDocInfoEditTimeInsert, or _LOWriter_FieldsDocInfoGetList function.
;                  $bIsFixed              - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $iTimeFormatKey        - [optional] an integer value. Default is Null. A Time Format Key returned from a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oDocInfoEditTimeField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iTimeFormatKey not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iTimeFormatKey not found in document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iTimeFormatKey
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoEditTimeInsert, _LOWriter_FieldsDocInfoGetList,
;					_LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyList, _LOWriter_DocGenProp
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoEditTimeModify(ByRef $oDoc, ByRef $oDocInfoEditTimeField, $bIsFixed = Null, $iTimeFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avDocInfoEditTm[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDocInfoEditTimeField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $iTimeFormatKey) Then
		; Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		; fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		; from the value.
		$iNumberFormat = $oDocInfoEditTimeField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : ($iNumberFormat)

		__LOWriter_ArrayFill($avDocInfoEditTm, $oDocInfoEditTimeField.IsFixed(), $iNumberFormat)
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoEditTm)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoEditTimeField.IsFixed = $bIsFixed
		$iError = ($oDocInfoEditTimeField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iTimeFormatKey <> Null) Then
		If Not IsInt($iTimeFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If Not IsInt($iTimeFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoEditTimeField.NumberFormat = $iTimeFormatKey
		$iError = ($oDocInfoEditTimeField.NumberFormat() = $iTimeFormatKey) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoEditTimeField.IsFixed() = False) Then $oDocInfoEditTimeField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoEditTimeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoKeywordsInsert
; Description ...: Insert a Document Information Keywords Field.
; Syntax ........: _LOWriter_FieldDocInfoKeywordsInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sKeywords = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sKeywords           - [optional] a string value. Default is Null. The Keywords text to display, note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sKeywords not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.Keywords" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Keywords Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoKeywordsModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoKeywordsInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sKeywords = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoKeywordField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoKeywordField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.KeyWords")
	If Not IsObj($oDocInfoKeywordField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoKeywordField.IsFixed = $bIsFixed
	EndIf

	If ($sKeywords <> Null) Then
		If Not IsString($sKeywords) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoKeywordField.Content = $sKeywords
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoKeywordField, $bOverwrite)

	If ($sKeywords <> Null) Then ; Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oDocInfoKeywordField.Content <> $sKeywords And ($oDocInfoKeywordField.IsFixed() = True) Then $oDocInfoKeywordField.Content = $sKeywords
	EndIf

	If ($oDocInfoKeywordField.IsFixed() = False) Then $oDocInfoKeywordField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoKeywordField)
EndFunc   ;==>_LOWriter_FieldDocInfoKeywordsInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoKeywordsModify
; Description ...: Set or Retrieve a Document Information Keywords Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoKeywordsModify(ByRef $oDocInfoKeywordField[, $bIsFixed = Null[, $sKeywords = Null]])
; Parameters ....: $oDocInfoKeywordField  - [in/out] an object. A Doc Info Keywords field Object from a previous _LOWriter_FieldDocInfoKeywordsInsert, or _LOWriter_FieldsDocInfoGetList function.
;                  $bIsFixed              - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sKeywords             - [optional] a string value. Default is Null. The Keywords text to display, note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoKeywordField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sKeywords not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sKeywords
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoKeywordsInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoKeywordsModify(ByRef $oDocInfoKeywordField, $bIsFixed = Null, $sKeywords = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoKyWrd[2]

	If Not IsObj($oDocInfoKeywordField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sKeywords) Then
		__LOWriter_ArrayFill($avDocInfoKyWrd, $oDocInfoKeywordField.IsFixed(), $oDocInfoKeywordField.Content())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoKyWrd)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoKeywordField.IsFixed = $bIsFixed
		$iError = ($oDocInfoKeywordField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sKeywords <> Null) Then
		If Not IsString($sKeywords) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoKeywordField.Content = $sKeywords
		$iError = ($oDocInfoKeywordField.Content() = $sKeywords) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoKeywordField.IsFixed() = False) Then $oDocInfoKeywordField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoKeywordsModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoModAuthInsert
; Description ...: Insert a Document Information Modification Author Field.
; Syntax ........: _LOWriter_FieldDocInfoModAuthInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sAuthor = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;				   $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed must be set to True in order for this to remain as set.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sAuthor not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.ChangeAuthor" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Modified By Author Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoModAuthModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocGenPropModification
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoModAuthInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoModAuthField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoModAuthField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.ChangeAuthor")
	If Not IsObj($oDocInfoModAuthField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoModAuthField.IsFixed = $bIsFixed
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoModAuthField.Author = $sAuthor
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoModAuthField, $bOverwrite)

	If ($sAuthor <> Null) Then ; Sometimes Author Disappears upon Insertion, make a check to re-set the Author value.
		If $oDocInfoModAuthField.Author <> $sAuthor And ($oDocInfoModAuthField.IsFixed() = True) Then $oDocInfoModAuthField.Author = $sAuthor
	EndIf

	If ($oDocInfoModAuthField.IsFixed() = False) Then $oDocInfoModAuthField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoModAuthField)
EndFunc   ;==>_LOWriter_FieldDocInfoModAuthInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoModAuthModify
; Description ...: Set or Retrieve a Document Information Modification Author Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoModAuthModify(ByRef $oDocInfoModAuthField[, $bIsFixed = Null[, $sAuthor = Null]])
; Parameters ....: $oDocInfoModAuthField  - [in/out] an object. A Modified By Author field Object from a previous _LOWriter_FieldDocInfoModAuthInsert, or _LOWriter_FieldsDocInfoGetList function.
;                  $bIsFixed              - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sAuthor               - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed must be set to True in order for this to remain as set.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoModAuthField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sAuthor not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sAuthor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoModAuthInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocGenPropModification
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoModAuthModify(ByRef $oDocInfoModAuthField, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoModAuth[2]

	If Not IsObj($oDocInfoModAuthField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sAuthor) Then
		__LOWriter_ArrayFill($avDocInfoModAuth, $oDocInfoModAuthField.IsFixed(), $oDocInfoModAuthField.Author())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoModAuth)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoModAuthField.IsFixed = $bIsFixed
		$iError = ($oDocInfoModAuthField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoModAuthField.Author = $sAuthor
		$iError = ($oDocInfoModAuthField.Author() = $sAuthor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoModAuthField.IsFixed() = False) Then $oDocInfoModAuthField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoModAuthModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoModDateTimeInsert
; Description ...: Insert a Document Information Modification Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoModDateTimeInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iDateFormatKey = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iDateFormatKey not found in document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.ChangeDateTime" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Modified Date/Time Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoModDateTimeModify, _LOWriter_DateFormatKeyCreate,
;					_LOWriter_DateFormatKeyList, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocGenPropModification
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoModDateTimeInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoModDtTmField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoModDtTmField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.ChangeDateTime")
	If Not IsObj($oDocInfoModDtTmField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoModDtTmField.IsFixed = $bIsFixed
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oDocInfoModDtTmField.NumberFormat = $iDateFormatKey
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoModDtTmField, $bOverwrite)

	If ($oDocInfoModDtTmField.IsFixed() = False) Then $oDocInfoModDtTmField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoModDtTmField)
EndFunc   ;==>_LOWriter_FieldDocInfoModDateTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoModDateTimeModify
; Description ...: Set or Retrieve a Document Information Modification Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoModDateTimeModify(ByRef $oDoc, ByRef $oDocInfoModDtTmField[, $bIsFixed = Null[, $iDateFormatKey = Null]])
; Parameters ....: $oDoc                  - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oDocInfoModDtTmField  - [in/out] an object. A Modified at Date/Time field Object from a previous _LOWriter_FieldDocInfoModDateTimeInsert,  or _LOWriter_FieldsDocInfoGetList function.
;                  $bIsFixed              - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey        - [optional] an integer value. Default is Null. A Date or Time Format Key returned from a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oDocInfoPrintAuthField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iDateFormatKey not found in document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iDateFormatKey
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoModDateTimeInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyList, _LOWriter_DocGenPropModification
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoModDateTimeModify(ByRef $oDoc, ByRef $oDocInfoModDtTmField, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avDocInfoModDate[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDocInfoModDtTmField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $iDateFormatKey) Then
		; Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		; fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		; from the value.
		$iNumberFormat = $oDocInfoModDtTmField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : ($iNumberFormat)

		__LOWriter_ArrayFill($avDocInfoModDate, $oDocInfoModDtTmField.IsFixed(), $iNumberFormat)
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoModDate)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoModDtTmField.IsFixed = $bIsFixed
		$iError = ($oDocInfoModDtTmField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoModDtTmField.NumberFormat = $iDateFormatKey
		$iError = ($oDocInfoModDtTmField.NumberFormat() = $iDateFormatKey) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoModDtTmField.IsFixed() = False) Then $oDocInfoModDtTmField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoModDateTimeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoPrintAuthInsert
; Description ...: Insert a Document Information Last Print Author Field.
; Syntax ........: _LOWriter_FieldDocInfoPrintAuthInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sAuthor = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sAuthor             - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed must be set to True in order for this to remain as set.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sAuthor not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.PrintAuthor" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Printed By Author Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoPrintAuthModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocGenPropPrint
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoPrintAuthInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoPrintAuthField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoPrintAuthField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.PrintAuthor")
	If Not IsObj($oDocInfoPrintAuthField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoPrintAuthField.IsFixed = $bIsFixed
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoPrintAuthField.Author = $sAuthor
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoPrintAuthField, $bOverwrite)

	If ($sAuthor <> Null) Then ; Sometimes Author Disappears upon Insertion, make a check to re-set the Author value.
		If $oDocInfoPrintAuthField.Author <> $sAuthor And ($oDocInfoPrintAuthField.IsFixed() = True) Then $oDocInfoPrintAuthField.Author = $sAuthor
	EndIf

	If ($oDocInfoPrintAuthField.IsFixed() = False) Then $oDocInfoPrintAuthField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoPrintAuthField)
EndFunc   ;==>_LOWriter_FieldDocInfoPrintAuthInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoPrintAuthModify
; Description ...: Set or Retrieve a Document Information Last Print Author Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoPrintAuthModify(ByRef $oDocInfoPrintAuthField[, $bIsFixed = Null[, $sAuthor = Null]])
; Parameters ....: $oDocInfoPrintAuthField  - [in/out] an object. A Printed By Author field Object from a previous _LOWriter_FieldDocInfoPrintAuthInsert, or _LOWriter_FieldsDocInfoGetList function.
;                  $bIsFixed                - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sAuthor                 - [optional] a string value. Default is Null. The Author's name, note, $bIsFixed must be set to True in order for this to remain as set.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoPrintAuthField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sAuthor not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sAuthor
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoPrintAuthInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocGenPropPrint
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoPrintAuthModify(ByRef $oDocInfoPrintAuthField, $bIsFixed = Null, $sAuthor = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoModAuth[2]

	If Not IsObj($oDocInfoPrintAuthField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sAuthor) Then
		__LOWriter_ArrayFill($avDocInfoModAuth, $oDocInfoPrintAuthField.IsFixed(), $oDocInfoPrintAuthField.Author())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoModAuth)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoPrintAuthField.IsFixed = $bIsFixed
		$iError = ($oDocInfoPrintAuthField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sAuthor <> Null) Then
		If Not IsString($sAuthor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoPrintAuthField.Author = $sAuthor
		$iError = ($oDocInfoPrintAuthField.Author() = $sAuthor) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoPrintAuthField.IsFixed() = False) Then $oDocInfoPrintAuthField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoPrintAuthModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoPrintDateTimeInsert
; Description ...: Insert a Document Information Last Print Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoPrintDateTimeInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iDateFormatKey = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey      - [optional] an integer value. Default is Null. A Date or Time Format Key returned from a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iDateFormatKey not found in document.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.PrintDateTime" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Printed Date/Time Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoPrintDateTimeModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DateFormatKeyCreate, _LOWriter_DateFormatKeyList, _LOWriter_DocGenPropPrint
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoPrintDateTimeInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoPrintDtTmField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoPrintDtTmField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.PrintDateTime")
	If Not IsObj($oDocInfoPrintDtTmField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoPrintDtTmField.IsFixed = $bIsFixed
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oDocInfoPrintDtTmField.NumberFormat = $iDateFormatKey
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoPrintDtTmField, $bOverwrite)

	If ($oDocInfoPrintDtTmField.IsFixed() = False) Then $oDocInfoPrintDtTmField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoPrintDtTmField)
EndFunc   ;==>_LOWriter_FieldDocInfoPrintDateTimeInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoPrintDateTimeModify
; Description ...: Set or Retrieve a Document Information Last Print Date/Time Field.
; Syntax ........: _LOWriter_FieldDocInfoPrintDateTimeModify(ByRef $oDoc, ByRef $oDocInfoPrintDtTmField[, $bIsFixed = Null[, $iDateFormatKey = Null]])
; Parameters ....: $oDoc                    - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oDocInfoPrintDtTmField  - [in/out] an object. A Printed at Date/Time field Object from a previous _LOWriter_FieldDocInfoPrintDateTimeInsert, or _LOWriter_FieldsDocInfoGetList function.
;                  $bIsFixed                - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $iDateFormatKey          - [optional] an integer value. Default is Null. A Date or Time Format Key returned from a previous _LOWriter_DateFormatKeyCreate or _LOWriter_DateFormatKeyList function.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oDocInfoPrintDtTmField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $iDateFormatKey not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iDateFormatKey not found in document.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iDateFormatKey
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoPrintDateTimeInsert,  _LOWriter_FieldsDocInfoGetList, _LOWriter_DateFormatKeyCreate,
;					_LOWriter_DateFormatKeyList, _LOWriter_DocGenPropPrint
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoPrintDateTimeModify(ByRef $oDoc, ByRef $oDocInfoPrintDtTmField, $bIsFixed = Null, $iDateFormatKey = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avDocInfoPrntDate[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oDocInfoPrintDtTmField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $iDateFormatKey) Then
		; Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		; fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		; from the value.
		$iNumberFormat = $oDocInfoPrintDtTmField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : ($iNumberFormat)

		__LOWriter_ArrayFill($avDocInfoPrntDate, $oDocInfoPrintDtTmField.IsFixed(), $iNumberFormat)
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoPrntDate)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoPrintDtTmField.IsFixed = $bIsFixed
		$iError = ($oDocInfoPrintDtTmField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iDateFormatKey <> Null) Then
		If Not IsInt($iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If Not _LOWriter_DateFormatKeyExists($oDoc, $iDateFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoPrintDtTmField.NumberFormat = $iDateFormatKey
		$iError = ($oDocInfoPrintDtTmField.NumberFormat() = $iDateFormatKey) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoPrintDtTmField.IsFixed() = False) Then $oDocInfoPrintDtTmField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoPrintDateTimeModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoRevNumInsert
; Description ...: Insert a Document Information Revision Number Field.
; Syntax ........: _LOWriter_FieldDocInfoRevNumInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iRevNum = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $iRevNum             - [optional] a Integer value. Default is Null. The Revision Number Integer to display, note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iRevNum not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.Revision" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Revision Number Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoRevNumModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocGenProp
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoRevNumInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iRevNum = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoRevNumField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoRevNumField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.Revision")
	If Not IsObj($oDocInfoRevNumField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoRevNumField.IsFixed = $bIsFixed
	EndIf

	If ($iRevNum <> Null) Then
		If Not IsInt($iRevNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoRevNumField.Revision = $iRevNum
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoRevNumField, $bOverwrite)

	If ($iRevNum <> Null) Then ; Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oDocInfoRevNumField.Revision <> $iRevNum And ($oDocInfoRevNumField.IsFixed() = True) Then $oDocInfoRevNumField.Revision = $iRevNum
	EndIf

	If ($oDocInfoRevNumField.IsFixed() = False) Then $oDocInfoRevNumField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoRevNumField)
EndFunc   ;==>_LOWriter_FieldDocInfoRevNumInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoRevNumModify
; Description ...: Set or Retrieve a Document Information Revision Number Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoRevNumModify(ByRef $oDocInfoRevNumField[, $bIsFixed = Null[, $iRevNum = Null]])
; Parameters ....: $oDocInfoRevNumField - [in/out] an object. A Doc Info Revision Number field Object from a previous _LOWriter_FieldDocInfoRevNumInsert, or _LOWriter_FieldsDocInfoGetList function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $iRevNum             - [optional] a Integer value. Default is Null. The Revision Number Integer to display, note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoRevNumField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $iRevNum not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iRevNum
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoRevNumInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocGenProp
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoRevNumModify(ByRef $oDocInfoRevNumField, $bIsFixed = Null, $iRevNum = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoRev[2]

	If Not IsObj($oDocInfoRevNumField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $iRevNum) Then
		__LOWriter_ArrayFill($avDocInfoRev, $oDocInfoRevNumField.IsFixed(), $oDocInfoRevNumField.Revision())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoRev)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoRevNumField.IsFixed = $bIsFixed
		$iError = ($oDocInfoRevNumField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRevNum <> Null) Then
		If Not IsInt($iRevNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoRevNumField.Revision = $iRevNum
		$iError = ($oDocInfoRevNumField.Revision() = $iRevNum) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoRevNumField.IsFixed() = False) Then $oDocInfoRevNumField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoRevNumModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoSubjectInsert
; Description ...: Insert a Document Information Subject Field.
; Syntax ........: _LOWriter_FieldDocInfoSubjectInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sSubject = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sSubject            - [optional] a string value. Default is Null. The Subject text to display, note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sSubject not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.Subject" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Subject Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoSubjectModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoSubjectInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sSubject = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoSubField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoSubField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.Subject")
	If Not IsObj($oDocInfoSubField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoSubField.IsFixed = $bIsFixed
	EndIf

	If ($sSubject <> Null) Then
		If Not IsString($sSubject) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoSubField.Content = $sSubject
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoSubField, $bOverwrite)

	If ($sSubject <> Null) Then ; Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oDocInfoSubField.Content <> $sSubject And ($oDocInfoSubField.IsFixed() = True) Then $oDocInfoSubField.Content = $sSubject
	EndIf

	If ($oDocInfoSubField.IsFixed() = False) Then $oDocInfoSubField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoSubField)
EndFunc   ;==>_LOWriter_FieldDocInfoSubjectInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoSubjectModify
; Description ...: Set or Retrieve a Document Information Subject Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoSubjectModify(ByRef $oDocInfoSubField[, $bIsFixed = Null[, $sSubject = Null]])
; Parameters ....: $oDocInfoSubField    - [in/out] an object. A Doc Info Subject field Object from a previous _LOWriter_FieldDocInfoSubjectInsert, or _LOWriter_FieldsDocInfoGetList function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sSubject            - [optional] a string value. Default is Null. The Subject text to display, note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoSubField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sSubject not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sSubject
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoSubjectInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoSubjectModify(ByRef $oDocInfoSubField, $bIsFixed = Null, $sSubject = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoSub[2]

	If Not IsObj($oDocInfoSubField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sSubject) Then
		__LOWriter_ArrayFill($avDocInfoSub, $oDocInfoSubField.IsFixed(), $oDocInfoSubField.Content())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoSub)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoSubField.IsFixed = $bIsFixed
		$iError = ($oDocInfoSubField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sSubject <> Null) Then
		If Not IsString($sSubject) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoSubField.Content = $sSubject
		$iError = ($oDocInfoSubField.Content() = $sSubject) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoSubField.IsFixed() = False) Then $oDocInfoSubField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoSubjectModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoTitleInsert
; Description ...: Insert a Document Information Title Field.
; Syntax ........: _LOWriter_FieldDocInfoTitleInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sTitle = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sTitle              - [optional] a string value. Default is Null. The Title text to display, note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sTitle not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.textfield.docinfo.Title" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Document Info Title Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldDocInfoTitleModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoTitleInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sTitle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oDocInfoTitleField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oDocInfoTitleField = $oDoc.createInstance("com.sun.star.text.textfield.docinfo.Title")
	If Not IsObj($oDocInfoTitleField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oDocInfoTitleField.IsFixed = $bIsFixed
	EndIf

	If ($sTitle <> Null) Then
		If Not IsString($sTitle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oDocInfoTitleField.Content = $sTitle
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oDocInfoTitleField, $bOverwrite)

	If ($sTitle <> Null) Then ; Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oDocInfoTitleField.Content <> $sTitle And ($oDocInfoTitleField.IsFixed() = True) Then $oDocInfoTitleField.Content = $sTitle
	EndIf

	If ($oDocInfoTitleField.IsFixed() = False) Then $oDocInfoTitleField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oDocInfoTitleField)
EndFunc   ;==>_LOWriter_FieldDocInfoTitleInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldDocInfoTitleModify
; Description ...: Set or Retrieve a Document Information Title Field's settings.
; Syntax ........: _LOWriter_FieldDocInfoTitleModify(ByRef $oDocInfoTitleField[, $bIsFixed = Null[, $sTitle = Null]])
; Parameters ....: $oDocInfoTitleField  - [in/out] an object. A Doc Info Title field Object from a previous _LOWriter_FieldDocInfoTitleInsert, or _LOWriter_FieldsDocInfoGetList function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sTitle              - [optional] a string value. Default is Null. The Title text to display, note, $bIsFixed must be True for this to be displayed.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDocInfoTitleField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sTitle not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sTitle
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldDocInfoTitleInsert, _LOWriter_FieldsDocInfoGetList, _LOWriter_DocDescription
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldDocInfoTitleModify(ByRef $oDocInfoTitleField, $bIsFixed = Null, $sTitle = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDocInfoTitle[2]

	If Not IsObj($oDocInfoTitleField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sTitle) Then
		__LOWriter_ArrayFill($avDocInfoTitle, $oDocInfoTitleField.IsFixed(), $oDocInfoTitleField.Content())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDocInfoTitle)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oDocInfoTitleField.IsFixed = $bIsFixed
		$iError = ($oDocInfoTitleField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sTitle <> Null) Then
		If Not IsString($sTitle) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oDocInfoTitleField.Content = $sTitle
		$iError = ($oDocInfoTitleField.Content() = $sTitle) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($oDocInfoTitleField.IsFixed() = False) Then $oDocInfoTitleField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldDocInfoTitleModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFileNameInsert
; Description ...: Insert a File Name Field.
; Syntax ........: _LOWriter_FieldFileNameInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $iFormat = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;				   $iFormat             - [optional] an integer value (0-3). Default is Null. The Data Format to display. See Constants, $LOW_FIELD_FILENAME_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iFormat not an Integer, less than 0, or greater than 3. See Constants, $LOW_FIELD_FILENAME_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.FileName" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted File Name field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Until at least L.O. Version 7.3.4.2, there is a bug where the wrong Path Format type is displayed when the content is set to Fixed = True.
;				   For example, $LOW_FIELD_FILENAME_NAME_AND_EXT, displays in the format of $LOW_FIELD_FILENAME_NAME.
; Related .......: _LOWriter_FieldFileNameModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFileNameInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFileNameField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oFileNameField = $oDoc.createInstance("com.sun.star.text.TextField.FileName")
	If Not IsObj($oFileNameField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oFileNameField.IsFixed = $bIsFixed
	EndIf

	If ($iFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iFormat, $LOW_FIELD_FILENAME_FULL_PATH, $LOW_FIELD_FILENAME_NAME_AND_EXT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oFileNameField.FileFormat = $iFormat
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oFileNameField, $bOverwrite)

	If ($oFileNameField.IsFixed() = False) Then $oFileNameField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFileNameField)
EndFunc   ;==>_LOWriter_FieldFileNameInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFileNameModify
; Description ...: Set or Retrieve a File Name Field's settings.
; Syntax ........: _LOWriter_FieldFileNameModify(ByRef $oFileNameField[, $bIsFixed = Null[, $iFormat = Null]])
; Parameters ....: $oFileNameField      - [in/out] an object. A File Name field Object from a previous _LOWriter_FieldFileNameInsert, or _LOWriter_FieldsGetList function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;				   $iFormat             - [optional] an integer value (0-3). Default is Null. The Data Format to display. See Constants, $LOW_FIELD_FILENAME_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oFileNameField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $iFormat not an Integer, less than 0, or greater than 3. See Constants, $LOW_FIELD_FILENAME_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $iFormat
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Until at least L.O. Version 7.3.4.2, there is a bug where the wrong Path Format type is displayed when the
;						content is set to Fixed = True. For example, $LOW_FIELD_FILENAME_NAME_AND_EXT, displays in the format
;							of $LOW_FIELD_FILENAME_NAME.
;				   Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldFileNameInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFileNameModify(ByRef $oFileNameField, $bIsFixed = Null, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avFileName[2]

	If Not IsObj($oFileNameField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iFormat, $bIsFixed) Then
		__LOWriter_ArrayFill($avFileName, $oFileNameField.IsFixed(), $oFileNameField.FileFormat())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avFileName)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oFileNameField.IsFixed = $bIsFixed
		$iError = ($oFileNameField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iFormat, $LOW_FIELD_FILENAME_FULL_PATH, $LOW_FIELD_FILENAME_NAME_AND_EXT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oFileNameField.FileFormat = $iFormat
		$iError = ($oFileNameField.FileFormat() = $iFormat) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($oFileNameField.IsFixed() = False) Then $oFileNameField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldFileNameModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncHiddenParInsert
; Description ...: Insert a Hidden Paragraph Field.
; Syntax ........: _LOWriter_FieldFuncHiddenParInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $sCondition = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $sCondition          - [optional] a string value. Default is Null. The condition to evaluate.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sCondition not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.HiddenParagraph" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Hidden Paragraph Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldFuncHiddenParModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncHiddenParInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sCondition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oHidParField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oHidParField = $oDoc.createInstance("com.sun.star.text.TextField.HiddenParagraph")
	If Not IsObj($oHidParField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oHidParField.Condition = $sCondition
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oHidParField, $bOverwrite)

	$oHidParField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oHidParField)
EndFunc   ;==>_LOWriter_FieldFuncHiddenParInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncHiddenParModify
; Description ...: Set or Retrieve a Hidden Paragraph Field's settings.
; Syntax ........: _LOWriter_FieldFuncHiddenParModify(ByRef $oHidParField[, $sCondition = Null])
; Parameters ....: $oHidParField        - [in/out] an object. A Hidden Paragraph field Object from a previous _LOWriter_FieldFuncHiddenParInsert, or _LOWriter_FieldsGetList function.
;                  $sCondition          - [optional] a string value. Default is Null. The condition to evaluate.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oHidParField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sCondition not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sCondition
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters. The second Element is a boolean whether the Paragraph is Hidden(True) or Visible(False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldFuncHiddenParInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncHiddenParModify(ByRef $oHidParField, $sCondition = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHidPar[2]

	If Not IsObj($oHidParField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sCondition) Then
		__LOWriter_ArrayFill($avHidPar, $oHidParField.Condition(), ($oHidParField.IsHidden()) ? (False) : (True)) ; "IsHidden" Is Backwards
		Return SetError($__LO_STATUS_SUCCESS, 1, $avHidPar)
	EndIf

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oHidParField.Condition = $sCondition
		$iError = ($oHidParField.Condition() = $sCondition) ? ($iError) : (BitOR($iError, 1))
	EndIf

	$oHidParField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldFuncHiddenParModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncHiddenTextInsert
; Description ...: Insert a Hidden Text Field.
; Syntax ........: _LOWriter_FieldFuncHiddenTextInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $sCondition = Null[, $sText = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $sCondition          - [optional] a string value. Default is Null. The Condition to evaluate.
;                  $sText               - [optional] a string value. Default is Null. The Text to show if the condition evaluates as true.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sCondition not a String.
;				   @Error 1 @Extended 6 Return 0 = $sText not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.HiddenText" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Hidden Text Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldFuncHiddenTextModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncHiddenTextInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sCondition = Null, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oHidTxtField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oHidTxtField = $oDoc.createInstance("com.sun.star.text.TextField.HiddenText")
	If Not IsObj($oHidTxtField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oHidTxtField.Condition = $sCondition
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oHidTxtField.Content = $sText
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oHidTxtField, $bOverwrite)

	$oHidTxtField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oHidTxtField)
EndFunc   ;==>_LOWriter_FieldFuncHiddenTextInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncHiddenTextModify
; Description ...: Set or Retrieve a Hidden Text Field's settings.
; Syntax ........: _LOWriter_FieldFuncHiddenTextModify(ByRef $oHidTxtField[, $sCondition = Null[, $sText = Null]])
; Parameters ....: $oHidTxtField        - [in/out] an object. A Hidden Text field Object from a previous _LOWriter_FieldFuncHiddenTextInsert, or _LOWriter_FieldsGetList function.
;                  $sCondition          - [optional] a string value. Default is Null. The Condition to evaluate.
;                  $sText               - [optional] a string value. Default is Null. The Text to show if the condition evaluates as true.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oHidTxtField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sCondition not a String.
;				   @Error 1 @Extended 3 Return 0 = $sText not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sCondition
;				   |								2 = Error setting $sText
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters. The Third Element is a boolean whether the Text is Hidden(True) Or Visible(False).
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldFuncHiddenTextInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncHiddenTextModify(ByRef $oHidTxtField, $sCondition = Null, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avHidPar[3]

	If Not IsObj($oHidTxtField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sCondition, $sText) Then
		__LOWriter_ArrayFill($avHidPar, $oHidTxtField.Condition(), $oHidTxtField.Content(), ($oHidTxtField.IsHidden()) ? (False) : (True)) ; "IsHidden" Is Backwards
		Return SetError($__LO_STATUS_SUCCESS, 1, $avHidPar)
	EndIf

	If ($sCondition <> Null) Then
		If Not IsString($sCondition) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oHidTxtField.Condition = $sCondition
		$iError = ($oHidTxtField.Condition() = $sCondition) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oHidTxtField.Content = $sText
		$iError = ($oHidTxtField.Content() = $sText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	$oHidTxtField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldFuncHiddenTextModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncInputInsert
; Description ...: Insert a Input Field.
; Syntax ........: _LOWriter_FieldFuncInputInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $sReference = Null[, $sText = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $sReference          - [optional] a string value. Default is Null. The Reference to display for the input field.
;                  $sText               - [optional] a string value. Default is Null. The Text to insert in the Input Field.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sReference not a String.
;				   @Error 1 @Extended 6 Return 0 = $sText not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.Input" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Input Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldFuncInputModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncInputInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $sReference = Null, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oInputField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oInputField = $oDoc.createInstance("com.sun.star.text.TextField.Input")
	If Not IsObj($oInputField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($sReference <> Null) Then
		If Not IsString($sReference) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oInputField.Hint = $sReference
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oInputField.Content = $sText
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oInputField, $bOverwrite)

	$oInputField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oInputField)
EndFunc   ;==>_LOWriter_FieldFuncInputInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncInputModify
; Description ...: Set or Retrieve a Input Field's settings.
; Syntax ........: _LOWriter_FieldFuncInputModify(ByRef $oInputField[, $sReference = Null[, $sText = Null]])
; Parameters ....: $oInputField         - [in/out] an object. A Input field Object from a previous _LOWriter_FieldFuncInputInsert, or _LOWriter_FieldsGetList function.
;                  $sReference          - [optional] a string value. Default is Null. The Reference to display for the input field.
;                  $sText               - [optional] a string value. Default is Null. The Text to insert in the Input Field.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oHidTxtField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sReference not a String.
;				   @Error 1 @Extended 3 Return 0 = $sText not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sReference
;				   |								2 = Error setting $sText
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldFuncInputInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncInputModify(ByRef $oInputField, $sReference = Null, $sText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asInput[2]

	If Not IsObj($oInputField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($sReference, $sText) Then
		__LOWriter_ArrayFill($asInput, $oInputField.Hint(), $oInputField.Content())
		Return SetError($__LO_STATUS_SUCCESS, 1, $asInput)
	EndIf

	If ($sReference <> Null) Then
		If Not IsString($sReference) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oInputField.Hint = $sReference
		$iError = ($oInputField.Hint() = $sReference) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sText <> Null) Then
		If Not IsString($sText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oInputField.Content = $sText
		$iError = ($oInputField.Content() = $sText) ? ($iError) : (BitOR($iError, 2))
	EndIf

	$oInputField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldFuncInputModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncPlaceholderInsert
; Description ...: Insert a Placeholder Field.
; Syntax ........: _LOWriter_FieldFuncPlaceholderInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $iPHolderType = Null[, $sPHolderName = Null[, $sReference = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iPHolderType        - [optional] an integer value (0-4). Default is Null. The type of Placeholder to insert. See Constants, $LOW_FIELD_PLACEHOLD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sPHolderName        - [optional] a string value. Default is Null. The Placeholder's name.
;                  $sReference          - [optional] a string value. Default is Null. A Reference to display when the mouse hovers the Placeholder.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iPHolderType not an Integer, less than 0, or greater than 4. See Constants, $LOW_FIELD_PLACEHOLD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $sPHolderName not a String.
;				   @Error 1 @Extended 7 Return 0 = $sReference not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.JumpEdit" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Placeholder Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldFuncPlaceholderModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncPlaceholderInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $iPHolderType = Null, $sPHolderName = Null, $sReference = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPHolderField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oPHolderField = $oDoc.createInstance("com.sun.star.text.TextField.JumpEdit")
	If Not IsObj($oPHolderField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($iPHolderType <> Null) Then
		If Not __LOWriter_IntIsBetween($iPHolderType, $LOW_FIELD_PLACEHOLD_TYPE_TEXT, $LOW_FIELD_PLACEHOLD_TYPE_OBJECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oPHolderField.PlaceHolderType = $iPHolderType
	EndIf

	If ($sPHolderName <> Null) Then
		If Not IsString($sPHolderName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oPHolderField.PlaceHolder = $sPHolderName
	EndIf

	If ($sReference <> Null) Then
		If Not IsString($sReference) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oPHolderField.Hint = $sReference
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oPHolderField, $bOverwrite)

	$oPHolderField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPHolderField)
EndFunc   ;==>_LOWriter_FieldFuncPlaceholderInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldFuncPlaceholderModify
; Description ...: Set or Retrieve a Placeholder Field's settings.
; Syntax ........: _LOWriter_FieldFuncPlaceholderModify(ByRef $oPHolderField[, $iPHolderType = Null[, $sPHolderName = Null[, $sReference = Null]]])
; Parameters ....: $oPHolderField       - [in/out] an object. A Placeholder field Object from a previous _LOWriter_FieldFuncPlaceholderInsert, or _LOWriter_FieldsGetList function.
;                  $iPHolderType        - [optional] an integer value (0-4). Default is Null. The type of Placeholder to insert. See Constants, $LOW_FIELD_PLACEHOLD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sPHolderName        - [optional] a string value. Default is Null. The Placeholder's name.
;                  $sReference          - [optional] a string value. Default is Null. A Reference to display when the mouse hovers the Placeholder.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oPHolderField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iPHolderType not an Integer, less than 0, or greater than 4. See Constants, $LOW_FIELD_PLACEHOLD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $sPHolderName not a String.
;				   @Error 1 @Extended 4 Return 0 = $sReference not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iPHolderType
;				   |								2 = Error setting $sPHolderName
;				   |								4 = Error setting $sReference
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldFuncPlaceholderInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldFuncPlaceholderModify(ByRef $oPHolderField, $iPHolderType = Null, $sPHolderName = Null, $sReference = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $asPHolder[3]

	If Not IsObj($oPHolderField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iPHolderType, $sPHolderName, $sReference) Then
		__LOWriter_ArrayFill($asPHolder, $oPHolderField.PlaceHolderType(), $oPHolderField.PlaceHolder(), $oPHolderField.Hint())
		Return SetError($__LO_STATUS_SUCCESS, 1, $asPHolder)
	EndIf

	If ($iPHolderType <> Null) Then
		If Not __LOWriter_IntIsBetween($iPHolderType, $LOW_FIELD_PLACEHOLD_TYPE_TEXT, $LOW_FIELD_PLACEHOLD_TYPE_OBJECT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oPHolderField.PlaceHolderType = $iPHolderType
		$iError = ($oPHolderField.PlaceHolderType() = $iPHolderType) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sPHolderName <> Null) Then
		If Not IsString($sPHolderName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oPHolderField.PlaceHolder = $sPHolderName
		$iError = ($oPHolderField.PlaceHolder() = $sPHolderName) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sReference <> Null) Then
		If Not IsString($sReference) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oPHolderField.Hint = $sReference
		$iError = ($oPHolderField.Hint() = $sReference) ? ($iError) : (BitOR($iError, 4))
	EndIf

	$oPHolderField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldFuncPlaceholderModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldGetAnchor
; Description ...: Retrieve the Anchor Cursor Object for a Field inserted in a document.
; Syntax ........: _LOWriter_FieldGetAnchor(ByRef $oField)
; Parameters ....: $oField              - [in/out] an object. A Field Object returned from a previous Insert, _LOWriter_FieldsGetList, _LOWriter_FieldsAdvGetList, or _LOWriter_FieldsDocInfoGetList function.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oField not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Field anchor Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning requested Field Anchor Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldsGetList, _LOWriter_FieldsAdvGetList, _LOWriter_FieldsDocInfoGetList, _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldGetAnchor(ByRef $oField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFieldAnchor

	If Not IsObj($oField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oFieldAnchor = $oField.Anchor.Text.createTextCursorByRange($oField.Anchor())
	If Not IsObj($oFieldAnchor) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFieldAnchor)
EndFunc   ;==>_LOWriter_FieldGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldInputListInsert
; Description ...: Insert a Input List Field.
; Syntax ........: _LOWriter_FieldInputListInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $asItems = Null[, $sName = Null[, $sSelectedItem = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $asItems             - [optional] an array of strings. Default is Null. A single column Array of Strings to colonize the List with.
;                  $sName               - [optional] a string value. Default is Null. The name of the Input List Field.
;                  $sSelectedItem       - [optional] a string value. Default is Null. The Item in the list to be currently selected. Defaults to "" if Item is not found.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $asItems not an Array.
;				   @Error 1 @Extended 6 Return 0 = $sName not a String.
;				   @Error 1 @Extended 7 Return 0 = $sSelectedItem not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.DropDown" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Input List field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldInputListModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldInputListInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $asItems = Null, $sName = Null, $sSelectedItem = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oInputField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oInputField = $oDoc.createInstance("com.sun.star.text.TextField.DropDown")
	If Not IsObj($oInputField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($asItems <> Null) Then
		If Not IsArray($asItems) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oInputField.Items = $asItems
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oInputField.Name = $sName
	EndIf

	If ($sSelectedItem <> Null) Then
		If Not IsString($sSelectedItem) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oInputField.SelectedItem = $sSelectedItem
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oInputField, $bOverwrite)

	$oInputField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oInputField)
EndFunc   ;==>_LOWriter_FieldInputListInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldInputListModify
; Description ...: Set or Retrieve a Input List Field's settings.
; Syntax ........: _LOWriter_FieldInputListModify(ByRef $oInputField[, $asItems = Null[, $sName = Null[, $sSelectedItem = Null]]])
; Parameters ....: $oInputField         - [in/out] an object. A Input List field Object from a previous _LOWriter_FieldInputListInsert, or _LOWriter_FieldsGetList function.
;                  $asItems             - [optional] an array of strings. Default is Null. A single column Array of Strings to colonize the List with.
;                  $sName               - [optional] a string value. Default is Null. The name of the Input List Field.
;                  $sSelectedItem       - [optional] a string value. Default is Null. The Item in the list to be currently selected. Defaults to "" if Item is not found.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oInputField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $asItems not an Array.
;				   @Error 1 @Extended 3 Return 0 = $sName not a String.
;				   @Error 1 @Extended 4 Return 0 = $sSelectedItem not a String.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $asItems
;				   |								2 = Error setting $sName
;				   |								4 = Error setting $sSelectedItem
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldInputListInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldInputListModify(ByRef $oInputField, $asItems = Null, $sName = Null, $sSelectedItem = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avDropDwn[3]

	If Not IsObj($oInputField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($asItems, $sName, $sSelectedItem) Then
		__LOWriter_ArrayFill($avDropDwn, $oInputField.Items(), $oInputField.Name(), $oInputField.SelectedItem())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avDropDwn)
	EndIf

	If ($asItems <> Null) Then
		If Not IsArray($asItems) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oInputField.Items = $asItems
		$iError = (UBound($oInputField.Items()) = UBound($asItems)) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sName <> Null) Then
		If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oInputField.Name = $sName
		$iError = ($oInputField.Name() = $sName) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($sSelectedItem <> Null) Then
		If Not IsString($sSelectedItem) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oInputField.SelectedItem = $sSelectedItem
		$iError = ($oInputField.SelectedItem() = $sSelectedItem) ? ($iError) : (BitOR($iError, 4))
	EndIf

	$oInputField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldInputListModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldPageNumberInsert
; Description ...: Insert a Page number field.
; Syntax ........: _LOWriter_FieldPageNumberInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $iNumFormat = Null[, $iOffset = Null[, $iPageNumType = Null[, $sUserText = Null]]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iNumFormat          - [optional] an integer value (0-71). Default is Null. The numbering format to use for Page numbering. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iOffset             - [optional] an integer value. Default is Null. The number of pages to minus or add to the page Number.
;                  $iPageNumType        - [optional] an integer value (0-2). Default is Null. The Page Number type, either previous, current or next page. See Constants, $LOW_PAGE_NUM_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sUserText           - [optional] a string value. Default is Null. The custom user text to display. Only valid if $iNumFormat is set to $LOW_NUM_STYLE_CHAR_SPECIAL(6).
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iNumFormat not an Integer, less than 0, or greater than 71. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $iOffset not an Integer.
;				   @Error 1 @Extended 7 Return 0 = $iPageNumType not an Integer, less than 0, or greater than 2. See Constants, $LOW_PAGE_NUM_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 8 Return 0 = $sUserText not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.PageNumber" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Page Number field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldPageNumberModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldPageNumberInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $iNumFormat = Null, $iOffset = Null, $iPageNumType = Null, $sUserText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPageField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oPageField = $oDoc.createInstance("com.sun.star.text.TextField.PageNumber")
	If Not IsObj($oPageField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oPageField.NumberingType = $iNumFormat
	Else
		$oPageField.NumberingType = $LOW_NUM_STYLE_PAGE_DESCRIPTOR
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oPageField.Offset = $iOffset
	EndIf

	If ($iPageNumType <> Null) Then
		If Not __LOWriter_IntIsBetween($iPageNumType, $LOW_PAGE_NUM_TYPE_PREV, $LOW_PAGE_NUM_TYPE_NEXT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oPageField.SubType = $iPageNumType

		If ($iPageNumType = $LOW_PAGE_NUM_TYPE_PREV) Then
			$oPageField.Offset = ($oPageField.Offset() - 1) ; If SubType is Set to Prev. Set offset to minus 1 of current value
		ElseIf ($iPageNumType = $LOW_PAGE_NUM_TYPE_NEXT) Then
			$oPageField.Offset = ($oPageField.Offset() + 1) ; If SubType is Set to Next. Set offset to plus 1 of current value
		EndIf
	Else
		$oPageField.SubType = $LOW_PAGE_NUM_TYPE_CURRENT ;If not set, page number Sub Type is auto set to Prev. Instead of current.
	EndIf

	If ($sUserText <> Null) Then
		If Not IsString($sUserText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$oPageField.UserText = $sUserText
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oPageField, $bOverwrite)

	$oPageField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPageField)
EndFunc   ;==>_LOWriter_FieldPageNumberInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldPageNumberModify
; Description ...: Set or Retrieve Page Number Field settings.
; Syntax ........: _LOWriter_FieldPageNumberModify(ByRef $oDoc, ByRef $oPageNumField[, $iNumFormat = Null[, $iOffset = Null[, $iPageNumType = Null[, $sUserText = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oPageNumField       - [in/out] an object. A Page Number field Object from a previous _LOWriter_FieldPageNumberInsert, or _LOWriter_FieldsGetList function.
;                  $iNumFormat          - [optional] an integer value (0-71). Default is Null. The numbering format to use for Page numbering. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iOffset             - [optional] an integer value. Default is Null. The number of pages to minus or add to the page Number.
;                  $iPageNumType        - [optional] an integer value (0-2). Default is Null. The Page Number type, either previous, current or next page. See Constants, $LOW_PAGE_NUM_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $sUserText           - [optional] a string value. Default is Null. The custom User text to display. Only valid if $iNumFormat is set to $LOW_NUM_STYLE_CHAR_SPECIAL(6).
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oPageNumField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iNumFormat not an Integer, less than 0, or greater than 71. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $iOffset not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iPageNumType not an Integer, less than 0, or greater than 2. See Constants, $LOW_PAGE_NUM_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 6 Return 0 = $sUserText not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.PageNumber" Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iNumFormat
;				   |								2 = Error setting $iOffset
;				   |								4 = Error setting $iPageNumType
;				   |								8 = Error setting $sUserText
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldPageNumberInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldPageNumberModify(ByRef $oDoc, ByRef $oPageNumField, $iNumFormat = Null, $iOffset = Null, $iPageNumType = Null, $sUserText = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $avField[4]
	Local $iError = 0
	Local $oNewPageNumField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oPageNumField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($iNumFormat, $iOffset, $iPageNumType, $sUserText) Then
		__LOWriter_ArrayFill($avField, $oPageNumField.NumberingType(), $oPageNumField.Offset(), $oPageNumField.SubType(), $oPageNumField.UserText())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avField)
	EndIf

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

		$oNewPageNumField = $oDoc.createInstance("com.sun.star.text.TextField.PageNumber")
		If Not IsObj($oNewPageNumField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

		; It doesn't work to just set a new Numbering type for an already inserted Page Number, so I have to create a new one and
		; then insert it.
		With $oNewPageNumField
			.NumberingType = $iNumFormat
			.Offset = $oPageNumField.Offset()
			.SubType = $oPageNumField.SubType()
			.UserText = $oPageNumField.UserText()
		EndWith

		$oDoc.Text.createTextCursorByRange($oPageNumField.Anchor()).Text.insertTextContent($oPageNumField.Anchor(), $oNewPageNumField, True)

		; Update the Old Page number Field Object to the new one.
		$oPageNumField = $oNewPageNumField

		$iError = ($oPageNumField.NumberingType() = $iNumFormat) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oPageNumField.Offset = $iOffset
		$iError = ($oPageNumField.Offset() = $iOffset) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iPageNumType <> Null) Then
		If Not __LOWriter_IntIsBetween($iPageNumType, $LOW_PAGE_NUM_TYPE_PREV, $LOW_PAGE_NUM_TYPE_NEXT) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oPageNumField.SubType = $iPageNumType
		$iError = ($oPageNumField.SubType() = $iPageNumType) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($sUserText <> Null) Then
		If Not IsString($sUserText) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oPageNumField.UserText = $sUserText
		$iError = ($oPageNumField.UserText() = $sUserText) ? ($iError) : (BitOR($iError, 8))
	EndIf

	$oPageNumField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldPageNumberModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefBookMarkInsert
; Description ...: Insert a Bookmark Reference Field.
; Syntax ........: _LOWriter_FieldRefBookMarkInsert(ByRef $oDoc, ByRef $oCursor, $sBookmarkName[, $bOverwrite = False[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $sBookmarkName       - a string value. The Bookmark name to Reference.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iRefUsing           - [optional] an integer value (0-4). Default is Null. The Type of reference to use to reference the bookmark, see $LOW_FIELD_REF_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $sBookmarkName not a String.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = Document does not contain a Bookmark by the same name as called in $sBookmarkName.
;				   @Error 1 @Extended 7 Return 0 = $iRefUsing not an Integer, less than 0, or greater than 4.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.GetReference" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Bookmark Reference Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefBookMarkModify, _LOWriter_DocBookmarkInsert, _LOWriter_DocBookmarksList,
;					 _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefBookMarkInsert(ByRef $oDoc, ByRef $oCursor, $sBookmarkName, $bOverwrite = False, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oBookmarkRefField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sBookmarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	If Not _LOWriter_DocBookmarksHasName($oDoc, $sBookmarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oBookmarkRefField = $oDoc.createInstance("com.sun.star.text.TextField.GetReference")
	If Not IsObj($oBookmarkRefField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oBookmarkRefField.SourceName = $sBookmarkName
	$oBookmarkRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_BOOKMARK

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oBookmarkRefField.ReferenceFieldPart = $iRefUsing
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oBookmarkRefField, $bOverwrite)

	$oBookmarkRefField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oBookmarkRefField)
EndFunc   ;==>_LOWriter_FieldRefBookMarkInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefBookMarkModify
; Description ...: Set or Retrieve a Bookmark Reference Field's settings.
; Syntax ........: _LOWriter_FieldRefBookMarkModify(ByRef $oDoc, ByRef $oBookmarkRefField[, $sBookmarkName = Null[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oBookmarkRefField   - [in/out] an object. A Bookmark Reference field Object from a previous _LOWriter_FieldRefBookMarkInsert, or _LOWriter_FieldsGetList function.
;                  $sBookmarkName       - [optional] a string value. Default is Null. The Bookmark name to Reference.
;                  $iRefUsing           - [optional] an integer value (0-4). Default is Null. The Type of reference to use to reference the bookmark, see $LOW_FIELD_REF_USING_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oBookmarkRefField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sBookmarkName not a String.
;				   @Error 1 @Extended 4 Return 0 = Document does not contain a Bookmark by the same name as called in $sBookmarkName.
;				   @Error 1 @Extended 5 Return 0 = $iRefUsing not an Integer, less than 0, or greater than 4.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sBookmarkName
;				   |								2 = Error setting $iRefUsing
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldRefBookMarkInsert, _LOWriter_DocBookmarkInsert, _LOWriter_DocBookmarksList,
;					_LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefBookMarkModify(ByRef $oDoc, ByRef $oBookmarkRefField, $sBookmarkName = Null, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avBook[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oBookmarkRefField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($sBookmarkName, $iRefUsing) Then
		__LOWriter_ArrayFill($avBook, $oBookmarkRefField.SourceName(), $oBookmarkRefField.ReferenceFieldPart())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avBook)
	EndIf

	If ($sBookmarkName <> Null) Then
		If Not IsString($sBookmarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		If Not _LOWriter_DocBookmarksHasName($oDoc, $sBookmarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oBookmarkRefField.SourceName = $sBookmarkName
		$oBookmarkRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_BOOKMARK ;Set Type to Bookmark in case input field Obj is a diff type.
		$iError = ($oBookmarkRefField.SourceName = $sBookmarkName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oBookmarkRefField.ReferenceFieldPart = $iRefUsing
		$iError = ($oBookmarkRefField.ReferenceFieldPart = $iRefUsing) ? ($iError) : (BitOR($iError, 2))
	EndIf

	$oBookmarkRefField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldRefBookMarkModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefEndnoteInsert
; Description ...: Insert a Endnote Reference Field.
; Syntax ........: _LOWriter_FieldRefEndnoteInsert(ByRef $oDoc, ByRef $oCursor, ByRef $oEndNote[, $bOverwrite = False[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $oEndNote            - [in/out] an object. A Endnote Object from a previous _LOWriter_EndnoteInsert, or _LOWriter_EndnotesGetList function.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iRefUsing           - [optional] an integer value (0-4). Default is Null. The Type of reference to use to reference the Endnote, see $LOW_FIELD_REF_USING_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $oEndNote not an Object.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iRefUsing not an Integer, less than 0, or greater than 4.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.GetReference" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Endnote Reference Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefEndnoteModify, _LOWriter_EndnoteInsert, _LOWriter_EndnotesGetList,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefEndnoteInsert(ByRef $oDoc, ByRef $oCursor, ByRef $oEndNote, $bOverwrite = False, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oENoteRefField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsObj($oEndNote) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oENoteRefField = $oDoc.createInstance("com.sun.star.text.TextField.GetReference")
	If Not IsObj($oENoteRefField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oENoteRefField.SourceName = ""
	$oENoteRefField.SequenceNumber = $oEndNote.ReferenceId()
	$oENoteRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_ENDNOTE

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oENoteRefField.ReferenceFieldPart = $iRefUsing
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oENoteRefField, $bOverwrite)

	$oENoteRefField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oENoteRefField)
EndFunc   ;==>_LOWriter_FieldRefEndnoteInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefEndnoteModify
; Description ...: Set or Retrieve a Endnote Reference Field's settings.
; Syntax ........: _LOWriter_FieldRefEndnoteModify(ByRef $oDoc, ByRef $oEndNoteRefField[, $oEndNote = Null[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oEndNoteRefField    - [in/out] an object. A Endnote Reference field Object from a previous _LOWriter_FieldRefEndnoteInsert, or _LOWriter_FieldsGetList function.
;                  $oEndNote            - [optional] an object. Default is Null. A Endnote Object from a previous _LOWriter_EndnoteInsert, or _LOWriter_EndnotesGetList function.
;                  $iRefUsing           - [optional] an integer value (0-4). Default is Null. The Type of reference to use to reference the Endnote, see $LOW_FIELD_REF_USING_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oEndNoteRefField not an Object.
;				   @Error 1 @Extended 3 Return 0 = Optional Parameters set to null, but $oEndNoteRefField object is not a listed as an Endnote Reference type field.
;				   @Error 1 @Extended 4 Return 0 = $oEndNote not an Object.
;				   @Error 1 @Extended 5 Return 0 = $iRefUsing not an Integer, less than 0, or greater than 4.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Endnote Object for setting return.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $oEndNote
;				   |								2 = Error setting $iRefUsing
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldRefEndnoteInsert, _LOWriter_EndnoteInsert, _LOWriter_EndnotesGetList, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefEndnoteModify(ByRef $oDoc, ByRef $oEndNoteRefField, $oEndNote = Null, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iSourceSeq
	Local $avFoot[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oEndNoteRefField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($oEndNote, $iRefUsing) Then
		If Not ($oEndNoteRefField.ReferenceFieldSource() = $LOW_FIELD_REF_TYPE_ENDNOTE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		If $oDoc.Endnotes.hasElements() Then
			$iSourceSeq = $oEndNoteRefField.SequenceNumber()
			For $i = 0 To $oDoc.Endnotes.Count() - 1 ;Locate referenced Endnote.
				If ($oDoc.Endnotes.getByIndex($i).ReferenceId() = $iSourceSeq) Then
					__LOWriter_ArrayFill($avFoot, $oDoc.Endnotes.getByIndex($i), $oEndNoteRefField.ReferenceFieldPart())
					Return SetError($__LO_STATUS_SUCCESS, 1, $avFoot)
				EndIf
				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
			Next

		EndIf
		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Error retrieving EndNote Obj
	EndIf

	If ($oEndNote <> Null) Then
		If Not IsObj($oEndNote) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oEndNoteRefField.SourceName = ""
		$oEndNoteRefField.SequenceNumber = $oEndNote.ReferenceId()
		$oEndNoteRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_ENDNOTE ;Set Type to Endnote in case input field Obj is a diff type.
		$iError = ($oEndNoteRefField.SequenceNumber = $oEndNote.ReferenceId()) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oEndNoteRefField.ReferenceFieldPart = $iRefUsing
		$iError = ($oEndNoteRefField.ReferenceFieldPart = $iRefUsing) ? ($iError) : (BitOR($iError, 2))
	EndIf

	$oEndNoteRefField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldRefEndnoteModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefFootnoteInsert
; Description ...: Insert a Footnote Reference Field.
; Syntax ........: _LOWriter_FieldRefFootnoteInsert(ByRef $oDoc, ByRef $oCursor, ByRef $oFootNote[, $bOverwrite = False[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $oFootNote           - [in/out] an object. A Footnote Object from a previous _LOWriter_FootnoteInsert, Or _LOWriter_FootnotesGetList function.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iRefUsing           - [optional] an integer value (0-4). Default is Null. The Type of reference to use to reference the Footnote, see $LOW_FIELD_REF_USING_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $oFootNote not an Object.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iRefUsing not an Integer, less than 0, or greater than 4.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.GetReference" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Footnote Reference Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefFootnoteModify, _LOWriter_FootnoteInsert, _LOWriter_FootnotesGetList,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefFootnoteInsert(ByRef $oDoc, ByRef $oCursor, ByRef $oFootNote, $bOverwrite = False, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oFNoteRefField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsObj($oFootNote) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oFNoteRefField = $oDoc.createInstance("com.sun.star.text.TextField.GetReference")
	If Not IsObj($oFNoteRefField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$oFNoteRefField.SourceName = ""
	$oFNoteRefField.SequenceNumber = $oFootNote.ReferenceId()
	$oFNoteRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_FOOTNOTE

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oFNoteRefField.ReferenceFieldPart = $iRefUsing
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oFNoteRefField, $bOverwrite)

	$oFNoteRefField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oFNoteRefField)
EndFunc   ;==>_LOWriter_FieldRefFootnoteInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefFootnoteModify
; Description ...: Set or Retrieve a Footnote Reference Field's settings.
; Syntax ........: _LOWriter_FieldRefFootnoteModify(ByRef $oDoc, ByRef $oFootNoteRefField[, $oFootNote = Null[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oFootNoteRefField   - [in/out] an object. A Footnote Reference field Object from a previous _LOWriter_FieldRefFootnoteInsert, or _LOWriter_FieldsGetList function.
;                  $oFootNote           - [optional] an object. Default is Null. A Footnote Object from a previous _LOWriter_FootnoteInsert, Or _LOWriter_FootnotesGetList function.
;                  $iRefUsing           - [optional] an integer value (0-4). Default is Null. The Type of reference to use to reference the Footnote, see $LOW_FIELD_REF_USING_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oFootNoteRefField not an Object.
;				   @Error 1 @Extended 3 Return 0 = Optional Parameters set to null, but $oFootNoteRefField object is not a listed as a Footnote Reference type field.
;				   @Error 1 @Extended 4 Return 0 = $oFootNote not an Object.
;				   @Error 1 @Extended 5 Return 0 = $iRefUsing not an Integer, less than 0, or greater than 4.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Error retrieving Footnote Object for setting return.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $oFootNote
;				   |								2 = Error setting $iRefUsing
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldRefFootnoteInsert, _LOWriter_FootnoteInsert, _LOWriter_FootnotesGetList, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefFootnoteModify(ByRef $oDoc, ByRef $oFootNoteRefField, $oFootNote = Null, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iSourceSeq
	Local $avFoot[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oFootNoteRefField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($oFootNote, $iRefUsing) Then
		If Not ($oFootNoteRefField.ReferenceFieldSource() = $LOW_FIELD_REF_TYPE_FOOTNOTE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		If $oDoc.Footnotes.hasElements() Then
			$iSourceSeq = $oFootNoteRefField.SequenceNumber()
			For $i = 0 To $oDoc.Footnotes.Count() - 1 ;Locate referenced Footnote.
				If ($oDoc.Footnotes.getByIndex($i).ReferenceId() = $iSourceSeq) Then
					__LOWriter_ArrayFill($avFoot, $oDoc.Footnotes.getByIndex($i), $oFootNoteRefField.ReferenceFieldPart())
					Return SetError($__LO_STATUS_SUCCESS, 1, $avFoot)
				EndIf
				Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV)) ? (10) : (0))
			Next

		EndIf
		Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0) ; Error retrieving FndNote Obj
	EndIf

	If ($oFootNote <> Null) Then
		If Not IsObj($oFootNote) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oFootNoteRefField.SourceName = ""
		$oFootNoteRefField.SequenceNumber = $oFootNote.ReferenceId()
		$oFootNoteRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_FOOTNOTE ;Set Type to Footnote in case input field Obj is a diff type.
		$iError = ($oFootNoteRefField.SequenceNumber = $oFootNote.ReferenceId()) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oFootNoteRefField.ReferenceFieldPart = $iRefUsing
		$iError = ($oFootNoteRefField.ReferenceFieldPart = $iRefUsing) ? ($iError) : (BitOR($iError, 2))
	EndIf

	$oFootNoteRefField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldRefFootnoteModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefGetType
; Description ...: Retrieve the type of Data a Reference Field is Referencing.
; Syntax ........: _LOWriter_FieldRefGetType(ByRef $oRefField)
; Parameters ....: $oRefField           - [in/out] an object. a Reference Field Object from a previous Insert or _LOWriter_FieldsGetList function.
; Return values .: Success: Integer
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oRefField not an Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Integer = Success. Returning the Data Type Source for the reference Field. See constants, $LOW_FIELD_REF_TYPE_* as defined in LibreOfficeWriter_Constants.au3
; Author ........: donnyh13
; Modified ......:
; Remarks .......: A Reference Field can be referencing multiple different types of Data, such as a Reference Mark, or Bookmark, etc.
; Related .......: _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefGetType(ByRef $oRefField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	If Not IsObj($oRefField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oRefField.ReferenceFieldSource())
EndFunc   ;==>_LOWriter_FieldRefGetType

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefInsert
; Description ...: Insert a Reference Field.
; Syntax ........: _LOWriter_FieldRefInsert(ByRef $oDoc, ByRef $oCursor, $sRefMarkName[, $bOverwrite = False[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $sRefMarkName        - a string value. The Reference Mark Name to reference.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iRefUsing           - [optional] an integer value (0-4). Default is Null. The Type of reference to insert, see $LOW_FIELD_REF_USING_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $sRefMarkName not a String.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = Document does not contain a Reference Mark by the same name as called in $sRefMarkName.
;				   @Error 1 @Extended 7 Return 0 = $iRefUsing not an Integer, less than 0, or greater than 4.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Reference Marks Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to create "com.sun.star.text.TextField.GetReference" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Reference Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefModify, _LOWriter_FieldRefMarkSet, _LOWriter_FieldRefMarkList,
;					_LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor,
;					_LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor,
;					_LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefInsert(ByRef $oDoc, ByRef $oCursor, $sRefMarkName, $bOverwrite = False, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMarks, $oMarkRefField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sRefMarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oRefMarks = $oDoc.getReferenceMarks()
	If Not IsObj($oRefMarks) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not $oRefMarks.hasByName($sRefMarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oMarkRefField = $oDoc.createInstance("com.sun.star.text.TextField.GetReference")
	If Not IsObj($oMarkRefField) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oMarkRefField.SourceName = $sRefMarkName
	$oMarkRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_REF_MARK

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oMarkRefField.ReferenceFieldPart = $iRefUsing
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oMarkRefField, $bOverwrite)

	$oMarkRefField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oMarkRefField)
EndFunc   ;==>_LOWriter_FieldRefInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefMarkDelete
; Description ...: Delete a Reference Mark by name.
; Syntax ........: _LOWriter_FieldRefMarkDelete(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sName               - a string value. The Reference Mark name to delete.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sName not a String.
;				   @Error 1 @Extended 3 Return 0 = Document does not contain a Reference Mark named the same as called in $sName
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Reference Marks Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Reference Mark object called in $sName.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to delete Reference Mark, but document still contains a Reference Mark by that name.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested Reference Mark.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefMarkSet, _LOWriter_FieldRefMarkList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefMarkDelete(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMark, $oRefMarks

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRefMarks = $oDoc.getReferenceMarks()
	If Not IsObj($oRefMarks) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not $oRefMarks.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oRefMark = $oRefMarks.getByName($sName)
	If Not IsObj($oRefMark) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oRefMark.dispose()

	Return ($oRefMarks.hasByName($sName)) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldRefMarkDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefMarkGetAnchor
; Description ...: Retrieve the Anchor Cursor Object of a Reference Mark by Name.
; Syntax ........: _LOWriter_FieldRefMarkGetAnchor(ByRef $oDoc, $sName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sName               - a string value. The Reference Mark name to retrieve the anchor for.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sName not a String.
;				   @Error 1 @Extended 3 Return 0 = Document does not contain a Reference Mark named the same as called in $sName
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Reference Marks Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Reference Mark object called in $sName.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Returning requested Reference Mark Anchor Cursor Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefMarkList, _LOWriter_CursorMove
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefMarkGetAnchor(ByRef $oDoc, $sName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMark, $oRefMarks

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oRefMarks = $oDoc.getReferenceMarks()
	If Not IsObj($oRefMarks) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not $oRefMarks.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oRefMark = $oRefMarks.getByName($sName)
	If Not IsObj($oRefMark) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oRefMark.Anchor.Text.createTextCursorByRange($oRefMark.Anchor()))
EndFunc   ;==>_LOWriter_FieldRefMarkGetAnchor

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefMarkList
; Description ...: Retrieve an Array of Reference Mark names.
; Syntax ........: _LOWriter_FieldRefMarkList(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Reference Marks Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Array of Reference Mark Names.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. No Reference Marks found in document.
;				   @Error 0 @Extended ? Return Array = Success. Successfully searched for Reference Marks, returning Array of Reference Mark Names, with @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefMarkDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefMarkList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMarks
	Local $asRefMarks[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oRefMarks = $oDoc.getReferenceMarks()
	If Not IsObj($oRefMarks) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$asRefMarks = $oRefMarks.getElementNames()
	If Not IsArray($asRefMarks) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return (UBound($asRefMarks) > 0) ? (SetError($__LO_STATUS_SUCCESS, UBound($asRefMarks), $asRefMarks)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldRefMarkList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefMarkSet
; Description ...: Create and Insert a Reference Mark at a Cursor position.
; Syntax ........: _LOWriter_FieldRefMarkSet(ByRef $oDoc, ByRef $oCursor, $sName[, $bOverwrite = False])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $sName               - a string value. The name of the Reference Mark to create.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $sName not a String.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = Document already contains a Reference Mark by the same name as called in $sName.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error retrieving Reference Marks Object.
;				   @Error 2 @Extended 2 Return 0 = Error creating "com.sun.star.text.ReferenceMark" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1. = Success. Successfully created a Reference Mark.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldRefMarkDelete, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefMarkSet(ByRef $oDoc, ByRef $oCursor, $sName, $bOverwrite = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMark, $oRefMarks

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$oRefMarks = $oDoc.getReferenceMarks()
	If Not IsObj($oRefMarks) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If $oRefMarks.hasByName($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oRefMark = $oDoc.createInstance("com.sun.star.text.ReferenceMark")
	If Not IsObj($oRefMark) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oRefMark.Name = $sName

	$oCursor.Text.insertTextContent($oCursor, $oRefMark, $bOverwrite)

	Return SetError($__LO_STATUS_SUCCESS, 0, 1)
EndFunc   ;==>_LOWriter_FieldRefMarkSet

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldRefModify
; Description ...: Set or Retrieve a Reference Field's settings.
; Syntax ........: _LOWriter_FieldRefModify(ByRef $oDoc, ByRef $oRefField[, $sRefMarkName = Null[, $iRefUsing = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oRefField           - [in/out] an object. A Reference field Object from a previous _LOWriter_FieldRefInsert or _LOWriter_FieldsGetList function.
;                  $sRefMarkName        - [optional] a string value. Default is Null. The Reference Mark Name to Reference.
;                  $iRefUsing           - [optional] an integer value (0-4). Default is Null. The Type of reference to insert, see $LOW_FIELD_REF_USING_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oRefField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sRefMarkName not a String.
;				   @Error 1 @Extended 4 Return 0 = Document does not contain a Reference Mark by the same name as called in $sRefMarkName.
;				   @Error 1 @Extended 6 Return 0 = $iRefUsing not an Integer, less than 0, or greater than 4.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve Reference Marks Object.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sRefMarkName
;				   |								2 = Error setting $iRefUsing
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldRefInsert, _LOWriter_FieldsGetList, _LOWriter_FieldRefMarkSet, _LOWriter_FieldRefMarkList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldRefModify(ByRef $oDoc, ByRef $oRefField, $sRefMarkName = Null, $iRefUsing = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oRefMarks
	Local $iError = 0
	Local $avRef[2]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oRefField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($sRefMarkName, $iRefUsing) Then
		__LOWriter_ArrayFill($avRef, $oRefField.SourceName(), $oRefField.ReferenceFieldPart())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avRef)
	EndIf

	If ($sRefMarkName <> Null) Then
		If Not IsString($sRefMarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oRefMarks = $oDoc.getReferenceMarks()
		If Not IsObj($oRefMarks) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
		If Not $oRefMarks.hasByName($sRefMarkName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oRefField.SourceName = $sRefMarkName
		$oRefField.ReferenceFieldSource = $LOW_FIELD_REF_TYPE_REF_MARK ;Set Type to RefMark in case input field Obj is a diff type.
		$iError = ($oRefField.SourceName = $sRefMarkName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iRefUsing <> Null) Then
		If Not __LOWriter_IntIsBetween($iRefUsing, $LOW_FIELD_REF_USING_PAGE_NUM_UNSTYLED, $LOW_FIELD_REF_USING_PAGE_NUM_STYLED) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oRefField.ReferenceFieldPart = $iRefUsing
		$iError = ($oRefField.ReferenceFieldPart = $iRefUsing) ? ($iError) : (BitOR($iError, 2))
	EndIf

	$oRefField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldRefModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldsAdvGetList
; Description ...: Retrieve an Array of Advanced Field Objects contained in a document.
; Syntax ........: _LOWriter_FieldsAdvGetList(ByRef $oDoc[, $iType = $LOW_FIELD_ADV_TYPE_ALL[, $bSupportedServices = True[, $bFieldType = True[, $bFieldTypeNum = True]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iType               - [optional] an integer value (1-1023). Default is $LOW_FIELD_ADV_TYPE_ALL. The type of Field to search for. See Constants, $LOW_FIELD_ADV_TYPE_* as defined in LibreOfficeWriter_Constants.au3. Can be BitOr'd together.
;                  $bSupportedServices  - [optional] a boolean value. Default is True. If True, adds a column to the array that has the supported service String for that particular Field, To assist in identifying the Field type.
;                  $bFieldType          - [optional] a boolean value. Default is True. If True, adds a column to the array that has the Field Type String for that particular Field as described by Libre Office. To assist in identifying the Field type.
;                  $bFieldTypeNum       - [optional] a boolean value. Default is True. If True, adds a column to the array that has the Field Type Constant Integer for that particular Field, to assist in identifying the Field type. See Constants, $LOW_FIELD_ADV_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than 1, or greater than 1023. (The total of all Constants added together.) See Constants, $LOW_FIELD_ADV_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $bSupportedServices not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bFieldType not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bFieldTypeNum not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $avFieldTypes passed to internal function not an Array. UDF needs fixed.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error converting Field type Constants.
;				   @Error 2 @Extended 2 Return 0 = Failed to create enumeration of fields in document.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning Array of Text Field Objects with @Extended set to number of results. See Remarks for Array sizing.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:The Array can vary in the number of columns, if $bSupportedServices, $bFieldType, and $bFieldTypeNum are set
;					to False, the Array will be a single column. With each of the above listed options being set to True, a
;					column will be added in the order they are listed in the UDF parameters. The First column will always be the
;					Field Object.
;					Setting $bSupportedServices to True will add a Supported Service String column for the found Field.
;					Setting $bFieldType to True will add a Field type column for the found Field.
;					Setting $bFieldTypeNum to True will add a Field type Number column, matching the constants, $LOW_FIELD_ADV_TYPE_* as defined in LibreOfficeWriter_Constants.au3 for the found Field.
;					Note: For simplicity, and also due to certain Bit limitations I have broken the different Field types into
;						three different categories, Regular Fields, ($LWFieldType), Advanced(Complex) Fields, ($LWFieldAdvType),
;						and Document Information fields (Found in the Document Information Tab in L.O. Fields dialog),
;						($LWFieldDocInfoType). Just because a field is listed in the constants below, does not necessarily mean
;						that I have made a function to create/modify it, you may still be able to update or delete it using the
;						_LOWriter_FieldUpdate, or _LOWriter_FieldDelete function, though. Some Fields are too complex to create a function for,
;						and others are literally impossible.
; Related .......: _LOWriter_FieldsDocInfoGetList, _LOWriter_FieldsGetList, _LOWriter_FieldDelete, _LOWriter_FieldGetAnchor,
;					_LOWriter_FieldUpdate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldsAdvGetList(ByRef $oDoc, $iType = $LOW_FIELD_ADV_TYPE_ALL, $bSupportedServices = True, $bFieldType = True, $bFieldTypeNum = True)
	Local $avFieldTypes[0][0]
	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_IntIsBetween($iType, $LOW_FIELD_ADV_TYPE_ALL, 1023) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	; 1023 is all possible Consts added together
	If Not IsBool($bSupportedServices) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bFieldType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bFieldTypeNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$avFieldTypes = __LOWriter_FieldTypeServices($iType, True, False)
	If @error > 0 Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$vReturn = __LOWriter_FieldsGetList($oDoc, $bSupportedServices, $bFieldType, $bFieldTypeNum, $avFieldTypes)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FieldsAdvGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldsDocInfoGetList
; Description ...: Retrieve an Array of Document Information Field Objects contained in a document.
; Syntax ........: _LOWriter_FieldsDocInfoGetList(ByRef $oDoc[, $iType = $LOW_FIELD_DOCINFO_TYPE_ALL[, $bSupportedServices = True[, $bFieldType = True[, $bFieldTypeNum = True]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iType               - [optional] an integer value (1-16383). Default is $LOW_FIELD_DOCINFO_TYPE_ALL. The type of Field to search for. See Constants, $LOW_FIELD_DOCINFO_TYPE_* as defined in LibreOfficeWriter_Constants.au3. Can be BitOr'd together.
;                  $bSupportedServices  - [optional] a boolean value. Default is True. If True, adds a column to the array that has the supported service String for that particular Field, To assist in identifying the Field type.
;                  $bFieldType          - [optional] a boolean value. Default is True. If True, adds a column to the array that has the Field Type String for that particular Field as described by Libre Office. To assist in identifying the Field type.
;                  $bFieldTypeNum       - [optional] a boolean value. Default is True. If True, adds a column to the array that has the Field Type Constant Integer for that particular Field, to assist in identifying the Field type. See Constants, $LOW_FIELD_DOCINFO_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than 1, or greater than 16383. (The total of all Constants added together.) See Constants, $LOW_FIELD_DOCINFO_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $bSupportedServices not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bFieldType not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bFieldTypeNum not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $avFieldTypes passed to internal function not an Array. UDF needs fixed.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error converting Field type Constants.
;				   @Error 2 @Extended 2 Return 0 = Failed to create enumeration of fields in document.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning Array of Text Field Objects with @Extended set to number of results. See Remarks for Array sizing.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Array can vary in the number of columns, if $bSupportedServices, $bFieldType, and $bFieldTypeNum are set
;					to False, the Array will be a single column. With each of the above listed options being set to True, a
;					column will be added in the order they are listed in the UDF parameters. The First column will always be the
;					Field Object.
;					Setting $bSupportedServices to True will add a Supported Service String column for the found Field.
;					Setting $bFieldType to True will add a Field type column for the found Field.
;					Setting $bFieldTypeNum to True will add a Field type Number column, matching the constants, $LOW_FIELD_DOCINFO_TYPE_* as defined in LibreOfficeWriter_Constants.au3 for the
;						found Field.
;					Note: For simplicity, and also due to certain Bit limitations I have broken the different Field types into
;						three different categories, Regular Fields, ($LWFieldType), Advanced(Complex) Fields, ($LWFieldAdvType),
;						and Document Information fields (Found in the Document Information Tab in L.O. Fields dialog),
;						($LWFieldDocInfoType). Just because a field is listed in the constants below, does not necessarily mean
;						that I have made a function to create/modify it, you may still be able to update or delete it using the
;						Field Update, or Field Delete function, though. Some Fields are too complex to create a function for,
;						and others are literally impossible.
; Related .......: _LOWriter_FieldsAdvGetList, _LOWriter_FieldsGetList, _LOWriter_FieldDelete, _LOWriter_FieldGetAnchor,
;					_LOWriter_FieldUpdate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldsDocInfoGetList(ByRef $oDoc, $iType = $LOW_FIELD_DOCINFO_TYPE_ALL, $bSupportedServices = True, $bFieldType = True, $bFieldTypeNum = True)
	Local $avFieldTypes[0][0]
	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_IntIsBetween($iType, $LOW_FIELD_ADV_TYPE_ALL, 16383) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	; 16383 is all possible Consts added together
	If Not IsBool($bSupportedServices) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bFieldType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bFieldTypeNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$avFieldTypes = __LOWriter_FieldTypeServices($iType, False, True)
	If @error > 0 Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$vReturn = __LOWriter_FieldsGetList($oDoc, $bSupportedServices, $bFieldType, $bFieldTypeNum, $avFieldTypes)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FieldsDocInfoGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSenderInsert
; Description ...: Insert a Sender Field.
; Syntax ........: _LOWriter_FieldSenderInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bIsFixed = Null[, $sContent = Null[, $iDataType = Null]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sContent            - [optional] a string value. Default is Null. The Content to Display, only valid if $bIsFixed is set to True.
;                  $iDataType           - [optional] an integer value (0-14). Default is Null. The Data Type to display. See Constants, $LOW_FIELD_USER_DATA_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $sContent not a String.
;				   @Error 1 @Extended 7 Return 0 = $iDataType not an Integer, less than 0, or greater than 14. See Constants, $LOW_FIELD_USER_DATA_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.ExtendedUser" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Sender field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldSenderModify,  _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSenderInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bIsFixed = Null, $sContent = Null, $iDataType = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSenderField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oSenderField = $oDoc.createInstance("com.sun.star.text.TextField.ExtendedUser")
	If Not IsObj($oSenderField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oSenderField.IsFixed = $bIsFixed
	EndIf

	If ($sContent <> Null) Then
		If Not IsString($sContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oSenderField.Content = $sContent
	EndIf

	If ($iDataType <> Null) Then
		If Not __LOWriter_IntIsBetween($iDataType, $LOW_FIELD_USER_DATA_COMPANY, $LOW_FIELD_USER_DATA_STATE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oSenderField.UserDataType = $iDataType
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oSenderField, $bOverwrite)

	If ($sContent <> Null) Then ; Sometimes Content Disappears upon Insertion, make a check to re-set the Content value.
		If $oSenderField.Content <> $sContent And ($oSenderField.IsFixed() = True) Then $oSenderField.Content = $sContent
	EndIf

	If ($oSenderField.IsFixed() = False) Then $oSenderField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oSenderField)
EndFunc   ;==>_LOWriter_FieldSenderInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSenderModify
; Description ...: Set or Retrieve a Sender Field's settings.
; Syntax ........: _LOWriter_FieldSenderModify(ByRef $oSenderField[, $bIsFixed = Null[, $sContent = Null[, $iDataType = Null]]])
; Parameters ....: $oSenderField        - [in/out] an object. A Sender field Object from a previous _LOWriter_FieldSenderInsert, or _LOWriter_FieldsGetList function.
;                  $bIsFixed            - [optional] a boolean value. Default is Null. If True, the value is static, that is, the value does not update if the source changes or all fields are updated.
;                  $sContent            - [optional] a string value. Default is Null. The Content to Display, only valid if $bIsFixed is set to True.
;                  $iDataType           - [optional] an integer value (0-14). Default is Null. The Data Type to display. See Constants, $LOW_FIELD_USER_DATA_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oSenderField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bIsFixed not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $sContent not a String.
;				   @Error 1 @Extended 4 Return 0 = $iDataType not an Integer, less than 0, or greater than 14. See Constants, $LOW_FIELD_USER_DATA_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bIsFixed
;				   |								2 = Error setting $sContent
;				   |								4 = Error setting $iDataType
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldSenderInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSenderModify(ByRef $oSenderField, $bIsFixed = Null, $sContent = Null, $iDataType = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avExtUser[3]

	If Not IsObj($oSenderField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bIsFixed, $sContent, $iDataType) Then
		__LOWriter_ArrayFill($avExtUser, $oSenderField.IsFixed(), $oSenderField.Content(), $oSenderField.UserDataType())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avExtUser)
	EndIf

	If ($bIsFixed <> Null) Then
		If Not IsBool($bIsFixed) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oSenderField.IsFixed = $bIsFixed
		$iError = ($oSenderField.IsFixed() = $bIsFixed) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($sContent <> Null) Then
		If Not IsString($sContent) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oSenderField.Content = $sContent
		$iError = ($oSenderField.Content() = $sContent) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($iDataType <> Null) Then
		If Not __LOWriter_IntIsBetween($iDataType, $LOW_FIELD_USER_DATA_COMPANY, $LOW_FIELD_USER_DATA_STATE) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oSenderField.UserDataType = $iDataType
		$iError = ($oSenderField.UserDataType() = $iDataType) ? ($iError) : (BitOR($iError, 4))
	EndIf

	If ($oSenderField.IsFixed() = False) Then $oSenderField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldSenderModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarInsert
; Description ...: Insert a Set Variable Field.
; Syntax ........: _LOWriter_FieldSetVarInsert(ByRef $oDoc, ByRef $oCursor, $sName, $sValue[, $bOverwrite = False[, $iNumFormatKey = Null[, $bIsVisible = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $sName               - a string value. The name of the Set Variable Field to Create, If the name matches an already existing Set Variable Master Field, that Master will be used, else a new Set Variable Masterfield will be created.
;                  $sValue              - a string value. The Set Variable Field's value.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iNumFormatKey       - [optional] an integer value. Default is Null. The Number Format Key to use for displaying this variable.
;                  $bIsVisible          - [optional] a boolean value. Default is Null. If False, the Set Variable Field is invisible. L.O.'s default is True.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $sName not a String.
;				   @Error 1 @Extended 5 Return 0 = $sValue not a String.
;				   @Error 1 @Extended 6 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 7 Return 0 = $iNumFormatKeyKey not an Integer.
;				   @Error 1 @Extended 8 Return 0 = $iNumFormatKeyKey not equal to -1 and Number Format key called in $iNumFormatKeyKey not found in document.
;				   @Error 1 @Extended 9 Return 0 = $bIsVisible not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.SetExpression" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Set Variable field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldSetVarModify,  _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor,
;					_LOWriter_FormatKeyCreate _LOWriter_FormatKeyList, _LOWriter_FieldSetVarMasterCreate,
;					_LOWriter_FieldSetVarMasterList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarInsert(ByRef $oDoc, ByRef $oCursor, $sName, $sValue, $bOverwrite = False, $iNumFormatKey = Null, $bIsVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oSetVarField, $oSetVarMaster
	Local $iExtended = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsString($sName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsString($sValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)

	$oSetVarField = $oDoc.createInstance("com.sun.star.text.TextField.SetExpression")
	If Not IsObj($oSetVarField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If _LOWriter_FieldSetVarMasterExists($oDoc, $sName) Then
		$oSetVarMaster = _LOWriter_FieldSetVarMasterGetObj($oDoc, $sName)
		$iExtended = 1 ;1 = Master already existed.
	Else
		$oSetVarMaster = _LOWriter_FieldSetVarMasterCreate($oDoc, $sName)
	EndIf

	If Not IsObj($oSetVarMaster) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oSetVarField.Content = $sValue

	If ($iNumFormatKey <> Null) Then
		If Not IsInt($iNumFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		If ($iNumFormatKey <> -1) And Not _LOWriter_FormatKeyExists($oDoc, $iNumFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$oSetVarField.NumberFormat = $iNumFormatKey
	Else
		$oSetVarField.NumberFormat = 0 ; If No Input, set to General
	EndIf

	If ($bIsVisible <> Null) Then
		If Not IsBool($bIsVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		$oSetVarField.IsVisible = $bIsVisible
	EndIf

	$oSetVarField.attachTextFieldMaster($oSetVarMaster)

	$oCursor.Text.insertTextContent($oCursor, $oSetVarField, $bOverwrite)

	$oSetVarField.Update()

	Return SetError($__LO_STATUS_SUCCESS, $iExtended, $oSetVarField)
EndFunc   ;==>_LOWriter_FieldSetVarInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterCreate
; Description ...: Create a Set Variable Master Field.
; Syntax ........: _LOWriter_FieldSetVarMasterCreate(ByRef $oDoc, $sMasterFieldName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sMasterFieldName    - a string value. The Set Variable Master Field name to create.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sMasterFieldName not a String.
;				   @Error 1 @Extended 3 Return 0 = Document already contains a MasterField by the name called in $sMasterFieldName
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve MasterFields Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to Create MasterField Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully created the MasterField, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldSetVarMasterDelete, _LOWriter_FieldSetVarInsert
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterCreate(ByRef $oDoc, $sMasterFieldName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMasterFields, $oMasterfield
	Local $sFullFieldName
	Local $sField = "com.sun.star.text.fieldmaster.SetExpression"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sMasterFieldName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$sFullFieldName = $sField & "." & $sMasterFieldName
	$oMasterFields = $oDoc.getTextFieldMasters()
	If Not IsObj($oMasterFields) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	If $oMasterFields.hasByName($sFullFieldName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oMasterfield = $oDoc.createInstance($sField)
	If Not IsObj($oMasterfield) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oMasterfield.Name = $sMasterFieldName

	Return SetError($__LO_STATUS_SUCCESS, 0, $oMasterfield)
EndFunc   ;==>_LOWriter_FieldSetVarMasterCreate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterDelete
; Description ...: Delete a Set Variable Master Field.
; Syntax ........: _LOWriter_FieldSetVarMasterDelete(ByRef $oDoc, $vMasterField)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $vMasterField        - a variant value. The Set Variable Master Field name or object from _LOWriter_FieldSetVarMasterCreate, _LOWriter_FieldSetVarMasterGetObj, or _LOWriter_FieldSetVarMasterList to delete.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $vMasterField not a String and not an Object.
;				   @Error 1 @Extended 3 Return 0 = $vMasterField is a String, but document does not contain a Masterfield by that name.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve MasterFields Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve MasterField object called in $vMasterField.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Attempted to delete MasterField, but document still contains a MasterField by that name.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully deleted requested MasterField.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldSetVarMasterCreate, _LOWriter_FieldSetVarMasterGetObj, _LOWriter_FieldSetVarMasterList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterDelete(ByRef $oDoc, $vMasterField)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMasterFields, $oMasterfield
	Local $sFullFieldName
	Local $sField = "com.sun.star.text.fieldmaster.SetExpression"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($vMasterField) And Not IsObj($vMasterField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$oMasterFields = $oDoc.getTextFieldMasters()
	If Not IsObj($oMasterFields) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If IsObj($vMasterField) Then
		$sFullFieldName = $sField & "." & $vMasterField.Name()
		$oMasterfield = $vMasterField
	Else
		$sFullFieldName = $sField & "." & $vMasterField
		If Not $oMasterFields.hasByName($sFullFieldName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oMasterfield = $oMasterFields.getByName($sFullFieldName)
	EndIf

	If Not IsObj($oMasterfield) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	$oMasterfield.dispose()

	Return ($oMasterFields.hasByName($sFullFieldName)) ? (SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldSetVarMasterDelete

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterExists
; Description ...: Check if a document contains a Set Variable Master Field by name.
; Syntax ........: _LOWriter_FieldSetVarMasterExists(ByRef $oDoc, $sMasterFieldName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sMasterFieldName    - a string value. The Set Variable Master Field name to look for.
; Return values .: Success: Boolean
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sMasterFieldName not a String.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve MasterFields Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Boolean = Success. If the document contains a MasterField called in $sMasterFieldName, then True is returned, Else false.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterExists(ByRef $oDoc, $sMasterFieldName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMasterFields
	Local $sFullFieldName = "com.sun.star.text.fieldmaster.SetExpression" & "."

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sMasterFieldName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$sFullFieldName &= $sMasterFieldName

	$oMasterFields = $oDoc.getTextFieldMasters()
	If Not IsObj($oMasterFields) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	If $oMasterFields.hasByName($sFullFieldName) Then Return SetError($__LO_STATUS_SUCCESS, 1, True)

	Return SetError($__LO_STATUS_SUCCESS, 0, False)
EndFunc   ;==>_LOWriter_FieldSetVarMasterExists

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterGetObj
; Description ...: Retrieve a Set Variable Master Field Object.
; Syntax ........: _LOWriter_FieldSetVarMasterGetObj(ByRef $oDoc, $sMasterFieldName)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $sMasterFieldName    - a string value. The Set Variable Master Field to retrieve the Object for.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $sMasterFieldName not an Object.
;				   @Error 1 @Extended 3 Return 0 = Document does not contain FieldMaster named as called in $sMasterFieldName.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve MasterFields Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve requested FieldMaster Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully retrieved requested FieldMaster Object, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldSetVarMasterList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterGetObj(ByRef $oDoc, $sMasterFieldName)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMasterFields, $oMasterfield
	Local $sFullFieldName = "com.sun.star.text.fieldmaster.SetExpression" & "."

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsString($sMasterFieldName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$sFullFieldName &= $sMasterFieldName

	$oMasterFields = $oDoc.getTextFieldMasters()
	If Not IsObj($oMasterFields) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)
	If Not $oMasterFields.hasByName($sFullFieldName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	$oMasterfield = $oMasterFields.getByName($sFullFieldName)
	If Not IsObj($oMasterfield) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	Return SetError($__LO_STATUS_SUCCESS, 0, $oMasterfield)
EndFunc   ;==>_LOWriter_FieldSetVarMasterGetObj

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterList
; Description ...: Retrieve a List of current Set Variable Master Fields in a document.
; Syntax ........: _LOWriter_FieldSetVarMasterList(ByRef $oDoc)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve MasterFields Object.
;				   @Error 2 @Extended 2 Return 0 = Failed to retrieve Array of MasterField Objects.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Successfully retrieved Array of Set Variable MasterField Names, returning Array of Set Variable MasterField Names with @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:Note: This function includes in the list about 5 built-in Master Fields from Libre Office, namely: Illustration, Table, Text, Drawing, and Figure.
; Related .......: _LOWriter_FieldSetVarMasterGetObj, _LOWriter_FieldSetVarMasterDelete
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterList(ByRef $oDoc)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oMasterFields
	Local $asMasterFields[0], $asSetVarMasters[0]
	Local $iCount = 0
	Local $sField = "com.sun.star.text.fieldmaster.SetExpression"

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	$oMasterFields = $oDoc.getTextFieldMasters()
	If Not IsObj($oMasterFields) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$asMasterFields = $oMasterFields.getElementNames()
	If Not IsArray($asMasterFields) Then Return SetError($__LO_STATUS_INIT_ERROR, 2, 0)

	ReDim $asSetVarMasters[UBound($asMasterFields)]

	For $i = 0 To UBound($asMasterFields) - 1
		If ($oMasterFields.getByName($asMasterFields[$i]).supportsService($sField)) Then
			$asSetVarMasters[$iCount] = $oMasterFields.getByName($asMasterFields[$i]).Name()
			$iCount += 1
		EndIf

		Sleep((IsInt($i / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	Next

	ReDim $asSetVarMasters[$iCount]

	Return SetError($__LO_STATUS_SUCCESS, $iCount, $asSetVarMasters)
EndFunc   ;==>_LOWriter_FieldSetVarMasterList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarMasterListFields
; Description ...: Return an Array of Objects of dependent fields for a specific Master Field.
; Syntax ........: _LOWriter_FieldSetVarMasterListFields(ByRef $oDoc, ByRef $oMasterfield)
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oMasterfield        - [in/out] an object. The Set Variable Master Field Object returned from a previous _LOWriter_FieldSetVarMasterCreate, or _LOWriter_FieldSetVarMasterGetObj function.
; Return values .: Success: 1 or Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oMasterfield not an Object.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve dependent fields Array.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully searched for dependent fields, but MasterField has none.
;				   @Error 0 @Extended ? Return Array = Success. Successfully searched for dependent fields, returning Array of dependent SetVariable Fields, with @Extended set to number of results.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Dependent Fields are SetVariable Fields that are referencing the Master field.
; Related .......: _LOWriter_FieldSetVarMasterGetObj
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarMasterListFields(ByRef $oDoc, ByRef $oMasterfield)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $aoDependFields[0]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oMasterfield) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	$aoDependFields = $oMasterfield.DependentTextFields()
	If Not IsArray($aoDependFields) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	Return (UBound($aoDependFields) > 0) ? (SetError($__LO_STATUS_SUCCESS, UBound($aoDependFields), $aoDependFields)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldSetVarMasterListFields

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldSetVarModify
; Description ...: Set or Retrieve a Set Variable Field's settings.
; Syntax ........: _LOWriter_FieldSetVarModify(ByRef $oDoc, ByRef $oSetVarField[, $sValue = Null[, $iNumFormatKey = Null[, $bIsVisible = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oSetVarField        - [in/out] an object. A Set Variable field Object from a previous _LOWriter_FieldSetVarInsert, or _LOWriter_FieldsGetList function.
;                  $sValue              - [optional] a string value. Default is Null. The Set Variable Field's value.
;                  $iNumFormatKey       - [optional] an integer value. Default is Null. The Number Format Key to use for displaying this variable.
;                  $bIsVisible          - [optional] a boolean value. Default is Null. If False, the Set Variable Field is invisible. L.O.'s default is True.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oSetVarField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sValue not a String.
;				   @Error 1 @Extended 4 Return 0 = $iNumFormatKeyKey not an Integer.
;				   @Error 1 @Extended 5 Return 0 = $iNumFormatKeyKey not equal to -1 and Number Format key called in $iNumFormatKeyKey not found in document.
;				   @Error 1 @Extended 6 Return 0 = $bIsVisible not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sValue
;				   |								2 = Error setting $iNumFormatKey
;				   |								4 = Error setting $bIsVisible
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 4 Element Array with values in order of function parameters. The fourth element is the Variable Name.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldSetVarInsert, _LOWriter_FieldsGetList, _LOWriter_FormatKeyCreate _LOWriter_FormatKeyList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldSetVarModify(ByRef $oDoc, ByRef $oSetVarField, $sValue = Null, $iNumFormatKey = Null, $bIsVisible = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avSetVar[4]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oSetVarField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($sValue, $iNumFormatKey, $bIsVisible) Then
		; Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		; fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		; from the value.
		$iNumberFormat = $oSetVarField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : ($iNumberFormat)

		__LOWriter_ArrayFill($avSetVar, $oSetVarField.Content(), $iNumberFormat, $oSetVarField.IsVisible(), $oSetVarField.VariableName())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avSetVar)
	EndIf

	If ($sValue <> Null) Then
		If Not IsString($sValue) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oSetVarField.Content = $sValue
		$iError = ($oSetVarField.Content() = $sValue) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iNumFormatKey <> Null) Then
		If Not IsInt($iNumFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		If ($iNumFormatKey <> -1) And Not _LOWriter_FormatKeyExists($oDoc, $iNumFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oSetVarField.NumberFormat = $iNumFormatKey
		$iError = ($oSetVarField.NumberFormat() = $iNumFormatKey) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bIsVisible <> Null) Then
		If Not IsBool($bIsVisible) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oSetVarField.IsVisible = $bIsVisible
		$iError = ($oSetVarField.IsVisible() = $bIsVisible) ? ($iError) : (BitOR($iError, 4))
	EndIf

	$oSetVarField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldSetVarModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldsGetList
; Description ...: Retrieve an Array of Field Objects contained in a document.
; Syntax ........: _LOWriter_FieldsGetList(ByRef $oDoc[, $iType = $LOW_FIELD_TYPE_ALL[, $bSupportedServices = True[, $bFieldType = True[, $bFieldTypeNum = True]]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $iType               - [optional] an integer value (1-2147483647). Default is $LOW_FIELD_TYPE_ALL. The type of Field to search for. See Constants, $LOW_FIELD_TYPE_* as defined in LibreOfficeWriter_Constants.au3. Can be BitOr'd together.
;                  $bSupportedServices  - [optional] a boolean value. Default is True. If True, adds a column to the array that has the supported service String for that particular Field, To assist in identifying the Field type.
;                  $bFieldType          - [optional] a boolean value. Default is True. If True, adds a column to the array that has the Field Type String for that particular Field as described by Libre Office. To assist in identifying the Field type.
;                  $bFieldTypeNum       - [optional] a boolean value. Default is True. If True, adds a column to the array that has the Field Type Constant Integer for that particular Field, to assist in identifying the Field type. See Constants, $LOW_FIELD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Array
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iType not an Integer, less than 1, or greater than 2147483647. (The total of all Constants added together.) See Constants, $LOW_FIELD_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 3 Return 0 = $bSupportedServices not a Boolean.
;				   @Error 1 @Extended 4 Return 0 = $bFieldType not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bFieldTypeNum not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $avFieldTypes passed to internal function not an Array. UDF needs fixed.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error converting Field type Constants.
;				   @Error 2 @Extended 2 Return 0 = Failed to create enumeration of fields in document.
;				   --Success--
;				   @Error 0 @Extended ? Return Array = Success. Returning Array of Text Field Objects with @Extended set to number of results. See Remarks for Array sizing.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: The Array can vary in the number of columns, if $bSupportedServices, $bFieldType, and $bFieldTypeNum are set
;					to False, the Array will be a single column. With each of the above listed options being set to True, a
;					column will be added in the order they are listed in the UDF parameters. The First column will always be the
;					Field Object.
;					Setting $bSupportedServices to True will add a Supported Service String column for the found Field.
;					Setting $bFieldType to True will add a Field type column for the found Field.
;					Setting $bFieldTypeNum to True will add a Field type Number column, matching the constants, $LOW_FIELD_TYPE_* as defined in LibreOfficeWriter_Constants.au3 for the found Field.
;					Note: For simplicity, and also due to certain Bit limitations I have broken the different Field types into
;						three different categories, Regular Fields, ($LWFieldType), Advanced(Complex) Fields, ($LWFieldAdvType),
;						and Document Information fields (Found in the Document Information Tab in L.O. Fields dialog),
;						($LWFieldDocInfoType). Just because a field is listed in the constants below, does not necessarily mean
;						that I have made a function to create/modify it, you may still be able to update or delete it using the
;						Field Update, or Field Delete function, though. Some Fields are too complex to create a function for,
;						and others are literally impossible.
; Related .......: _LOWriter_FieldsAdvGetList, _LOWriter_FieldsDocInfoGetList, _LOWriter_FieldDelete, _LOWriter_FieldGetAnchor,
;					_LOWriter_FieldUpdate
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldsGetList(ByRef $oDoc, $iType = $LOW_FIELD_TYPE_ALL, $bSupportedServices = True, $bFieldType = True, $bFieldTypeNum = True)
	Local $avFieldTypes[0][0]
	Local $vReturn

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not __LOWriter_IntIsBetween($iType, $LOW_FIELD_TYPE_ALL, 2147483647) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	; 2147483647 is all possible Consts added together
	If Not IsBool($bSupportedServices) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bFieldType) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bFieldTypeNum) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$avFieldTypes = __LOWriter_FieldTypeServices($iType)
	If (@error > 0) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	$vReturn = __LOWriter_FieldsGetList($oDoc, $bSupportedServices, $bFieldType, $bFieldTypeNum, $avFieldTypes)

	Return SetError(@error, @extended, $vReturn)
EndFunc   ;==>_LOWriter_FieldsGetList

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldShowVarInsert
; Description ...: Insert a Show Variable Field.
; Syntax ........: _LOWriter_FieldShowVarInsert(ByRef $oDoc, ByRef $oCursor, $sSetVarName[, $bOverwrite = False[, $iNumFormatKey = Null[, $bShowName = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $sSetVarName         - a string value. The Set Variable name to show the value of.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iNumFormatKey       - [optional] an integer value. Default is Null. The Number Format Key to display the content in.
;                  $bShowName           - [optional] a boolean value. Default is Null. If True, the Set Variable name is displayed rather than its value.
; Return values .: Success: Object.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $sSetVarName not a String.
;				   @Error 1 @Extended 6 Return 0 = Did not find a Set Var Field Master with same name as $sSetVarName.
;				   @Error 1 @Extended 7 Return 0 = $iNumFormatKey not an Integer.
;				   @Error 1 @Extended 8 Return 0 = Number Format key called in $iNumFormatKey not found in document.
;				   @Error 1 @Extended 9 Return 0 = $bShowName not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Error creating "com.sun.star.text.TextField.GetExpression" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object. = Success. Successfully inserted Show Variable field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: This function checks if there is a Set Variable matching the name called in $sSetVarName.
; Related .......: _LOWriter_FieldShowVarModify, _LOWriter_FieldSetVarInsert, _LOWriter_FieldsGetList,
;					_LOWriter_FormatKeyCreate _LOWriter_FormatKeyList, _LOWriter_DocGetViewCursor,
;					_LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor,
;					 _LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor,
;					_LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldShowVarInsert(ByRef $oDoc, ByRef $oCursor, $sSetVarName, $bOverwrite = False, $iNumFormatKey = Null, $bShowName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oShowVarField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oShowVarField = $oDoc.createInstance("com.sun.star.text.TextField.GetExpression")
	If Not IsObj($oShowVarField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If Not IsString($sSetVarName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
	If Not _LOWriter_FieldSetVarMasterExists($oDoc, $sSetVarName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
	$oShowVarField.Content = $sSetVarName

	If ($iNumFormatKey <> Null) Then
		If Not IsInt($iNumFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		If Not _LOWriter_FormatKeyExists($oDoc, $iNumFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 8, 0)
		$oShowVarField.NumberFormat = $iNumFormatKey
	EndIf

	If ($bShowName <> Null) Then
		If Not IsBool($bShowName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 9, 0)
		$oShowVarField.IsShowFormula = $bShowName
		If ($bShowName = True) Then $oShowVarField.NumberFormat = -1
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oShowVarField, $bOverwrite)

	$oShowVarField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oShowVarField)
EndFunc   ;==>_LOWriter_FieldShowVarInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldShowVarModify
; Description ...:Set or Retrieve a Show Variable Field's settings.
; Syntax ........: _LOWriter_FieldShowVarModify(ByRef $oDoc, ByRef $oShowVarField[, $sSetVarName = Null[, $iNumFormatKey = Null[, $bShowName = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oShowVarField       - [in/out] an object. A Show Variable field Object from a previous _LOWriter_FieldShowVarInsert, or _LOWriter_FieldsGetList function.
;                  $sSetVarName         - [optional] a string value. Default is Null. The Set Variable name to show the value of.
;                  $iNumFormatKey       - [optional] an integer value. Default is Null. The Number Format Key to display the content in.
;                  $bShowName           - [optional] a boolean value. Default is Null. If True, the Set Variable name is displayed rather than its value.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oShowVarField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $sSetVarName not a String.
;				   @Error 1 @Extended 4 Return 0 = Did not find a Set Var Field Master with same name as $sSetVarName.
;				   @Error 1 @Extended 5 Return 0 = $iNumFormatKey not an Integer.
;				   @Error 1 @Extended 6 Return 0 = Number Format key called in $iNumFormatKey not found in document.
;				   @Error 1 @Extended 7 Return 0 = $bShowName not a Boolean.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $sSetVarName
;				   |								2 = Error setting $iNumFormatKey
;				   |								4 = Error setting $bShowName
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 3 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;				   Note: This function checks if there is a Set Variable matching the name called in $sSetVarName.
; Related .......: _LOWriter_FieldShowVarInsert, _LOWriter_FieldsGetList, _LOWriter_FormatKeyCreate,  _LOWriter_FormatKeyList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldShowVarModify(ByRef $oDoc, ByRef $oShowVarField, $sSetVarName = Null, $iNumFormatKey = Null, $bShowName = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0, $iNumberFormat
	Local $avShowVar[3]

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oShowVarField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($sSetVarName, $iNumFormatKey, $bShowName) Then
		; Libre Office Seems to insert its Number formats by adding 10,000 to the number, but if I insert that same value, it
		; fails/causes the wrong format to be used, so, If the Number format is greater than or equal to 10,000, Minus 10,000
		; from the value.
		$iNumberFormat = $oShowVarField.NumberFormat()
		$iNumberFormat = ($iNumberFormat >= 10000) ? ($iNumberFormat - 10000) : ($iNumberFormat)

		__LOWriter_ArrayFill($avShowVar, $oShowVarField.Content(), $iNumberFormat, $oShowVarField.IsShowFormula())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avShowVar)
	EndIf

	If ($sSetVarName <> Null) Then
		If Not IsString($sSetVarName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		If Not _LOWriter_FieldSetVarMasterExists($oDoc, $sSetVarName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oShowVarField.Content = $sSetVarName
		$iError = ($oShowVarField.Content() = $sSetVarName) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iNumFormatKey <> Null) Then
		If Not IsInt($iNumFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		If Not _LOWriter_FormatKeyExists($oDoc, $iNumFormatKey) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oShowVarField.NumberFormat = $iNumFormatKey
		$iError = ($oShowVarField.NumberFormat() = ($iNumFormatKey)) ? ($iError) : (BitOR($iError, 2))
	EndIf

	If ($bShowName <> Null) Then
		If Not IsBool($bShowName) Then Return SetError($__LO_STATUS_INPUT_ERROR, 7, 0)
		$oShowVarField.IsShowFormula = $bShowName
		$iError = ($oShowVarField.IsShowFormula() = $bShowName) ? ($iError) : (BitOR($iError, 4))
	EndIf

	$oShowVarField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldShowVarModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldStatCountInsert
; Description ...: Insert a Count Field.
; Syntax ........: _LOWriter_FieldStatCountInsert(ByRef $oDoc, ByRef $oCursor, $iCountType[, $bOverwrite = False[, $iNumFormat = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $iCountType          - an integer value (0-6). The Type of Data to Count. See Constants, $LOW_FIELD_COUNT_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iNumFormat          - [optional] an integer value (0-71). Default is Null. The numbering format to use for Count field numbering. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $iCountType not an integer, less than 0, or greater than 6. See Constants, $LOW_FIELD_COUNT_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 5 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iNumFormat not an integer, less than 0, or greater than 71. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create requested Count Field Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve Field Count Service Type. Check Constants.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Count Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: After insertion there seems to be a necessary delay before the value to display is available, thus when a
;						new count field is inserted, the value will be "0". If you call a _LOWriter_FieldUpdate for this
;						field after a few seconds, the value should appear.
; Related .......: _LOWriter_FieldStatCountModify, _LOWriter_DocGetViewCursor, _LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldStatCountInsert(ByRef $oDoc, ByRef $oCursor, $iCountType, $bOverwrite = False, $iNumFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oCountField
	Local $sFieldType

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not __LOWriter_IntIsBetween($iCountType, $LOW_FIELD_COUNT_TYPE_CHARACTERS, $LOW_FIELD_COUNT_TYPE_WORDS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)

	$sFieldType = __LOWriter_FieldCountType($iCountType)
	If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

	$oCountField = $oDoc.createInstance($sFieldType)
	If Not IsObj($oCountField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oCountField.NumberingType = $iNumFormat
	Else
		$oCountField.NumberingType = $LOW_NUM_STYLE_PAGE_DESCRIPTOR
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oCountField, $bOverwrite)

	$oCountField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oCountField)
EndFunc   ;==>_LOWriter_FieldStatCountInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldStatCountModify
; Description ...: Set or Retrieve a Count Field's settings.
; Syntax ........: _LOWriter_FieldStatCountModify(ByRef $oDoc, ByRef $oCountField[, $iCountType = Null[, $iNumFormat = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCountField         - [in/out] an object. A Count field Object from a previous _LOWriter_FieldStatCountInsert, or _LOWriter_FieldsGetList function.
;                  $iCountType          - [optional] an integer value (0-6). Default is Null. The type of data to count. See Constants, $LOW_FIELD_COUNT_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;                  $iNumFormat          - [optional] an integer value (0-71). Default is Null. The numbering format to use for Count field numbering. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCountField not an Object.
;				   @Error 1 @Extended 3 Return 0 = $iCountType not an integer, less than 0, or greater than 6. See Constants, $LOW_FIELD_COUNT_TYPE_* as defined in LibreOfficeWriter_Constants.au3.
;				   @Error 1 @Extended 4 Return 0 = $iNumFormat not an integer, less than 0, or greater than 71. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create requested Count Field Object.
;				   --Processing Errors--
;				   @Error 3 @Extended 1 Return 0 = Failed to retrieve Field Count Service Type. Check Constants.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iCountType
;				   |								2 = Error setting $iNumFormat
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
;					 After changing the Count type there may be a delay before the value to display is available,
;						thus when the count field is inserted, the value will be "0". If you call a _LOWriter_FieldUpdate for
;						this field after a few seconds, the value should appear.
; Related .......: _LOWriter_FieldStatCountInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldStatCountModify(ByRef $oDoc, ByRef $oCountField, $iCountType = Null, $iNumFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avCountField[2]
	Local $oNewCountField
	Local $sFieldType

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCountField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)

	If __LOWriter_VarsAreNull($iNumFormat) Then
		__LOWriter_ArrayFill($avCountField, __LOWriter_FieldCountType($oCountField), $oCountField.NumberingType())
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)
		Return SetError($__LO_STATUS_SUCCESS, 1, $avCountField)
	EndIf

	If ($iCountType <> Null) Then
		If Not __LOWriter_IntIsBetween($iCountType, $LOW_FIELD_COUNT_TYPE_CHARACTERS, $LOW_FIELD_COUNT_TYPE_WORDS) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$sFieldType = __LOWriter_FieldCountType($iCountType)
		If (@error > 0) Then Return SetError($__LO_STATUS_PROCESSING_ERROR, 1, 0)

		If Not $oCountField.supportsService($sFieldType) Then ; If the Field is already that type, skip this and do nothing.

			$oNewCountField = $oDoc.createInstance($sFieldType)
			If Not IsObj($oNewCountField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

			; It doesn't work to just set a new Count type for an already inserted Count Field, so I have to create a new one and then
			; insert it.
			$oNewCountField.NumberingType = $oCountField.NumberingType()

			$oDoc.Text.createTextCursorByRange($oCountField.Anchor()).Text.insertTextContent($oCountField.Anchor(), $oNewCountField, True)

			; Update the Old Count Field Object to the new one.
			$oCountField = $oNewCountField

			$oCountField.Update()

			$iError = ($oCountField.supportsService($sFieldType)) ? ($iError) : (BitOR($iError, 1))
		EndIf
	EndIf

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)
		$oCountField.NumberingType = $iNumFormat
		$iError = ($oCountField.NumberingType() = $iNumFormat) ? ($iError) : (BitOR($iError, 2))
	EndIf

	$oCountField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldStatCountModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldStatTemplateInsert
; Description ...: Insert a Template Field.
; Syntax ........: _LOWriter_FieldStatTemplateInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $iFormat = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iFormat             - [optional] an integer value (0-5). Default is Null. The Format to display the template data in. See Constants, $LOW_FIELD_FILENAME_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iFormat not an integer, less than 0, or greater than 5. See Constants, $LOW_FIELD_FILENAME_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.TemplateName" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Template Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldStatTemplateModify, _LOWriter_DocGenPropTemplate,  _LOWriter_DocGetViewCursor,
;					_LOWriter_DocCreateTextCursor, _LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor,
;					_LOWriter_DocHeaderGetTextCursor, _LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor,
;					_LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldStatTemplateInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTemplateField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oTemplateField = $oDoc.createInstance("com.sun.star.text.TextField.TemplateName")
	If Not IsObj($oTemplateField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($iFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iFormat, $LOW_FIELD_FILENAME_FULL_PATH, $LOW_FIELD_FILENAME_TEMPLATE_NAME) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oTemplateField.FileFormat = $iFormat
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oTemplateField, $bOverwrite)

	$oTemplateField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oTemplateField)
EndFunc   ;==>_LOWriter_FieldStatTemplateInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldStatTemplateModify
; Description ...: Set or Retrieve a Template Field's settings.
; Syntax ........: _LOWriter_FieldStatTemplateModify(ByRef $oTemplateField[, $iFormat = Null])
; Parameters ....: $oTemplateField      - [in/out] an object. A Template field Object from a previous _LOWriter_FieldStatTemplateInsert, or _LOWriter_FieldsGetList function.
;                  $iFormat             - [optional] an integer value (0-5). Default is Null. The Format to display the template data in. See Constants, $LOW_FIELD_FILENAME_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oTemplateField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iFormat not an integer, less than 0, or greater than 5. See Constants, $LOW_FIELD_FILENAME_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iFormat
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current Template Format Type setting, in Integer format. See File Name Constants.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldStatTemplateInsert, _LOWriter_DocGenPropTemplate, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldStatTemplateModify(ByRef $oTemplateField, $iFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oTemplateField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iFormat) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oTemplateField.FileFormat())

	If Not __LOWriter_IntIsBetween($iFormat, $LOW_FIELD_FILENAME_FULL_PATH, $LOW_FIELD_FILENAME_TEMPLATE_NAME) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$oTemplateField.FileFormat = $iFormat
	$iError = ($oTemplateField.FileFormat() = $iFormat) ? ($iError) : (BitOR($iError, 1))

	$oTemplateField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldStatTemplateModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldUpdate
; Description ...: Update a Field or all fields in a document.
; Syntax ........: _LOWriter_FieldUpdate(ByRef $oDoc[, $oField = Null[, $bForceUpdate = False]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oField              - [optional] an object. Default is Null. A Field Object returned from a previous Insert, _LOWriter_FieldsGetList, _LOWriter_FieldsAdvGetList, or _LOWriter_FieldsDocInfoGetList function. If left as Null, all Fields will be updated.
;                  $bForceUpdate        - [optional] a boolean value. Default is False. If True, Field(s) will be updated whether fixed or not. See remarks.
; Return values .: Success: 1
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object
;				   @Error 1 @Extended 2 Return 0 = $oField not set to Null, and not an Object.
;				   @Error 1 @Extended 3 Return 0 = $bForceUpdate not a Boolean.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to retrieve enumeration of all fields.
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Successfully updated requested field.
;				   @Error 0 @Extended 1 Return 1 = Success. Requested field is set to Fixed and $bForceUpdate is set to false, Field was not updated.
;				   @Error 0 @Extended ? Return 1 = Success. Successfully updated all fields, @Extended set to number of fields updated.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Updating a fixed field will usually erase any user-provided content, such as an author name, creation date
;						etc. If a Field is fixed, the field wont be updated unless $bForceUpdate is set to true.
; Related .......: _LOWriter_FieldsGetList, _LOWriter_FieldsAdvGetList, _LOWriter_FieldsDocInfoGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldUpdate(ByRef $oDoc, $oField = Null, $bForceUpdate = False)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oTextFields, $oTextField
	Local $iCount = 0, $iUpdated = 0

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If $oField <> Null And Not IsObj($oField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If Not IsBool($bForceUpdate) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)

	If ($oField <> Null) Then
		If ($oField.getPropertySetInfo.hasPropertyByName("IsFixed") = True) Then
			If ($oField.IsFixed() = True) And ($bForceUpdate = False) Then Return SetError($__LO_STATUS_SUCCESS, 1, 1) ; Updating a fixed field, causes its content to be removed.
		EndIf
		$oField.Update()
		Return SetError($__LO_STATUS_SUCCESS, 0, 1)
	EndIf

	; Update All Fields.
	$oTextFields = $oDoc.getTextFields.createEnumeration()
	If Not IsObj($oTextFields) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	While $oTextFields.hasMoreElements()
		$oTextField = $oTextFields.nextElement()

		If ($bForceUpdate = False) Then
			If ($oTextField.getPropertySetInfo.hasPropertyByName("IsFixed") = True) Then
				If ($oTextField.IsFixed() = False) Then
					$oTextField.Update()
					$iUpdated += 1
				EndIf ;Updating a fixed field, causes its content to be removed.
			Else
				$oTextField.Update()
				$iUpdated += 1
			EndIf

		Else
			$oTextField.Update()
			$iUpdated += 1
		EndIf

		$iCount += 1
		Sleep((IsInt($iCount / $__LOWCONST_SLEEP_DIV) ? (10) : (0)))
	WEnd

	Return SetError($__LO_STATUS_SUCCESS, $iUpdated, 1)
EndFunc   ;==>_LOWriter_FieldUpdate

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldVarSetPageInsert
; Description ...: Insert a Set Page Variable Field.
; Syntax ........: _LOWriter_FieldVarSetPageInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $bRefOn = Null[, $iOffset = Null]]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $bRefOn              - [optional] a boolean value. Default is Null. If True, Reference point is enabled, else disabled.
;                  $iOffset             - [optional] an integer value. Default is Null. The offset to start the page count from.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $bRefOn not a Boolean.
;				   @Error 1 @Extended 6 Return 0 = $iOffset not an Integer.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.ReferencePageSet" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Set Page Variable Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldVarSetPageModify, _LOWriter_DocGetViewCursor,	_LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldVarSetPageInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $bRefOn = Null, $iOffset = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPageVarSetField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oPageVarSetField = $oDoc.createInstance("com.sun.star.text.TextField.ReferencePageSet")
	If Not IsObj($oPageVarSetField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($bRefOn <> Null) Then
		If Not IsBool($bRefOn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oPageVarSetField.On = $bRefOn
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LO_STATUS_INPUT_ERROR, 6, 0)
		$oPageVarSetField.Offset = $iOffset
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oPageVarSetField, $bOverwrite)

	$oPageVarSetField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPageVarSetField)
EndFunc   ;==>_LOWriter_FieldVarSetPageInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldVarSetPageModify
; Description ...: Set or retrieve a Set Page Variable Field's settings.
; Syntax ........: _LOWriter_FieldVarSetPageModify(ByRef $oPageVarSetField[, $bRefOn = Null[, $iOffset = Null]]])
; Parameters ....: $oPageVarSetField    - [in/out] an object. A Set Page Variable field Object from a previous _LOWriter_FieldVarSetPageInsert, or _LOWriter_FieldsGetList function.
;                  $bRefOn              - [optional] a boolean value. Default is Null. If True, Reference point is enabled, else disabled.
;                  $iOffset             - [optional] an integer value. Default is Null. The offset to start the page count from.
; Return values .: Success: 1 or Array.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oPageVarSetField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $bRefOn not a Boolean.
;				   @Error 1 @Extended 3 Return 0 = $iOffset not an Integer.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $bRefOn
;				   |								2 = Error setting $iOffset
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Array = Success. All optional parameters were set to Null, returning current settings in a 2 Element Array with values in order of function parameters.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldVarSetPageInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldVarSetPageModify(ByRef $oPageVarSetField, $bRefOn = Null, $iOffset = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0
	Local $avPage[2]

	If Not IsObj($oPageVarSetField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($bRefOn, $iOffset) Then
		__LOWriter_ArrayFill($avPage, $oPageVarSetField.On(), $oPageVarSetField.Offset())
		Return SetError($__LO_STATUS_SUCCESS, 1, $avPage)
	EndIf

	If ($bRefOn <> Null) Then
		If Not IsBool($bRefOn) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
		$oPageVarSetField.On = $bRefOn
		$iError = ($oPageVarSetField.On() = $bRefOn) ? ($iError) : (BitOR($iError, 1))
	EndIf

	If ($iOffset <> Null) Then
		If Not IsInt($iOffset) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
		$oPageVarSetField.Offset = $iOffset
		$iError = ($oPageVarSetField.Offset() = $iOffset) ? ($iError) : (BitOR($iError, 2))
	EndIf

	$oPageVarSetField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldVarSetPageModify

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldVarShowPageInsert
; Description ...: Insert a Show Page Variable Field.
; Syntax ........: _LOWriter_FieldVarShowPageInsert(ByRef $oDoc, ByRef $oCursor[, $bOverwrite = False[, $iNumFormat = Null]])
; Parameters ....: $oDoc                - [in/out] an object. A Document object returned by previous _LOWriter_DocOpen, _LOWriter_DocConnect, or _LOWriter_DocCreate function.
;                  $oCursor             - [in/out] an object. A Cursor Object returned from any Cursor Object creation Or retrieval function. Cannot be a Table Cursor.
;                  $bOverwrite          - [optional] a boolean value. Default is False. If True, any content selected by the cursor will be overwritten. If False, content will be inserted to the left of any selection.
;                  $iNumFormat          - [optional] an integer value (0-71). Default is Null. The numbering format to use for Show Page Variable numbering. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: Object
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oDoc not an Object.
;				   @Error 1 @Extended 2 Return 0 = $oCursor not an Object.
;				   @Error 1 @Extended 3 Return 0 = $oCursor is a Table Cursor, and is not supported.
;				   @Error 1 @Extended 4 Return 0 = $bOverwrite not a Boolean.
;				   @Error 1 @Extended 5 Return 0 = $iNumFormat not an integer, less than 0, or greater than 71. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Initialization Errors--
;				   @Error 2 @Extended 1 Return 0 = Failed to create "com.sun.star.text.TextField.ReferencePageGet" Object.
;				   --Success--
;				   @Error 0 @Extended 0 Return Object = Success. Successfully Inserted a Show Page Variable Field, returning its Object.
; Author ........: donnyh13
; Modified ......:
; Remarks .......:
; Related .......: _LOWriter_FieldVarShowPageModify, _LOWriter_DocGetViewCursor,	_LOWriter_DocCreateTextCursor,
;					_LOWriter_CellCreateTextCursor, _LOWriter_FrameCreateTextCursor, _LOWriter_DocHeaderGetTextCursor,
;					_LOWriter_DocFooterGetTextCursor, _LOWriter_EndnoteGetTextCursor, _LOWriter_FootnoteGetTextCursor
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldVarShowPageInsert(ByRef $oDoc, ByRef $oCursor, $bOverwrite = False, $iNumFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $oPageShowField

	If Not IsObj($oDoc) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)
	If Not IsObj($oCursor) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	If (__LOWriter_Internal_CursorGetType($oCursor) = $LOW_CURTYPE_TABLE_CURSOR) Then Return SetError($__LO_STATUS_INPUT_ERROR, 3, 0)
	If Not IsBool($bOverwrite) Then Return SetError($__LO_STATUS_INPUT_ERROR, 4, 0)

	$oPageShowField = $oDoc.createInstance("com.sun.star.text.TextField.ReferencePageGet")
	If Not IsObj($oPageShowField) Then Return SetError($__LO_STATUS_INIT_ERROR, 1, 0)

	If ($iNumFormat <> Null) Then
		If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 5, 0)
		$oPageShowField.NumberingType = $iNumFormat
	Else
		$oPageShowField.NumberingType = $LOW_NUM_STYLE_PAGE_DESCRIPTOR
	EndIf

	$oCursor.Text.insertTextContent($oCursor, $oPageShowField, $bOverwrite)

	$oPageShowField.Update()

	Return SetError($__LO_STATUS_SUCCESS, 0, $oPageShowField)
EndFunc   ;==>_LOWriter_FieldVarShowPageInsert

; #FUNCTION# ====================================================================================================================
; Name ..........: _LOWriter_FieldVarShowPageModify
; Description ...: Set or Retrieve a Show Page Variable Field's settings.
; Syntax ........: _LOWriter_FieldVarShowPageModify(ByRef $oPageShowField[, $iNumFormat = Null])
; Parameters ....: $oPageShowField        - [in/out] an object. A Show Page Variable field Object from a previous _LOWriter_FieldVarShowPageInsert, or _LOWriter_FieldsGetList function.
;                  $iNumFormat            - [optional] an integer value (0-71). Default is Null. The numbering format to use for Show Page Variable numbering. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
; Return values .: Success: 1 or Integer.
;				   Failure: 0 and sets the @Error and @Extended flags to non-zero.
;				   --Input Errors--
;				   @Error 1 @Extended 1 Return 0 = $oPageShowField not an Object.
;				   @Error 1 @Extended 2 Return 0 = $iNumFormat not an integer, less than 0, or greater than 71. See Constants, $LOW_NUM_STYLE_* as defined in LibreOfficeWriter_Constants.au3.
;				   --Property Setting Errors--
;				   @Error 4 @Extended ? Return 0 = Some settings were not successfully set. Use BitAND to test @Extended for the following values:
;				   |								1 = Error setting $iNumFormat
;				   --Success--
;				   @Error 0 @Extended 0 Return 1 = Success. Settings were successfully set.
;				   @Error 0 @Extended 1 Return Integer = Success. All optional parameters were set to Null, returning current numbering type setting, in Integer format.
; Author ........: donnyh13
; Modified ......:
; Remarks .......: Call this function with only the required parameters (or with all other parameters set to Null keyword), to get the current settings.
;				   Call any optional parameter with Null keyword to skip it.
; Related .......: _LOWriter_FieldVarShowPageInsert, _LOWriter_FieldsGetList
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _LOWriter_FieldVarShowPageModify(ByRef $oPageShowField, $iNumFormat = Null)
	Local $oCOM_ErrorHandler = ObjEvent("AutoIt.Error", __LOWriter_InternalComErrorHandler)
	#forceref $oCOM_ErrorHandler

	Local $iError = 0

	If Not IsObj($oPageShowField) Then Return SetError($__LO_STATUS_INPUT_ERROR, 1, 0)

	If __LOWriter_VarsAreNull($iNumFormat) Then Return SetError($__LO_STATUS_SUCCESS, 1, $oPageShowField.NumberingType())

	If Not __LOWriter_IntIsBetween($iNumFormat, $LOW_NUM_STYLE_CHARS_UPPER_LETTER, $LOW_NUM_STYLE_NUMBER_LEGAL_KO) Then Return SetError($__LO_STATUS_INPUT_ERROR, 2, 0)
	$oPageShowField.NumberingType = $iNumFormat
	$iError = ($oPageShowField.NumberingType() = $iNumFormat) ? ($iError) : (BitOR($iError, 1))

	$oPageShowField.Update()

	Return ($iError > 0) ? (SetError($__LO_STATUS_PROP_SETTING_ERROR, $iError, 0)) : (SetError($__LO_STATUS_SUCCESS, 0, 1))
EndFunc   ;==>_LOWriter_FieldVarShowPageModify
